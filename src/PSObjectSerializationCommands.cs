// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

#nullable disable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Xml;
using Microsoft.PowerShell.Commands;

using Dbg = System.Management.Automation.Diagnostics;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Serialization format for unified PSObject serialization cmdlets.
    /// </summary>
    public enum PSSerializationFormat
    {
        /// <summary>
        /// CBOR (RFC 8949) PSObject serialization format (Self-Described CBOR tag 55799).
        /// </summary>
        Cbor = 0,

        /// <summary>
        /// Protocol Buffers (proto3) PSObject serialization format ("PSPB" header).
        /// </summary>
        Protobuf = 1,

        /// <summary>
        /// PowerShell CLIXML serialization format.
        /// </summary>
        Clixml = 2
    }

    /// <summary>
    /// Transforms friendly format aliases (e.g. "Proto", "CliProto", "CliCbor", "CliXml", "Xml") into <see cref="PSSerializationFormat"/>.
    /// </summary>
    internal sealed class PSSerializationFormatTransformationAttribute : ArgumentTransformationAttribute
    {
        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            object raw = inputData is PSObject pso ? pso.BaseObject : inputData;
            switch (raw)
            {
                case PSSerializationFormat format:
                    return format;

                case PSHostGrpcEncoding grpcEncoding:
                    return grpcEncoding switch
                    {
                        PSHostGrpcEncoding.Protobuf => PSSerializationFormat.Protobuf,
                        PSHostGrpcEncoding.Clixml => PSSerializationFormat.Clixml,
                        _ => PSSerializationFormat.Cbor
                    };

                case string text:
                    return text.Trim().ToLowerInvariant() switch
                    {
                        "protobuf" or "proto" or "cliproto" or "pb" => PSSerializationFormat.Protobuf,
                        "cbor" or "clicbor" => PSSerializationFormat.Cbor,
                        "clixml" or "xml" => PSSerializationFormat.Clixml,
                        _ => inputData
                    };

                default:
                    return inputData;
            }
        }
    }

    /// <summary>
    /// Format detection utilities for unified PSObject serialization cmdlets.
    /// </summary>
    internal static class PSSerializationFormatDetector
    {
        internal static PSSerializationFormat InferFromPath(string path, PSSerializationFormat fallback = PSSerializationFormat.Cbor)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return fallback;
            }

            string ext = System.IO.Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext))
            {
                return fallback;
            }

            return ext.ToLowerInvariant() switch
            {
                ".cbor" => PSSerializationFormat.Cbor,
                ".xml" or ".clixml" => PSSerializationFormat.Clixml,
                ".pb" or ".proto" or ".cliproto" => PSSerializationFormat.Protobuf,
                _ => fallback
            };
        }

        internal static PSSerializationFormat DetectFromBytes(ReadOnlySpan<byte> bytes, PSSerializationFormat fallback = PSSerializationFormat.Cbor)
        {
            if (bytes.Length >= 4 &&
                bytes[0] == (byte)'P' &&
                bytes[1] == (byte)'S' &&
                bytes[2] == (byte)'P' &&
                bytes[3] == (byte)'B')
            {
                return PSSerializationFormat.Protobuf;
            }

            if (bytes.Length >= 3 &&
                bytes[0] == 0xD9 &&
                bytes[1] == 0xD9 &&
                bytes[2] == 0xF7)
            {
                return PSSerializationFormat.Cbor;
            }

            if (LooksLikeXmlBytes(bytes))
            {
                return PSSerializationFormat.Clixml;
            }

            return fallback;
        }

        internal static PSSerializationFormat DetectFromStream(Stream stream, string pathHint = null)
        {
            PSSerializationFormat fallback = InferFromPath(pathHint, PSSerializationFormat.Cbor);
            if (stream is null || !stream.CanSeek)
            {
                return fallback;
            }

            long originalPosition = stream.Position;
            Span<byte> header = stackalloc byte[16];
            int read = stream.Read(header);
            stream.Position = originalPosition;

            if (read == 0)
            {
                return fallback;
            }

            return DetectFromBytes(header.Slice(0, read), fallback);
        }

        private static bool LooksLikeXmlBytes(ReadOnlySpan<byte> bytes)
        {
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                bytes = bytes.Slice(3);
            }
            else if (bytes.Length >= 2 &&
                ((bytes[0] == 0xFF && bytes[1] == 0xFE) || (bytes[0] == 0xFE && bytes[1] == 0xFF)))
            {
                return true;
            }

            int idx = 0;
            while (idx < bytes.Length &&
                (bytes[idx] == (byte)' ' || bytes[idx] == (byte)'\t' || bytes[idx] == (byte)'\r' || bytes[idx] == (byte)'\n'))
            {
                idx++;
            }

            return idx < bytes.Length && (bytes[idx] == (byte)'<' || bytes[idx] == (byte)'#');
        }
    }

    #region ExportPSObjectCommand

    /// <summary>
    /// Implementation for the unified Export-PSObject command.
    /// </summary>
    [Cmdlet(VerbsData.Export, "PSObject", SupportsShouldProcess = true, DefaultParameterSetName = "ByPath")]
    public sealed class ExportPSObjectCommand : PSCmdlet, IDisposable
    {
        /// <summary>
        /// Gets or sets the serialization format (Cbor, Protobuf, or Clixml). Defaults to Cbor.
        /// </summary>
        [Parameter]
        [PSSerializationFormatTransformation]
        public PSSerializationFormat Format { get; set; } = PSSerializationFormat.Cbor;

        /// <summary>
        /// Specifies the depth of serialization.
        /// </summary>
        [Parameter]
        [ValidateRange(1, int.MaxValue)]
        public int Depth { get; set; } = 2;

        /// <summary>
        /// Gets or sets the path of the file to export to.
        /// </summary>
        [Parameter(Mandatory = true, Position = 0, ParameterSetName = "ByPath", ValueFromPipelineByPropertyName = true)]
        public string Path { get; set; }

        /// <summary>
        /// Gets or sets the literal path of the file to export to.
        /// </summary>
        [Parameter(Mandatory = true, ParameterSetName = "ByLiteralPath", ValueFromPipelineByPropertyName = true)]
        [Alias("PSPath", "LP")]
        public string LiteralPath
        {
            get => Path;
            set
            {
                Path = value;
                _isLiteralPath = true;
            }
        }

        private bool _isLiteralPath;

        /// <summary>
        /// Gets or sets the input object to be exported.
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipeline = true)]
        [AllowNull]
        public PSObject InputObject { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether read-only files should be overwritten.
        /// </summary>
        [Parameter]
        public SwitchParameter Force { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether existing files should not be overwritten.
        /// </summary>
        [Parameter]
        [Alias("NoOverwrite")]
        public SwitchParameter NoClobber { get; set; }

        /// <summary>
        /// Gets or sets the character encoding used when exporting in Clixml format. Defaults to utf8NoBOM.
        /// </summary>
        [Parameter]
        [EncodingTransformation]
        public Encoding Encoding { get; set; } = EncodingTransformationAttribute.Utf8NoBom;

        private FileStream _fileStream;
        private StreamWriter _streamWriter;
        private XmlWriter _xmlWriter;
        private ProtoSerializer _protoSerializer;
        private CborSerializer _cborSerializer;
        private CliXmlSerializer _cliXmlSerializer;
        private FileInfo _readOnlyFileInfo;
        private PSSerializationFormat _effectiveFormat;

        /// <summary>
        /// Opens the destination file and initializes the selected serializer.
        /// </summary>
        protected override void BeginProcessing()
        {
            Dbg.Assert(Path is not null, "Path is mandatory parameter");

            if (!ShouldProcess(Path))
            {
                return;
            }

            _effectiveFormat = MyInvocation.BoundParameters.ContainsKey(nameof(Format))
                ? Format
                : PSSerializationFormatDetector.InferFromPath(Path, PSSerializationFormat.Cbor);

            PathUtils.MasterStreamOpen(
                this,
                Path,
                EncodingConversion.Unicode,
                defaultEncoding: true,
                Append: false,
                Force: Force,
                NoClobber: NoClobber,
                fileStream: out _fileStream,
                streamWriter: out _,
                readOnlyFileInfo: out _readOnlyFileInfo,
                isLiteralPath: _isLiteralPath);

            switch (_effectiveFormat)
            {
                case PSSerializationFormat.Protobuf:
                    _protoSerializer = new ProtoSerializer(_fileStream, Depth, useDepthFromTypes: true);
                    break;

                case PSSerializationFormat.Cbor:
                    _cborSerializer = new CborSerializer(_fileStream, Depth, useDepthFromTypes: true);
                    break;

                case PSSerializationFormat.Clixml:
                    _streamWriter = new StreamWriter(_fileStream, Encoding ?? EncodingTransformationAttribute.Utf8NoBom);
                    XmlWriterSettings xmlSettings = new()
                    {
                        CloseOutput = true,
                        Encoding = _streamWriter.Encoding,
                        Indent = true,
                        OmitXmlDeclaration = true
                    };
                    _xmlWriter = XmlWriter.Create(_streamWriter, xmlSettings);
                    _cliXmlSerializer = new CliXmlSerializer(_xmlWriter, Depth, useDepthFromTypes: true);
                    break;
            }
        }

        /// <summary>
        /// Serializes each input object into the active stream.
        /// </summary>
        protected override void ProcessRecord()
        {
            if (_protoSerializer is not null)
            {
                _protoSerializer.Serialize(InputObject);
                _fileStream.Flush();
            }
            else if (_cborSerializer is not null)
            {
                _cborSerializer.Serialize(InputObject);
                _fileStream.Flush();
            }
            else if (_cliXmlSerializer is not null)
            {
                _cliXmlSerializer.Serialize(InputObject);
                _xmlWriter.Flush();
            }
        }

        /// <summary>
        /// Finalizes the stream and closes the file.
        /// </summary>
        protected override void EndProcessing()
        {
            if (_protoSerializer is not null)
            {
                _protoSerializer.Done();
                _protoSerializer = null;
            }

            if (_cborSerializer is not null)
            {
                _cborSerializer.Done();
                _cborSerializer = null;
            }

            if (_cliXmlSerializer is not null)
            {
                _cliXmlSerializer.Done();
                _cliXmlSerializer = null;
            }

            CleanUp();
        }

        /// <summary>
        /// Stops an in-progress serialization.
        /// </summary>
        protected override void StopProcessing()
        {
            _protoSerializer?.Stop();
            _cborSerializer?.Stop();
            _cliXmlSerializer?.Stop();
        }

        private void CleanUp()
        {
            if (_xmlWriter is not null)
            {
                _xmlWriter.Dispose();
                _xmlWriter = null;
            }

            if (_streamWriter is not null)
            {
                _streamWriter.Dispose();
                _streamWriter = null;
            }

            if (_fileStream is not null)
            {
                _fileStream.Dispose();
                _fileStream = null;
            }

            if (_readOnlyFileInfo is not null)
            {
                _readOnlyFileInfo.Attributes |= FileAttributes.ReadOnly;
                _readOnlyFileInfo = null;
            }
        }

        private bool _disposed;

        /// <summary>
        /// Releases all resources used by the cmdlet.
        /// </summary>
        public void Dispose()
        {
            if (!_disposed)
            {
                CleanUp();
            }

            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }

    #endregion ExportPSObjectCommand

    #region ImportPSObjectCommand

    /// <summary>
    /// Implementation for the unified Import-PSObject command.
    /// </summary>
    [Cmdlet(VerbsData.Import, "PSObject", DefaultParameterSetName = "ByPath", SupportsPaging = true)]
    public sealed class ImportPSObjectCommand : PSCmdlet, IDisposable
    {
        /// <summary>
        /// Gets or sets the file paths to import from.
        /// </summary>
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "ByPath")]
        public string[] Path { get; set; }

        /// <summary>
        /// Gets or sets the literal file paths to import from.
        /// </summary>
        [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = "ByLiteralPath")]
        [Alias("PSPath", "LP")]
        public string[] LiteralPath
        {
            get => Path;
            set
            {
                Path = value;
                _isLiteralPath = true;
            }
        }

        private bool _isLiteralPath;

        /// <summary>
        /// Gets or sets the serialization format (Protobuf, Cbor, or Clixml). If omitted, auto-detected from file header/extension.
        /// </summary>
        [Parameter]
        [PSSerializationFormatTransformation]
        public PSSerializationFormat Format { get; set; }

        private ImportProtoHelper _protoHelper;
        private ImportCborHelper _cborHelper;
        private ImportCliXmlHelper _cliXmlHelper;

        /// <summary>
        /// Imports serialized PSObject files specified by Path or LiteralPath.
        /// </summary>
        protected override void ProcessRecord()
        {
            if (Path is null)
            {
                return;
            }

            bool explicitFormat = MyInvocation.BoundParameters.ContainsKey(nameof(Format));
            foreach (string path in Path)
            {
                FileStream fs = null;
                try
                {
                    fs = PathUtils.OpenFileStream(path, this, _isLiteralPath);
                    PSSerializationFormat effectiveFormat = explicitFormat
                        ? Format
                        : PSSerializationFormatDetector.DetectFromStream(fs, path);

                    switch (effectiveFormat)
                    {
                        case PSSerializationFormat.Protobuf:
                            _protoHelper = new ImportProtoHelper(this, path, _isLiteralPath);
                            _protoHelper.ImportFromOpenStream(fs);
                            break;

                        case PSSerializationFormat.Cbor:
                            _cborHelper = new ImportCborHelper(this, path, _isLiteralPath);
                            _cborHelper.ImportFromOpenStream(fs);
                            break;

                        case PSSerializationFormat.Clixml:
                            _cliXmlHelper = new ImportCliXmlHelper(this, path, _isLiteralPath);
                            _cliXmlHelper.ImportFromOpenStream(fs);
                            break;
                    }
                }
                catch (Exception e) when (e is IOException or ArgumentException or FormatException or UnauthorizedAccessException or XmlException)
                {
                    WriteError(new ErrorRecord(e, "ImportPSObjectError", ErrorCategory.InvalidData, path));
                }
                finally
                {
                    fs?.Dispose();
                }
            }
        }

        /// <summary>
        /// Stops an in-progress deserialization.
        /// </summary>
        protected override void StopProcessing()
        {
            _protoHelper?.Stop();
            _cborHelper?.Stop();
            _cliXmlHelper?.Stop();
        }

        /// <summary>
        /// Releases all resources used by the cmdlet.
        /// </summary>
        public void Dispose()
        {
            _protoHelper?.Dispose();
            _cborHelper?.Dispose();
            _cliXmlHelper?.Dispose();
            GC.SuppressFinalize(this);
        }
    }

    #endregion ImportPSObjectCommand

    #region ConvertToPSObjectCommand

    /// <summary>
    /// Implementation for the unified ConvertTo-PSObject command.
    /// </summary>
    [Cmdlet(VerbsData.ConvertTo, "PSObject")]
    [OutputType(typeof(byte[]), typeof(string))]
    public sealed class ConvertToPSObjectCommand : PSCmdlet, IDisposable
    {
        /// <summary>
        /// Gets or sets the serialization format (Cbor, Protobuf, or Clixml). Defaults to Cbor.
        /// </summary>
        [Parameter]
        [PSSerializationFormatTransformation]
        public PSSerializationFormat Format { get; set; } = PSSerializationFormat.Cbor;

        /// <summary>
        /// Gets or sets the depth of serialization.
        /// </summary>
        [Parameter]
        [ValidateRange(1, int.MaxValue)]
        public int Depth { get; set; } = 2;

        /// <summary>
        /// Gets or sets the input object to serialize.
        /// </summary>
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        [AllowNull]
        public PSObject InputObject { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to return the serialized payload as a Base64-encoded string.
        /// </summary>
        [Parameter]
        public SwitchParameter AsBase64String { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to return the serialized payload as a byte array (even when Format is Clixml).
        /// </summary>
        [Parameter]
        public SwitchParameter AsByteArray { get; set; }

        private MemoryStream _reusableMemoryStream;
        private StringBuilder _reusableStringBuilder;
        private ProtoSerializer _activeProtoSerializer;
        private CborSerializer _activeCborSerializer;
        private CliXmlSerializer _activeCliXmlSerializer;

        /// <summary>
        /// Serializes each input object and writes its payload immediately to the pipeline.
        /// </summary>
        protected override void ProcessRecord()
        {
            switch (Format)
            {
                case PSSerializationFormat.Protobuf:
                    {
                        _reusableMemoryStream ??= new MemoryStream(256);
                        _reusableMemoryStream.Position = 0;
                        _reusableMemoryStream.SetLength(0);
                        _activeProtoSerializer = new ProtoSerializer(_reusableMemoryStream, Depth, useDepthFromTypes: true);
                        try
                        {
                            _activeProtoSerializer.Serialize(InputObject);
                            _activeProtoSerializer.Done();
                        }
                        finally
                        {
                            _activeProtoSerializer = null;
                        }

                        byte[] bytes = _reusableMemoryStream.ToArray();
                        if (AsBase64String.IsPresent)
                        {
                            WriteObject(Convert.ToBase64String(bytes));
                        }
                        else
                        {
                            WriteObject(bytes, enumerateCollection: false);
                        }

                        break;
                    }

                case PSSerializationFormat.Cbor:
                    {
                        _reusableMemoryStream ??= new MemoryStream(256);
                        _reusableMemoryStream.Position = 0;
                        _reusableMemoryStream.SetLength(0);
                        _activeCborSerializer = new CborSerializer(_reusableMemoryStream, Depth, useDepthFromTypes: true);
                        try
                        {
                            _activeCborSerializer.Serialize(InputObject);
                            _activeCborSerializer.Done();
                        }
                        finally
                        {
                            _activeCborSerializer = null;
                        }

                        byte[] bytes = _reusableMemoryStream.ToArray();
                        if (AsBase64String.IsPresent)
                        {
                            WriteObject(Convert.ToBase64String(bytes));
                        }
                        else
                        {
                            WriteObject(bytes, enumerateCollection: false);
                        }

                        break;
                    }

                case PSSerializationFormat.Clixml:
                    {
                        _reusableStringBuilder ??= new StringBuilder(256);
                        _reusableStringBuilder.Clear();
                        XmlWriterSettings xmlSettings = new()
                        {
                            CloseOutput = false,
                            Encoding = Encoding.Unicode,
                            Indent = true,
                            OmitXmlDeclaration = true
                        };
                        using (XmlWriter xmlWriter = XmlWriter.Create(_reusableStringBuilder, xmlSettings))
                        {
                            _activeCliXmlSerializer = new CliXmlSerializer(xmlWriter, Depth, useDepthFromTypes: true);
                            try
                            {
                                _activeCliXmlSerializer.Serialize(InputObject);
                                _activeCliXmlSerializer.Done();
                            }
                            finally
                            {
                                _activeCliXmlSerializer = null;
                            }
                        }

                        string xml = _reusableStringBuilder.ToString();
                        if (AsByteArray.IsPresent)
                        {
                            byte[] bytes = EncodingTransformationAttribute.Utf8NoBom.GetBytes(xml);
                            WriteObject(bytes, enumerateCollection: false);
                        }
                        else if (AsBase64String.IsPresent)
                        {
                            byte[] bytes = EncodingTransformationAttribute.Utf8NoBom.GetBytes(xml);
                            WriteObject(Convert.ToBase64String(bytes));
                        }
                        else
                        {
                            WriteObject(xml);
                        }

                        break;
                    }
            }
        }

        /// <summary>
        /// Stops an in-progress serialization.
        /// </summary>
        protected override void StopProcessing()
        {
            _activeProtoSerializer?.Stop();
            _activeCborSerializer?.Stop();
            _activeCliXmlSerializer?.Stop();
        }

        /// <summary>
        /// Releases all resources used by the cmdlet.
        /// </summary>
        public void Dispose()
        {
            _reusableMemoryStream?.Dispose();
            _reusableMemoryStream = null;
            _reusableStringBuilder = null;
            GC.SuppressFinalize(this);
        }
    }

    #endregion ConvertToPSObjectCommand

    #region ConvertFromPSObjectCommand

    /// <summary>
    /// Implementation for the unified ConvertFrom-PSObject command.
    /// </summary>
    [Cmdlet(VerbsData.ConvertFrom, "PSObject")]
    public sealed class ConvertFromPSObjectCommand : PSCmdlet, IDisposable
    {
        /// <summary>
        /// Gets or sets the serialized payload (byte[], pipeline bytes, Stream, CLIXML string, or Base64 string) to deserialize.
        /// </summary>
        [Parameter(Position = 0, Mandatory = true, ValueFromPipeline = true)]
        [AllowNull]
        public object InputObject { get; set; }

        /// <summary>
        /// Gets or sets the serialization format (Protobuf, Cbor, or Clixml). If omitted, auto-detected from payload header.
        /// </summary>
        [Parameter]
        [PSSerializationFormatTransformation]
        public PSSerializationFormat Format { get; set; }

        private ImportProtoHelper _protoHelper;
        private ImportCborHelper _cborHelper;
        private ImportCliXmlHelper _cliXmlHelper;
        private MemoryStream _pipedBytesBuffer;

        /// <summary>
        /// Processes serialized input and writes deserialized objects to the pipeline.
        /// </summary>
        protected override void ProcessRecord()
        {
            if (InputObject is null)
            {
                return;
            }

            ProcessInputValue(InputObject);
        }

        private void ProcessInputValue(object input)
        {
            if (input is null)
            {
                return;
            }

            object raw = input is PSObject pso ? pso.BaseObject : input;
            switch (raw)
            {
                case byte[] bytes:
                    FlushBufferedBytesIfAny();
                    DeserializeFromBytes(bytes);
                    break;

                case byte singleByte:
                    _pipedBytesBuffer ??= new MemoryStream();
                    _pipedBytesBuffer.WriteByte(singleByte);
                    break;

                case IList<byte> byteList:
                    {
                        FlushBufferedBytesIfAny();
                        byte[] arr = new byte[byteList.Count];
                        byteList.CopyTo(arr, 0);
                        DeserializeFromBytes(arr);
                        break;
                    }

                case Stream stream:
                    FlushBufferedBytesIfAny();
                    DeserializeFromStream(stream);
                    break;

                case string text:
                    FlushBufferedBytesIfAny();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        DeserializeFromString(text);
                    }
                    break;

                case System.Collections.IEnumerable enumerable:
                    foreach (object item in enumerable)
                    {
                        ProcessInputValue(item);
                    }
                    FlushBufferedBytesIfAny();
                    break;

                default:
                    ThrowTerminatingError(
                        new ErrorRecord(
                            new ArgumentException($"Cannot convert input of type '{raw.GetType().FullName}'. Expected byte[], byte, Stream, CLIXML string, or Base64 string."),
                            "InvalidPSObjectInputType",
                            ErrorCategory.InvalidArgument,
                            input));
                    break;
            }
        }

        /// <summary>
        /// Deserializes any individual bytes accumulated from pipeline enumeration.
        /// </summary>
        protected override void EndProcessing()
        {
            FlushBufferedBytesIfAny();
        }

        /// <summary>
        /// Stops an in-progress deserialization.
        /// </summary>
        protected override void StopProcessing()
        {
            _protoHelper?.Stop();
            _cborHelper?.Stop();
            _cliXmlHelper?.Stop();
        }

        private void FlushBufferedBytesIfAny()
        {
            if (_pipedBytesBuffer is not null && _pipedBytesBuffer.Length > 0)
            {
                byte[] bytes = _pipedBytesBuffer.ToArray();
                _pipedBytesBuffer.SetLength(0);
                DeserializeFromBytes(bytes);
            }
        }

        private void DeserializeFromString(string text)
        {
            bool explicitFormat = MyInvocation.BoundParameters.ContainsKey(nameof(Format));

            ReadOnlySpan<char> trimmed = text.AsSpan().TrimStart();
            if (!trimmed.IsEmpty && trimmed[0] == '\uFEFF')
            {
                trimmed = trimmed.Slice(1).TrimStart();
            }

            bool looksLikeXmlText = !trimmed.IsEmpty && (trimmed[0] == '<' || trimmed[0] == '#');
            if (looksLikeXmlText && (!explicitFormat || Format == PSSerializationFormat.Clixml))
            {
                _cliXmlHelper = new ImportCliXmlHelper(this);
                _cliXmlHelper.ImportStringContent(text);
                return;
            }

            if (explicitFormat && Format == PSSerializationFormat.Clixml && !looksLikeXmlText)
            {
                try
                {
                    byte[] decoded = Convert.FromBase64String(text);
                    DeserializeFromBytes(decoded);
                    return;
                }
                catch (FormatException)
                {
                    _cliXmlHelper = new ImportCliXmlHelper(this);
                    _cliXmlHelper.ImportStringContent(text);
                    return;
                }
            }

            byte[] bytes = Convert.FromBase64String(text);
            DeserializeFromBytes(bytes);
        }

        private void DeserializeFromBytes(byte[] bytes)
        {
            bool explicitFormat = MyInvocation.BoundParameters.ContainsKey(nameof(Format));
            PSSerializationFormat effectiveFormat = explicitFormat
                ? Format
                : PSSerializationFormatDetector.DetectFromBytes(bytes, PSSerializationFormat.Cbor);

            using MemoryStream ms = new(bytes, writable: false);
            DispatchStreamDeserialization(ms, effectiveFormat);
        }

        private void DeserializeFromStream(Stream stream)
        {
            bool explicitFormat = MyInvocation.BoundParameters.ContainsKey(nameof(Format));
            if (explicitFormat || stream.CanSeek)
            {
                PSSerializationFormat effectiveFormat = explicitFormat
                    ? Format
                    : PSSerializationFormatDetector.DetectFromStream(stream);
                DispatchStreamDeserialization(stream, effectiveFormat);
            }
            else
            {
                using MemoryStream ms = new();
                stream.CopyTo(ms);
                byte[] bytes = ms.ToArray();
                PSSerializationFormat effectiveFormat = PSSerializationFormatDetector.DetectFromBytes(bytes, PSSerializationFormat.Cbor);
                ms.Position = 0;
                DispatchStreamDeserialization(ms, effectiveFormat);
            }
        }

        private void DispatchStreamDeserialization(Stream stream, PSSerializationFormat format)
        {
            switch (format)
            {
                case PSSerializationFormat.Protobuf:
                    _protoHelper = new ImportProtoHelper(this);
                    _protoHelper.ImportStreamContent(stream);
                    break;

                case PSSerializationFormat.Cbor:
                    _cborHelper = new ImportCborHelper(this);
                    _cborHelper.ImportStreamContent(stream);
                    break;

                case PSSerializationFormat.Clixml:
                    _cliXmlHelper = new ImportCliXmlHelper(this);
                    _cliXmlHelper.ImportStreamContent(stream);
                    break;
            }
        }

        /// <summary>
        /// Releases all resources used by the cmdlet.
        /// </summary>
        public void Dispose()
        {
            _protoHelper?.Dispose();
            _cborHelper?.Dispose();
            _cliXmlHelper?.Dispose();
            _pipedBytesBuffer?.Dispose();
            GC.SuppressFinalize(this);
        }
    }

    #endregion ConvertFromPSObjectCommand

    #region Internal Serialization & Import Helpers

    /// <summary>
    /// Transforms PowerShell encoding arguments (strings, code pages, or Encoding instances) into <see cref="Encoding"/>.
    /// </summary>
    internal sealed class EncodingTransformationAttribute : ArgumentTransformationAttribute
    {
        internal static readonly Encoding Utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
        {
            object raw = inputData is PSObject pso ? pso.BaseObject : inputData;
            switch (raw)
            {
                case Encoding encoding:
                    return encoding;

                case string name:
                    return ResolveEncodingName(name);

                case int codePage:
                    return Encoding.GetEncoding(codePage);

                default:
                    return inputData;
            }
        }

        internal static Encoding ResolveEncodingName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return Utf8NoBom;
            }

            switch (name.Trim().ToLowerInvariant())
            {
                case "utf8":
                case "utf8nobom":
                case "default":
                    return Utf8NoBom;
                case "utf8bom":
                    return Encoding.UTF8;
                case "unicode":
                case "utf16":
                case "utf16le":
                    return Encoding.Unicode;
                case "bigendianunicode":
                case "utf16be":
                    return Encoding.BigEndianUnicode;
                case "utf32":
                case "utf32le":
                    return Encoding.UTF32;
                case "bigendianutf32":
                case "utf32be":
                    return new UTF32Encoding(bigEndian: true, byteOrderMark: true);
                case "ascii":
                    return Encoding.ASCII;
                case "oem":
                    return Console.OutputEncoding ?? Utf8NoBom;
#pragma warning disable SYSLIB0001
                case "utf7":
                    return Encoding.UTF7;
#pragma warning restore SYSLIB0001
                default:
                    if (int.TryParse(name, NumberStyles.Integer, CultureInfo.InvariantCulture, out int cp))
                    {
                        return Encoding.GetEncoding(cp);
                    }

                    return Encoding.GetEncoding(name);
            }
        }
    }

    /// <summary>
    /// Reflection-backed wrapper around System.Management.Automation.Serializer.
    /// </summary>
    internal sealed class CliXmlSerializer
    {
        private static readonly Type s_smaSerializerType =
            typeof(PSObject).Assembly.GetType("System.Management.Automation.Serializer");

        private static readonly ConstructorInfo s_ctorWithDepth =
            s_smaSerializerType?.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                new Type[] { typeof(XmlWriter), typeof(int), typeof(bool) },
                modifiers: null);

        private static readonly MethodInfo s_serializeMethod =
            s_smaSerializerType?.GetMethod(
                "Serialize",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                new Type[] { typeof(object) },
                modifiers: null);

        private static readonly MethodInfo s_doneMethod =
            s_smaSerializerType?.GetMethod(
                "Done",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                Type.EmptyTypes,
                modifiers: null);

        private static readonly MethodInfo s_stopMethod =
            s_smaSerializerType?.GetMethod(
                "Stop",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                Type.EmptyTypes,
                modifiers: null);

        private readonly object _instance;
        private readonly XmlWriter _xmlWriter;

        internal CliXmlSerializer(XmlWriter xmlWriter, int depth, bool useDepthFromTypes)
        {
            if (xmlWriter is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(xmlWriter));
            }

            if (s_ctorWithDepth is null || s_serializeMethod is null || s_doneMethod is null)
            {
                throw new InvalidOperationException("Failed to initialize System.Management.Automation.Serializer reflection bindings.");
            }

            _xmlWriter = xmlWriter;
            try
            {
                _instance = s_ctorWithDepth.Invoke(new object[] { xmlWriter, depth, useDepthFromTypes });
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
            }
        }

        internal void Serialize(object source)
        {
            try
            {
                s_serializeMethod.Invoke(_instance, new object[] { source });
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
            }
        }

        internal void Done()
        {
            try
            {
                s_doneMethod.Invoke(_instance, Array.Empty<object>());
                _xmlWriter.Flush();
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
            }
        }

        internal void Stop()
        {
            try
            {
                s_stopMethod?.Invoke(_instance, Array.Empty<object>());
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
            }
        }
    }

    /// <summary>
    /// Reflection-backed wrapper around System.Management.Automation.Deserializer.
    /// </summary>
    internal sealed class CliXmlDeserializer
    {
        private static readonly Type s_smaDeserializerType =
            typeof(PSObject).Assembly.GetType("System.Management.Automation.Deserializer");

        private static readonly Type s_smaInternalDeserializerType =
            typeof(PSObject).Assembly.GetType("System.Management.Automation.InternalDeserializer");

        private static readonly ConstructorInfo s_ctor =
            s_smaDeserializerType?.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                new Type[] { typeof(XmlReader) },
                modifiers: null);

        private static readonly MethodInfo s_doneMethod =
            s_smaDeserializerType?.GetMethod(
                "Done",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                Type.EmptyTypes,
                modifiers: null);

        private static readonly MethodInfo s_deserializeMethod =
            s_smaDeserializerType?.GetMethod(
                "Deserialize",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                Type.EmptyTypes,
                modifiers: null);

        private static readonly MethodInfo s_stopMethod =
            s_smaDeserializerType?.GetMethod(
                "Stop",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                Type.EmptyTypes,
                modifiers: null);

        private static readonly PropertyInfo s_xmlReaderSettingsProp =
            s_smaInternalDeserializerType?.GetProperty(
                "XmlReaderSettingsForCliXml",
                BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);

        private readonly object _instance;

        internal CliXmlDeserializer(XmlReader xmlReader)
        {
            if (xmlReader is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(xmlReader));
            }

            if (s_ctor is null || s_doneMethod is null || s_deserializeMethod is null)
            {
                throw new InvalidOperationException("Failed to initialize System.Management.Automation.Deserializer reflection bindings.");
            }

            try
            {
                _instance = s_ctor.Invoke(new object[] { xmlReader });
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
            }
        }

        internal static XmlReaderSettings GetXmlReaderSettingsForCliXml()
        {
            if (s_xmlReaderSettingsProp?.GetValue(null) is XmlReaderSettings settings)
            {
                return settings.Clone();
            }

            return new XmlReaderSettings
            {
                CheckCharacters = false,
                IgnoreComments = false,
                IgnoreProcessingInstructions = false,
                MaxCharactersInDocument = 0,
                XmlResolver = null,
                CloseInput = false
            };
        }

        internal static XmlReader CreateXmlReader(Stream stream)
        {
            StreamReader textReader = new(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);

            const string cliXmlDirective = "#< CLIXML";
            if (textReader.Peek() == (int)cliXmlDirective[0])
            {
                string line = textReader.ReadLine();
                if (!string.Equals(line, cliXmlDirective, StringComparison.Ordinal) && stream.CanSeek)
                {
                    stream.Seek(0, SeekOrigin.Begin);
                    textReader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true);
                }
            }

            return XmlReader.Create(textReader, GetXmlReaderSettingsForCliXml());
        }

        internal static XmlReader CreateXmlReader(string xmlContent)
        {
            StringReader textReader = new(xmlContent);

            const string cliXmlDirective = "#< CLIXML";
            if (textReader.Peek() == (int)cliXmlDirective[0])
            {
                string line = textReader.ReadLine();
                if (!string.Equals(line, cliXmlDirective, StringComparison.Ordinal))
                {
                    textReader = new StringReader(xmlContent);
                }
            }

            return XmlReader.Create(textReader, GetXmlReaderSettingsForCliXml());
        }

        internal bool Done()
        {
            try
            {
                return (bool)s_doneMethod.Invoke(_instance, Array.Empty<object>());
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
                return true;
            }
        }

        internal object Deserialize()
        {
            try
            {
                return s_deserializeMethod.Invoke(_instance, Array.Empty<object>());
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
                return null;
            }
        }

        internal void Stop()
        {
            try
            {
                s_stopMethod?.Invoke(_instance, Array.Empty<object>());
            }
            catch (TargetInvocationException tie) when (tie.InnerException is not null)
            {
                ExceptionDispatchInfo.Capture(tie.InnerException).Throw();
            }
        }
    }

    /// <summary>
    /// Helper for Protobuf stream deserialization.
    /// </summary>
    internal sealed class ImportProtoHelper : IDisposable
    {
        private readonly PSCmdlet _cmdlet;
        private ProtoDeserializer _deserializer;

        internal ImportProtoHelper(PSCmdlet cmdlet, string path = null, bool isLiteralPath = false)
        {
            _cmdlet = cmdlet;
        }

        internal void ImportFromOpenStream(Stream fs)
        {
            if (_cmdlet.PagingParameters.IncludeTotalCount)
            {
                PSObject totalCount = _cmdlet.PagingParameters.NewTotalCount(0, 0.0);
                _cmdlet.WriteObject(totalCount);
            }

            ImportStreamContent(fs, usePaging: true);
        }

        internal void ImportStreamContent(Stream stream, bool usePaging = false)
        {
            _deserializer = new ProtoDeserializer(stream);
            ulong skipped = 0;
            ulong paged = 0;

            while (!_deserializer.Done())
            {
                object result = _deserializer.Deserialize();
                if (usePaging)
                {
                    if (skipped++ < _cmdlet.PagingParameters.Skip)
                    {
                        continue;
                    }

                    if (paged++ >= _cmdlet.PagingParameters.First)
                    {
                        break;
                    }
                }

                _cmdlet.WriteObject(result);
            }
        }

        internal void Stop()
        {
            _deserializer?.Stop();
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }

    /// <summary>
    /// Helper for CBOR stream deserialization.
    /// </summary>
    internal sealed class ImportCborHelper : IDisposable
    {
        private readonly PSCmdlet _cmdlet;
        private CborDeserializer _deserializer;

        internal ImportCborHelper(PSCmdlet cmdlet, string path = null, bool isLiteralPath = false)
        {
            _cmdlet = cmdlet;
        }

        internal void ImportFromOpenStream(Stream fs)
        {
            if (_cmdlet.PagingParameters.IncludeTotalCount)
            {
                PSObject totalCount = _cmdlet.PagingParameters.NewTotalCount(0, 0.0);
                _cmdlet.WriteObject(totalCount);
            }

            ImportStreamContent(fs, usePaging: true);
        }

        internal void ImportStreamContent(Stream stream, bool usePaging = false)
        {
            _deserializer = new CborDeserializer(stream);
            ulong skipped = 0;
            ulong paged = 0;

            while (!_deserializer.Done())
            {
                object result = _deserializer.Deserialize();
                if (usePaging)
                {
                    if (skipped++ < _cmdlet.PagingParameters.Skip)
                    {
                        continue;
                    }

                    if (paged++ >= _cmdlet.PagingParameters.First)
                    {
                        break;
                    }
                }

                _cmdlet.WriteObject(result);
            }
        }

        internal void Stop()
        {
            _deserializer?.Stop();
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }

    /// <summary>
    /// Helper for CLIXML stream/string deserialization.
    /// </summary>
    internal sealed class ImportCliXmlHelper : IDisposable
    {
        private readonly PSCmdlet _cmdlet;
        private CliXmlDeserializer _deserializer;

        internal ImportCliXmlHelper(PSCmdlet cmdlet, string path = null, bool isLiteralPath = false)
        {
            _cmdlet = cmdlet;
        }

        internal void ImportFromOpenStream(Stream fs)
        {
            if (_cmdlet.PagingParameters.IncludeTotalCount)
            {
                PSObject totalCount = _cmdlet.PagingParameters.NewTotalCount(0, 0.0);
                _cmdlet.WriteObject(totalCount);
            }

            ImportStreamContent(fs, usePaging: true);
        }

        internal void ImportStreamContent(Stream stream, bool usePaging = false)
        {
            using XmlReader xmlReader = CliXmlDeserializer.CreateXmlReader(stream);
            _deserializer = new CliXmlDeserializer(xmlReader);
            ulong skipped = 0;
            ulong paged = 0;

            while (!_deserializer.Done())
            {
                object result = _deserializer.Deserialize();
                if (usePaging)
                {
                    if (skipped++ < _cmdlet.PagingParameters.Skip)
                    {
                        continue;
                    }

                    if (paged++ >= _cmdlet.PagingParameters.First)
                    {
                        break;
                    }
                }

                _cmdlet.WriteObject(result);
            }
        }

        internal void ImportStringContent(string xmlContent)
        {
            using XmlReader xmlReader = CliXmlDeserializer.CreateXmlReader(xmlContent);
            _deserializer = new CliXmlDeserializer(xmlReader);
            while (!_deserializer.Done())
            {
                object result = _deserializer.Deserialize();
                _cmdlet.WriteObject(result);
            }
        }

        internal void Stop()
        {
            _deserializer?.Stop();
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }

    #endregion Internal Serialization & Import Helpers
}
