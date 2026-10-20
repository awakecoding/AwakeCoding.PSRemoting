using System;
using System.Buffers.Binary;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Reflection;
using System.Text;
using System.Xml;
using AwakeCoding.PSRemoting.PowerShell.Grpc;
using Google.Protobuf;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Converts PSRP stdio XML packets and CLIXML-encoded fragments to and from
    /// gRPC frames carrying Protobuf-encoded PSObject payloads (<see cref="PSProtoSerializer"/>).
    /// </summary>
    internal sealed class PSHostGrpcPsrpCodec
    {
        private const int FragmentHeaderLength = 21;
        private const int RemoteDataObjectHeaderLength = 40;
        private const byte FragmentFlagStart = 0x01;
        private const byte FragmentFlagEnd = 0x02;
        private const byte FragmentFlagStartAndEnd = FragmentFlagStart | FragmentFlagEnd;
        private const int DefaultMaxFragmentBlobSize = 32768 - FragmentHeaderLength;

        private static readonly Encoding s_utf8NoBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

        private static readonly Type? s_smaSerializerType = typeof(PSObject).Assembly.GetType("System.Management.Automation.Serializer");
        private static readonly Type? s_smaDeserializerType = typeof(PSObject).Assembly.GetType("System.Management.Automation.Deserializer");
        private static readonly Type? s_smaSerializationContextType = typeof(PSObject).Assembly.GetType("System.Management.Automation.SerializationContext");
        private static readonly Type? s_smaDeserializationContextType = typeof(PSObject).Assembly.GetType("System.Management.Automation.DeserializationContext");
        private static readonly Type? s_smaSerializationOptionsType = typeof(PSObject).Assembly.GetType("System.Management.Automation.SerializationOptions");
        private static readonly Type? s_smaDeserializationOptionsType = typeof(PSObject).Assembly.GetType("System.Management.Automation.DeserializationOptions");

        private static readonly object? s_smaRemotingSerOptions = s_smaSerializationOptionsType != null
            ? Enum.Parse(s_smaSerializationOptionsType, "RemotingOptions")
            : null;
        private static readonly object? s_smaRemotingDesOptions = s_smaDeserializationOptionsType != null
            ? Enum.Parse(s_smaDeserializationOptionsType, "RemotingOptions")
            : null;

        private static readonly ConstructorInfo? s_smaSerCtxCtor = s_smaSerializationContextType?.GetConstructor(
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            new Type[] { typeof(int), s_smaSerializationOptionsType!, typeof(PSRemotingCryptoHelper) },
            modifiers: null);

        private static readonly ConstructorInfo? s_smaDesCtxCtor = s_smaDeserializationContextType?.GetConstructor(
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            new Type[] { s_smaDeserializationOptionsType!, typeof(PSRemotingCryptoHelper) },
            modifiers: null);

        private static readonly ConstructorInfo? s_smaSerializerCtor = s_smaSerializerType != null && s_smaSerializationContextType != null
            ? s_smaSerializerType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                new Type[] { typeof(XmlWriter), s_smaSerializationContextType },
                modifiers: null)
            : null;

        private static readonly ConstructorInfo? s_smaDeserializerCtor = s_smaDeserializerType != null && s_smaDeserializationContextType != null
            ? s_smaDeserializerType.GetConstructor(
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                binder: null,
                new Type[] { typeof(XmlReader), s_smaDeserializationContextType },
                modifiers: null)
            : null;

        private static readonly MethodInfo? s_smaSerializeMethod = s_smaSerializerType?.GetMethod(
            "Serialize",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            new Type[] { typeof(object) },
            modifiers: null);

        private static readonly MethodInfo? s_smaSerializerDoneMethod = s_smaSerializerType?.GetMethod(
            "Done",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            Type.EmptyTypes,
            modifiers: null);

        private static readonly MethodInfo? s_smaDeserializeMethod = s_smaDeserializerType?.GetMethod(
            "Deserialize",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            Type.EmptyTypes,
            modifiers: null);

        private static readonly MethodInfo? s_smaDeserializerDoneMethod = s_smaDeserializerType?.GetMethod(
            "Done",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            Type.EmptyTypes,
            modifiers: null);

        private static readonly Type? s_baseClientTransportManagerType = typeof(PSObject).Assembly.GetType(
            "System.Management.Automation.Remoting.Client.BaseClientTransportManager");

        private static readonly MethodInfo? s_onDataAvailableHandlerMethod = s_baseClientTransportManagerType?.GetMethod(
            "OnDataAvailableHandler",
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);

        private static readonly Type? s_remoteDataObjectOfPSObjectType = s_onDataAvailableHandlerMethod?.GetParameters()[0].ParameterType;

        private static readonly MethodInfo? s_rdoCreateFromMethod = FindRdoCreateFromMethod();
        private static readonly Type? s_remotingDestinationType = s_rdoCreateFromMethod?.GetParameters()[0].ParameterType;
        private static readonly Type? s_remotingDataTypeType = s_rdoCreateFromMethod?.GetParameters()[1].ParameterType;

        private static MethodInfo? FindRdoCreateFromMethod()
        {
            if (s_remoteDataObjectOfPSObjectType == null)
            {
                return null;
            }

            foreach (var method in s_remoteDataObjectOfPSObjectType.GetMethods(BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
            {
                if (method.Name == "CreateFrom" && method.GetParameters().Length == 5)
                {
                    return method;
                }
            }

            return null;
        }

        private readonly Dictionary<ulong, MemoryStream> _outboundInFlightObjects = new Dictionary<ulong, MemoryStream>();
        private readonly Dictionary<ulong, ReassemblyBuffer> _inboundInFlightObjects = new Dictionary<ulong, ReassemblyBuffer>();

        public PSHostGrpcEncoding Encoding { get; set; }

        public PSHostGrpcPsrpCodec()
            : this(PSHostGrpcEncoding.Protobuf)
        {
        }

        public PSHostGrpcPsrpCodec(PSHostGrpcEncoding encoding)
        {
            Encoding = encoding;
        }

        private sealed class ReassemblyBuffer
        {
            public uint Destination;
            public uint DataType;
            public PSHostGrpcPayloadEncoding PayloadEncoding;
            public byte[] RunspacePoolId = Array.Empty<byte>();
            public byte[] PowerShellId = Array.Empty<byte>();
            public readonly MemoryStream PayloadStream = new MemoryStream();
        }

        /// <summary>
        /// Encodes a single OutOfProcess PSRP XML packet line into one or more Protobuf-backed <see cref="PSHostGrpcFrame"/> messages.
        /// </summary>
        public IEnumerable<PSHostGrpcFrame> EncodePacketLineToFrames(string line)
        {
            string trimmed = line.Trim();
            if (trimmed.Length == 0)
            {
                yield break;
            }

            using var stringReader = new StringReader(trimmed);
            var readerSettings = new XmlReaderSettings
            {
                CheckCharacters = false,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = false,
                XmlResolver = null,
                ConformanceLevel = ConformanceLevel.Fragment
            };

            using var xmlReader = XmlReader.Create(stringReader, readerSettings);
            if (!xmlReader.Read() || xmlReader.NodeType != XmlNodeType.Element)
            {
                throw new PSRemotingTransportException($"Invalid PSRP packet line: {trimmed}");
            }

            string elementName = xmlReader.Name;
            string psGuid = xmlReader.GetAttribute("PSGuid") ?? string.Empty;
            string streamAttr = xmlReader.GetAttribute("Stream") ?? "Default";

            switch (elementName)
            {
                case "DataAck":
                    yield return new PSHostGrpcFrame { Kind = PSHostGrpcFrameKind.DataAck, PsGuid = psGuid };
                    yield break;

                case "Command":
                    yield return new PSHostGrpcFrame { Kind = PSHostGrpcFrameKind.Command, PsGuid = psGuid };
                    yield break;

                case "CommandAck":
                    yield return new PSHostGrpcFrame { Kind = PSHostGrpcFrameKind.CommandAck, PsGuid = psGuid };
                    yield break;

                case "Close":
                    yield return new PSHostGrpcFrame { Kind = PSHostGrpcFrameKind.ClosePacket, PsGuid = psGuid };
                    yield break;

                case "CloseAck":
                    yield return new PSHostGrpcFrame { Kind = PSHostGrpcFrameKind.CloseAck, PsGuid = psGuid };
                    yield break;

                case "Signal":
                    yield return new PSHostGrpcFrame { Kind = PSHostGrpcFrameKind.Signal, PsGuid = psGuid };
                    yield break;

                case "SignalAck":
                    yield return new PSHostGrpcFrame { Kind = PSHostGrpcFrameKind.SignalAck, PsGuid = psGuid };
                    yield break;

                case "Data":
                    string base64Content = xmlReader.ReadElementContentAsString();
                    byte[] rawBytes = Convert.FromBase64String(base64Content);
                    foreach (var frame in EncodeRawFragmentToFrames(rawBytes, psGuid, streamAttr))
                    {
                        yield return frame;
                    }
                    yield break;

                default:
                    throw new PSRemotingTransportException($"Unsupported PSRP packet element '{elementName}'.");
            }
        }

        /// <summary>
        /// Encodes raw PSRP fragment bytes directly into one or more Protobuf-backed <see cref="PSHostGrpcFrame"/> messages,
        /// bypassing <c>OutOfProcessUtils.CreateDataPacket</c> XML and Base64 encoding.
        /// </summary>
        public IEnumerable<PSHostGrpcFrame> EncodeRawFragmentToFrames(byte[] rawBytes, string psGuid, string streamAttr = "Default")
        {
            int offset = 0;
            while (offset + FragmentHeaderLength <= rawBytes.Length)
            {
                ulong objectId = BinaryPrimitives.ReadUInt64BigEndian(rawBytes.AsSpan(offset, 8));
                ulong fragmentId = BinaryPrimitives.ReadUInt64BigEndian(rawBytes.AsSpan(offset + 8, 8));
                byte flags = rawBytes[offset + 16];
                int blobLength = BinaryPrimitives.ReadInt32BigEndian(rawBytes.AsSpan(offset + 17, 4));
                offset += FragmentHeaderLength;

                if (blobLength < 0 || offset + blobLength > rawBytes.Length)
                {
                    throw new PSRemotingTransportException("Invalid PSRP fragment length.");
                }

                bool isStart = (flags & FragmentFlagStart) != 0;
                bool isEnd = (flags & FragmentFlagEnd) != 0;

                MemoryStream? objectStream;
                if (isStart && isEnd)
                {
                    objectStream = new MemoryStream(rawBytes, offset, blobLength, writable: false);
                }
                else if (isStart)
                {
                    objectStream = new MemoryStream();
                    objectStream.Write(rawBytes, offset, blobLength);
                    _outboundInFlightObjects[objectId] = objectStream;
                    offset += blobLength;
                    continue;
                }
                else
                {
                    if (!_outboundInFlightObjects.TryGetValue(objectId, out objectStream))
                    {
                        throw new PSRemotingTransportException($"Missing start fragment for PSRP ObjectId {objectId}.");
                    }

                    objectStream.Write(rawBytes, offset, blobLength);
                    if (!isEnd)
                    {
                        offset += blobLength;
                        continue;
                    }

                    _outboundInFlightObjects.Remove(objectId);
                    objectStream.Position = 0;
                }

                offset += blobLength;

                byte[] fullObjectBytes = objectStream.ToArray();
                objectStream.Dispose();

                if (fullObjectBytes.Length < RemoteDataObjectHeaderLength)
                {
                    throw new PSRemotingTransportException("Reassembled PSRP RemoteDataObject is shorter than 40-byte header.");
                }

                uint destination = BinaryPrimitives.ReadUInt32LittleEndian(fullObjectBytes.AsSpan(0, 4));
                uint dataType = BinaryPrimitives.ReadUInt32LittleEndian(fullObjectBytes.AsSpan(4, 4));
                byte[] runspacePoolId = fullObjectBytes.AsSpan(8, 16).ToArray();
                byte[] powerShellId = fullObjectBytes.AsSpan(24, 16).ToArray();

                PSHostGrpcPayloadEncoding wireEncoding = Encoding == PSHostGrpcEncoding.Cbor
                    ? PSHostGrpcPayloadEncoding.Cbor
                    : PSHostGrpcPayloadEncoding.Protobuf;

                byte[] binaryPayloadBytes;
                int clixmlLength = fullObjectBytes.Length - RemoteDataObjectHeaderLength;
                if (clixmlLength > 0)
                {
                    object? psObj = DeserializeClixmlRemotingPayload(fullObjectBytes, RemoteDataObjectHeaderLength, clixmlLength);
                    binaryPayloadBytes = wireEncoding == PSHostGrpcPayloadEncoding.Cbor
                        ? SerializeCborRemotingPayload(psObj)
                        : SerializeProtoRemotingPayload(psObj);
                }
                else
                {
                    binaryPayloadBytes = Array.Empty<byte>();
                }

                if (binaryPayloadBytes.Length <= DefaultMaxFragmentBlobSize)
                {
                    yield return new PSHostGrpcFrame
                    {
                        Kind = PSHostGrpcFrameKind.Data,
                        PsGuid = psGuid,
                        Stream = streamAttr,
                        ObjectId = objectId,
                        FragmentId = 0,
                        BlobFlags = FragmentFlagStartAndEnd,
                        Destination = destination,
                        DataType = dataType,
                        RunspacePoolId = ByteString.CopyFrom(runspacePoolId),
                        PowershellId = ByteString.CopyFrom(powerShellId),
                        Payload = ByteString.CopyFrom(binaryPayloadBytes),
                        PayloadEncoding = wireEncoding
                    };
                }
                else
                {
                    int protoOffset = 0;
                    ulong outFragId = 0;
                    while (protoOffset < binaryPayloadBytes.Length)
                    {
                        int chunkLen = Math.Min(DefaultMaxFragmentBlobSize, binaryPayloadBytes.Length - protoOffset);
                        byte outFlags = 0;
                        if (protoOffset == 0)
                        {
                            outFlags |= FragmentFlagStart;
                        }
                        if (protoOffset + chunkLen >= binaryPayloadBytes.Length)
                        {
                            outFlags |= FragmentFlagEnd;
                        }

                        yield return new PSHostGrpcFrame
                        {
                            Kind = PSHostGrpcFrameKind.Data,
                            PsGuid = psGuid,
                            Stream = streamAttr,
                            ObjectId = objectId,
                            FragmentId = outFragId++,
                            BlobFlags = outFlags,
                            Destination = destination,
                            DataType = dataType,
                            RunspacePoolId = ByteString.CopyFrom(runspacePoolId),
                            PowershellId = ByteString.CopyFrom(powerShellId),
                            Payload = ByteString.CopyFrom(binaryPayloadBytes, protoOffset, chunkLen),
                            PayloadEncoding = wireEncoding
                        };

                        protoOffset += chunkLen;
                    }
                }
            }
        }

        /// <summary>
        /// Decodes a <see cref="PSHostGrpcFrame"/> into an OutOfProcess PSRP XML packet line,
        /// or returns <c>null</c> if the frame is a non-final fragment of a multi-fragment PSRP object.
        /// </summary>
        public string? DecodeFrameToPacketLine(PSHostGrpcFrame frame)
        {
            switch (frame.Kind)
            {
                case PSHostGrpcFrameKind.DataAck:
                    return $"<DataAck PSGuid='{frame.PsGuid}' />";

                case PSHostGrpcFrameKind.Command:
                    return $"<Command PSGuid='{frame.PsGuid}' />";

                case PSHostGrpcFrameKind.CommandAck:
                    return $"<CommandAck PSGuid='{frame.PsGuid}' />";

                case PSHostGrpcFrameKind.ClosePacket:
                    return $"<Close PSGuid='{frame.PsGuid}' />";

                case PSHostGrpcFrameKind.CloseAck:
                    return $"<CloseAck PSGuid='{frame.PsGuid}' />";

                case PSHostGrpcFrameKind.Signal:
                    return $"<Signal PSGuid='{frame.PsGuid}' />";

                case PSHostGrpcFrameKind.SignalAck:
                    return $"<SignalAck PSGuid='{frame.PsGuid}' />";

                case PSHostGrpcFrameKind.Data:
                    return DecodeDataFrameToPacketLine(frame);

                default:
                    return null;
            }
        }

        /// <summary>
        /// Decodes an inbound <see cref="PSHostGrpcFrameKind.Data"/> frame directly into a
        /// <c>RemoteDataObject&lt;PSObject&gt;</c> using <see cref="ProtoDeserializer"/> or <see cref="CborDeserializer"/>
        /// without re-serializing to CLIXML, Base64, or <c>&lt;Data&gt;</c> XML.
        /// Returns <c>true</c> when a complete PSRP object has been reassembled and decoded.
        /// </summary>
        public bool TryDecodeDataFrameToRemoteDataObject(
            PSHostGrpcFrame frame,
            out Guid psGuid,
            out object? remoteDataObject)
        {
            psGuid = ParsePsGuid(frame.PsGuid);
            remoteDataObject = null;

            if (!TryReassembleInboundDataFrame(
                frame,
                out _,
                out uint destination,
                out uint dataType,
                out PSHostGrpcPayloadEncoding payloadEncoding,
                out byte[] runspacePoolId,
                out byte[] powerShellId,
                out byte[] binaryPayloadBytes))
            {
                return false;
            }

            if (s_rdoCreateFromMethod == null || s_remotingDestinationType == null || s_remotingDataTypeType == null)
            {
                throw new InvalidOperationException("Failed to resolve RemoteDataObject<PSObject>.CreateFrom reflection binding.");
            }

            PSObject? psObj = null;
            if (binaryPayloadBytes.Length > 0)
            {
                psObj = (payloadEncoding == PSHostGrpcPayloadEncoding.Cbor
                    ? DeserializeCborRemotingPayload(binaryPayloadBytes)
                    : DeserializeProtoRemotingPayload(binaryPayloadBytes)) as PSObject;
            }

            Guid rsPoolGuid = runspacePoolId.Length == 16 ? new Guid(runspacePoolId) : Guid.Empty;
            Guid psInstanceGuid = powerShellId.Length == 16 ? new Guid(powerShellId) : Guid.Empty;

            remoteDataObject = s_rdoCreateFromMethod.Invoke(
                obj: null,
                new object?[]
                {
                    Enum.ToObject(s_remotingDestinationType, destination),
                    Enum.ToObject(s_remotingDataTypeType, dataType),
                    rsPoolGuid,
                    psInstanceGuid,
                    psObj!
                });

            return remoteDataObject != null;
        }

        internal static Guid ParsePsGuid(string? psGuidText)
        {
            if (string.IsNullOrEmpty(psGuidText))
            {
                return Guid.Empty;
            }

            return Guid.TryParse(psGuidText, out Guid parsed) ? parsed : Guid.Empty;
        }

        private bool TryReassembleInboundDataFrame(
            PSHostGrpcFrame frame,
            out ulong objectId,
            out uint destination,
            out uint dataType,
            out PSHostGrpcPayloadEncoding payloadEncoding,
            out byte[] runspacePoolId,
            out byte[] powerShellId,
            out byte[] protoPayloadBytes)
        {
            objectId = frame.ObjectId;
            byte flags = (byte)frame.BlobFlags;
            bool isStart = (flags & FragmentFlagStart) != 0;
            bool isEnd = (flags & FragmentFlagEnd) != 0;

            if (isStart && isEnd)
            {
                destination = frame.Destination;
                dataType = frame.DataType;
                payloadEncoding = frame.PayloadEncoding;
                runspacePoolId = frame.RunspacePoolId.ToByteArray();
                powerShellId = frame.PowershellId.ToByteArray();
                protoPayloadBytes = frame.Payload.ToByteArray();
                return true;
            }

            if (isStart)
            {
                var buffer = new ReassemblyBuffer
                {
                    Destination = frame.Destination,
                    DataType = frame.DataType,
                    PayloadEncoding = frame.PayloadEncoding,
                    RunspacePoolId = frame.RunspacePoolId.ToByteArray(),
                    PowerShellId = frame.PowershellId.ToByteArray()
                };
                frame.Payload.WriteTo(buffer.PayloadStream);
                _inboundInFlightObjects[objectId] = buffer;

                destination = 0;
                dataType = 0;
                payloadEncoding = PSHostGrpcPayloadEncoding.Protobuf;
                runspacePoolId = Array.Empty<byte>();
                powerShellId = Array.Empty<byte>();
                protoPayloadBytes = Array.Empty<byte>();
                return false;
            }

            if (!_inboundInFlightObjects.TryGetValue(objectId, out var existing))
            {
                throw new PSRemotingTransportException($"Missing start fragment for inbound gRPC PSRP ObjectId {objectId}.");
            }

            frame.Payload.WriteTo(existing.PayloadStream);
            if (!isEnd)
            {
                destination = 0;
                dataType = 0;
                payloadEncoding = PSHostGrpcPayloadEncoding.Protobuf;
                runspacePoolId = Array.Empty<byte>();
                powerShellId = Array.Empty<byte>();
                protoPayloadBytes = Array.Empty<byte>();
                return false;
            }

            _inboundInFlightObjects.Remove(objectId);
            destination = existing.Destination;
            dataType = existing.DataType;
            payloadEncoding = existing.PayloadEncoding;
            runspacePoolId = existing.RunspacePoolId;
            powerShellId = existing.PowerShellId;
            protoPayloadBytes = existing.PayloadStream.ToArray();
            existing.PayloadStream.Dispose();
            return true;
        }

        private string? DecodeDataFrameToPacketLine(PSHostGrpcFrame frame)
        {
            if (!TryReassembleInboundDataFrame(
                frame,
                out ulong objectId,
                out uint destination,
                out uint dataType,
                out PSHostGrpcPayloadEncoding payloadEncoding,
                out byte[] runspacePoolId,
                out byte[] powerShellId,
                out byte[] binaryPayloadBytes))
            {
                return null;
            }

            byte[] clixmlPayloadBytes;
            if (binaryPayloadBytes.Length > 0)
            {
                object? psObj = payloadEncoding == PSHostGrpcPayloadEncoding.Cbor
                    ? DeserializeCborRemotingPayload(binaryPayloadBytes)
                    : DeserializeProtoRemotingPayload(binaryPayloadBytes);
                clixmlPayloadBytes = SerializeClixmlRemotingPayload(psObj);
            }
            else
            {
                clixmlPayloadBytes = Array.Empty<byte>();
            }

            int totalRdoLength = RemoteDataObjectHeaderLength + clixmlPayloadBytes.Length;
            byte[] rdoBytes = new byte[totalRdoLength];
            BinaryPrimitives.WriteUInt32LittleEndian(rdoBytes.AsSpan(0, 4), destination);
            BinaryPrimitives.WriteUInt32LittleEndian(rdoBytes.AsSpan(4, 4), dataType);
            if (runspacePoolId.Length == 16)
            {
                Buffer.BlockCopy(runspacePoolId, 0, rdoBytes, 8, 16);
            }
            if (powerShellId.Length == 16)
            {
                Buffer.BlockCopy(powerShellId, 0, rdoBytes, 24, 16);
            }
            if (clixmlPayloadBytes.Length > 0)
            {
                Buffer.BlockCopy(clixmlPayloadBytes, 0, rdoBytes, RemoteDataObjectHeaderLength, clixmlPayloadBytes.Length);
            }

            using var fragmentsStream = new MemoryStream();
            Span<byte> header = stackalloc byte[FragmentHeaderLength];
            int rdoOffset = 0;
            ulong fragId = 0;
            while (rdoOffset < rdoBytes.Length)
            {
                int chunkLen = Math.Min(DefaultMaxFragmentBlobSize, rdoBytes.Length - rdoOffset);
                byte outFlags = 0;
                if (rdoOffset == 0)
                {
                    outFlags |= FragmentFlagStart;
                }
                if (rdoOffset + chunkLen >= rdoBytes.Length)
                {
                    outFlags |= FragmentFlagEnd;
                }

                BinaryPrimitives.WriteUInt64BigEndian(header.Slice(0, 8), objectId);
                BinaryPrimitives.WriteUInt64BigEndian(header.Slice(8, 8), fragId++);
                header[16] = outFlags;
                BinaryPrimitives.WriteInt32BigEndian(header.Slice(17, 4), chunkLen);

                fragmentsStream.Write(header);
                fragmentsStream.Write(rdoBytes, rdoOffset, chunkLen);
                rdoOffset += chunkLen;
            }

            string base64Payload = Convert.ToBase64String(fragmentsStream.ToArray());
            string streamName = string.IsNullOrEmpty(frame.Stream) ? "Default" : frame.Stream;
            return string.Format(
                CultureInfo.InvariantCulture,
                "<Data Stream='{0}' PSGuid='{1}'>{2}</Data>",
                streamName,
                frame.PsGuid,
                base64Payload);
        }

        internal static byte[] SerializeProtoRemotingPayload(object? obj)
        {
            using var ms = new MemoryStream();
            var context = new SerializationContext(
                depth: 1,
                options: SerializationOptions.RemotingOptions,
                cryptoHelper: null);
            var serializer = new ProtoSerializer(ms, context);
            serializer.Serialize(obj!);
            serializer.Done();
            return ms.ToArray();
        }

        internal static object? DeserializeProtoRemotingPayload(byte[] protoBytes)
        {
            using var ms = new MemoryStream(protoBytes, writable: false);
            var context = new DeserializationContext(
                options: DeserializationOptions.RemotingOptions,
                cryptoHelper: null);
            var deserializer = new ProtoDeserializer(ms, context);
            object? result = deserializer.Deserialize();
            deserializer.Done();
            return result == null ? null : PSObject.AsPSObject(result);
        }

        internal static byte[] SerializeCborRemotingPayload(object? obj)
        {
            using var ms = new MemoryStream();
            var context = new SerializationContext(
                depth: 1,
                options: SerializationOptions.RemotingOptions,
                cryptoHelper: null);
            var serializer = new CborSerializer(ms, context);
            serializer.Serialize(obj!);
            serializer.Done();
            return ms.ToArray();
        }

        internal static object? DeserializeCborRemotingPayload(byte[] cborBytes)
        {
            using var ms = new MemoryStream(cborBytes, writable: false);
            var context = new DeserializationContext(
                options: DeserializationOptions.RemotingOptions,
                cryptoHelper: null);
            var deserializer = new CborDeserializer(ms, context);
            object? result = deserializer.Deserialize();
            deserializer.Done();
            return result == null ? null : PSObject.AsPSObject(result);
        }

        internal static object? DeserializeClixmlRemotingPayload(byte[] buffer, int offset, int count)
        {
            if (s_smaDeserializerCtor == null || s_smaDesCtxCtor == null || s_smaDeserializeMethod == null || s_smaDeserializerDoneMethod == null)
            {
                throw new InvalidOperationException("Failed to initialize System.Management.Automation.Deserializer reflection bindings.");
            }

            using var ms = new MemoryStream(buffer, offset, count, writable: false);
            var xmlSettings = new XmlReaderSettings
            {
                CheckCharacters = false,
                IgnoreComments = false,
                IgnoreProcessingInstructions = false,
                MaxCharactersInDocument = 0,
                XmlResolver = null,
                ConformanceLevel = ConformanceLevel.Fragment,
                CloseInput = false
            };

            using var xmlReader = XmlReader.Create(ms, xmlSettings);
            object desCtx = s_smaDesCtxCtor.Invoke(new object?[] { s_smaRemotingDesOptions, null });
            object deserializer = s_smaDeserializerCtor.Invoke(new object?[] { xmlReader, desCtx });
            object? result = s_smaDeserializeMethod.Invoke(deserializer, Array.Empty<object>());
            s_smaDeserializerDoneMethod.Invoke(deserializer, Array.Empty<object>());
            return result == null ? null : PSObject.AsPSObject(result);
        }

        internal static byte[] SerializeClixmlRemotingPayload(object? obj)
        {
            if (s_smaSerializerCtor == null || s_smaSerCtxCtor == null || s_smaSerializeMethod == null || s_smaSerializerDoneMethod == null)
            {
                throw new InvalidOperationException("Failed to initialize System.Management.Automation.Serializer reflection bindings.");
            }

            using var ms = new MemoryStream();
            var xmlSettings = new XmlWriterSettings
            {
                CheckCharacters = false,
                Indent = false,
                CloseOutput = false,
                Encoding = s_utf8NoBom,
                NewLineHandling = NewLineHandling.None,
                OmitXmlDeclaration = true,
                ConformanceLevel = ConformanceLevel.Fragment
            };

            using (var xmlWriter = XmlWriter.Create(ms, xmlSettings))
            {
                object serCtx = s_smaSerCtxCtor.Invoke(new object?[] { 1, s_smaRemotingSerOptions, null });
                object serializer = s_smaSerializerCtor.Invoke(new object?[] { xmlWriter, serCtx });
                s_smaSerializeMethod.Invoke(serializer, new object?[] { obj });
                s_smaSerializerDoneMethod.Invoke(serializer, Array.Empty<object>());
                xmlWriter.Flush();
            }

            return ms.ToArray();
        }
    }
}
