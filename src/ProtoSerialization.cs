// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

#pragma warning disable CS8600, CS8601, CS8602, CS8603, CS8604, CS8618, CS8625

using System;
using System.Buffers.Binary;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Security;
using System.Text;
using System.Xml;

using Dbg = System.Management.Automation.Diagnostics;

namespace System.Management.Automation
{
    /// <summary>
    /// Provides public functionality for serializing and deserializing PowerShell objects
    /// using Protocol Buffers (proto3) wire encoding.
    /// </summary>
    public static class PSProtoSerializer
    {
        private const int DefaultSerializationDepth = 1;

        /// <summary>
        /// Serializes an object into PowerShell Protobuf format with a default depth of 1.
        /// </summary>
        /// <param name="source">The input object to serialize.</param>
        /// <returns>The serialized object as a Protobuf byte array.</returns>
        public static byte[] Serialize(object? source)
        {
            return Serialize(source, DefaultSerializationDepth);
        }

        /// <summary>
        /// Serializes an object into PowerShell Protobuf format.
        /// </summary>
        /// <param name="source">The input object to serialize.</param>
        /// <param name="depth">The depth of the members to serialize.</param>
        /// <returns>The serialized object as a Protobuf byte array.</returns>
        public static byte[] Serialize(object? source, int depth)
        {
            using MemoryStream ms = new();
            Serialize(source, ms, depth);
            return ms.ToArray();
        }

        /// <summary>
        /// Serializes an object into PowerShell Protobuf format and writes it to the specified stream.
        /// </summary>
        /// <param name="source">The input object to serialize.</param>
        /// <param name="stream">The destination stream.</param>
        /// <param name="depth">The depth of the members to serialize.</param>
        public static void Serialize(object? source, Stream stream, int depth)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            ProtoSerializer serializer = new(stream, depth, useDepthFromTypes: true);
            serializer.Serialize(source);
            serializer.Done();
        }

        /// <summary>
        /// Serializes a list of objects into PowerShell Protobuf format.
        /// </summary>
        internal static byte[] Serialize(IList<object?> source, int depth, bool enumerate)
        {
            using MemoryStream ms = new();
            ProtoSerializer serializer = new(ms, depth, useDepthFromTypes: true);

            if (enumerate)
            {
                foreach (object? item in source)
                {
                    serializer.Serialize(item);
                }
            }
            else
            {
                serializer.Serialize(source);
            }

            serializer.Done();
            return ms.ToArray();
        }

        /// <summary>
        /// Serializes an object into a Base64-encoded PowerShell Protobuf string.
        /// </summary>
        public static string SerializeToBase64String(object? source, int depth = DefaultSerializationDepth)
        {
            return Convert.ToBase64String(Serialize(source, depth));
        }

        /// <summary>
        /// Deserializes PowerShell Protobuf bytes into an object.
        /// </summary>
        /// <param name="source">The Protobuf bytes representing the object to deserialize.</param>
        /// <returns>An object that represents the serialized content.</returns>
        public static object? Deserialize(byte[] source)
        {
            object?[] results = DeserializeAsList(source);

            if (results.Length == 0)
            {
                return null;
            }
            else if (results.Length == 1)
            {
                return results[0];
            }
            else
            {
                return results;
            }
        }

        /// <summary>
        /// Deserializes PowerShell Protobuf from a stream into an object.
        /// </summary>
        /// <param name="stream">The stream containing Protobuf data.</param>
        /// <returns>An object that represents the serialized content.</returns>
        public static object? Deserialize(Stream stream)
        {
            object?[] results = DeserializeAsList(stream);

            if (results.Length == 0)
            {
                return null;
            }
            else if (results.Length == 1)
            {
                return results[0];
            }
            else
            {
                return results;
            }
        }

        /// <summary>
        /// Deserializes PowerShell Protobuf bytes into a list of objects.
        /// </summary>
        public static object?[] DeserializeAsList(byte[] source)
        {
            if (source is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(source));
            }

            using MemoryStream ms = new(source, writable: false);
            return DeserializeAsList(ms);
        }

        /// <summary>
        /// Deserializes PowerShell Protobuf from a stream into a list of objects.
        /// </summary>
        public static object?[] DeserializeAsList(Stream stream)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            List<object?> results = new();
            ProtoDeserializer deserializer = new(stream);
            while (!deserializer.Done())
            {
                object? result = deserializer.Deserialize();
                results.Add(result);
            }

            return results.ToArray();
        }
    }

    /// <summary>
    /// Wire constants and field numbers for PowerShell Protobuf (proto3) serialization.
    /// </summary>
    internal static class ProtoSchema
    {
        // 4-byte stream magic header: "PSPB" (0x50, 0x53, 0x50, 0x42)
        internal static ReadOnlySpan<byte> StreamMagic => "PSPB"u8;

        // Protobuf wire types
        internal const int WireVarint = 0;
        internal const int WireFixed64 = 1;
        internal const int WireLengthDelimited = 2;
        internal const int WireFixed32 = 5;

        // PSValue fields
        internal const int Val_Name = 1;               // string (Len)
        internal const int Val_Null = 2;               // bool (Varint)
        internal const int Val_Bool = 3;               // bool (Varint)
        internal const int Val_String = 4;             // string (Len)
        internal const int Val_Int32 = 5;              // sint32 (ZigZag Varint)
        internal const int Val_Int64 = 6;              // sint64 (ZigZag Varint)
        internal const int Val_UInt64 = 7;             // uint64 (Varint)
        internal const int Val_Double = 8;             // double (Fixed64)
        internal const int Val_Bytes = 9;              // bytes (Len)
        internal const int Val_Guid = 10;              // bytes (Len, 16B)
        internal const int Val_DateTime = 11;          // fixed64 (Ticks | (Kind << 62))
        internal const int Val_TimeSpan = 12;          // sint64 (ZigZag Varint Ticks)
        internal const int Val_ObjectRef = 13;         // string RefId (Len)
        internal const int Val_Complex = 14;           // PSComplexObject (Len)
        internal const int Val_UInt32 = 15;            // uint32 (Varint)
        internal const int Val_Byte = 16;              // uint32 (Varint)
        internal const int Val_SByte = 17;             // sint32 (ZigZag Varint)
        internal const int Val_Int16 = 18;             // sint32 (ZigZag Varint)
        internal const int Val_UInt16 = 19;            // uint32 (Varint)
        internal const int Val_Char = 20;              // uint32 (Varint)
        internal const int Val_Single = 21;            // float (Fixed32)
        internal const int Val_Decimal = 22;           // bytes (Len, 16B)
        internal const int Val_Uri = 23;               // string (Len)
        internal const int Val_Version = 24;           // string (Len)
        internal const int Val_ScriptBlock = 25;       // string (Len)
        internal const int Val_XmlDocument = 26;       // string (Len)
        internal const int Val_ProgressRecord = 27;    // PSProgressRecord (Len)
        internal const int Val_SecureString = 28;      // string (Len)
        internal const int Val_StreamName = 29;        // string (Len)

        // PSComplexObject fields (all 1..12 => 1-byte tags!)
        internal const int Obj_RefId = 1;              // string (Len)
        internal const int Obj_TypeNames = 2;          // PSTypeNameList (Len)
        internal const int Obj_TypeNamesRef = 3;       // string (Len)
        internal const int Obj_ToString = 4;           // string (Len)
        internal const int Obj_BaseValue = 5;          // PSValue (Len)
        internal const int Obj_ContainerKind = 6;      // uint32 (Varint: ContainerType)
        internal const int Obj_CollectionItems = 7;    // repeated PSValue (Len)
        internal const int Obj_DictEntries = 8;        // repeated PSDictionaryEntry (Len)
        internal const int Obj_AdaptedProps = 9;       // repeated PSValue (Len, with Val_Name)
        internal const int Obj_NoteProperties = 10;    // repeated PSValue (Len, with Val_Name)
        internal const int Obj_MemberSets = 11;        // repeated PSMemberSet (Len)

        // PSTypeNameList fields
        internal const int TN_RefId = 1;               // string (Len)
        internal const int TN_Types = 2;               // repeated string (Len)

        // PSDictionaryEntry fields
        internal const int Dict_Key = 1;               // PSValue (Len)
        internal const int Dict_Value = 2;             // PSValue (Len)

        // PSMemberSet fields
        internal const int MS_Name = 1;                // string (Len)
        internal const int MS_NoteProperties = 2;      // repeated PSValue (Len, with Val_Name)
        internal const int MS_NestedSets = 3;          // repeated PSMemberSet (Len)

        // PSProgressRecord fields
        internal const int PR_Activity = 1;            // string (Len)
        internal const int PR_ActivityId = 2;          // sint32 (ZigZag Varint)
        internal const int PR_CurrentOperation = 3;    // PSValue (Len)
        internal const int PR_ParentActivityId = 4;    // sint32 (ZigZag Varint)
        internal const int PR_PercentComplete = 5;     // sint32 (ZigZag Varint)
        internal const int PR_RecordType = 6;          // uint32 (Varint)
        internal const int PR_SecondsRemaining = 7;    // sint32 (ZigZag Varint)
        internal const int PR_StatusDescription = 8;   // string (Len)
    }

    /// <summary>
    /// Provides functionality for serializing a PSObject into a Protobuf stream.
    /// </summary>
    internal sealed class ProtoSerializer
    {
        private readonly InternalProtoSerializer _serializer;

        internal ProtoSerializer(Stream stream)
            : this(stream, new SerializationContext())
        {
        }

        internal ProtoSerializer(Stream stream, int depth, bool useDepthFromTypes)
            : this(stream, new SerializationContext(depth, useDepthFromTypes))
        {
        }

        internal ProtoSerializer(Stream stream, SerializationContext context)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            if (context is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(context));
            }

            _serializer = new InternalProtoSerializer(stream, context);
            _serializer.Start();
        }

        internal TypeTable? TypeTable
        {
            get => _serializer.TypeTable;
            set => _serializer.TypeTable = value;
        }

        internal void Serialize(object? source)
        {
            Serialize(source, null);
        }

        internal void Serialize(object? source, string? streamName)
        {
            _serializer.WriteOneTopLevelObject(source, streamName);
        }

        internal void Done()
        {
            _serializer.End();
        }

        internal void Stop()
        {
            _serializer.Stop();
        }
    }

    /// <summary>
    /// Provides functionality for deserializing a PSObject from a Protobuf stream.
    /// </summary>
    internal sealed class ProtoDeserializer
    {
        private readonly InternalProtoDeserializer _deserializer;

        internal ProtoDeserializer(Stream stream)
            : this(stream, new DeserializationContext())
        {
        }

        internal ProtoDeserializer(Stream stream, DeserializationContext context)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            if (context is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(context));
            }

            _deserializer = new InternalProtoDeserializer(stream, context);
            _deserializer.Start();
        }

        internal TypeTable? TypeTable
        {
            get => _deserializer.TypeTable;
            set => _deserializer.TypeTable = value;
        }

        internal bool Done()
        {
            return _deserializer.Done();
        }

        internal void Stop()
        {
            _deserializer.Stop();
        }

        internal object? Deserialize()
        {
            return Deserialize(out _);
        }

        internal object? Deserialize(out string? streamName)
        {
            if (Done())
            {
                throw PSTraceSource.NewInvalidOperationException(Serialization.ReadCalledAfterDone);
            }

            return _deserializer.ReadOneTopLevelObject(out streamName);
        }

        /// <summary>
        /// Skips the next top-level length-delimited Protobuf object in O(1) without decoding its fields.
        /// </summary>
        internal bool TrySkipTopLevelObject()
        {
            if (Done())
            {
                return false;
            }

            return _deserializer.TrySkipOneTopLevelObject();
        }
    }

    /// <summary>
    /// Single-buffer in-memory Protocol Buffers v3 wire writer supporting zero-allocation nested submessages.
    /// </summary>
    internal sealed class ProtoBufferWriter
    {
        private byte[] _buffer;
        private int _pos;

        internal ProtoBufferWriter(int initialCapacity = 1024)
        {
            _buffer = new byte[initialCapacity];
            _pos = 0;
        }

        internal int Length => _pos;

        internal ReadOnlySpan<byte> WrittenSpan => _buffer.AsSpan(0, _pos);

        internal void Reset()
        {
            _pos = 0;
        }

        internal void WriteToStreamDelimited(Stream stream)
        {
            Span<byte> lenBuf = stackalloc byte[5];
            int lenBytes = EncodeVarint32((uint)_pos, lenBuf);
            stream.Write(lenBuf.Slice(0, lenBytes));
            if (_pos > 0)
            {
                stream.Write(_buffer, 0, _pos);
            }
        }

        internal void WriteTag(int fieldNumber, int wireType)
        {
            WriteVarint32((uint)((fieldNumber << 3) | wireType));
        }

        internal void WriteVarint32(uint value)
        {
            EnsureCapacity(5);
            while (value >= 0x80u)
            {
                _buffer[_pos++] = (byte)(value | 0x80u);
                value >>= 7;
            }

            _buffer[_pos++] = (byte)value;
        }

        internal void WriteVarint64(ulong value)
        {
            EnsureCapacity(10);
            while (value >= 0x80UL)
            {
                _buffer[_pos++] = (byte)(value | 0x80UL);
                value >>= 7;
            }

            _buffer[_pos++] = (byte)value;
        }

        internal void WriteZigZag32(int value)
        {
            uint zz = (uint)((value << 1) ^ (value >> 31));
            WriteVarint32(zz);
        }

        internal void WriteZigZag64(long value)
        {
            ulong zz = (ulong)((value << 1) ^ (value >> 63));
            WriteVarint64(zz);
        }

        internal void WriteBoolField(int fieldNumber, bool value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireVarint);
            EnsureCapacity(1);
            _buffer[_pos++] = value ? (byte)1 : (byte)0;
        }

        internal void WriteUInt32Field(int fieldNumber, uint value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireVarint);
            WriteVarint32(value);
        }

        internal void WriteUInt64Field(int fieldNumber, ulong value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireVarint);
            WriteVarint64(value);
        }

        internal void WriteSInt32Field(int fieldNumber, int value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireVarint);
            WriteZigZag32(value);
        }

        internal void WriteSInt64Field(int fieldNumber, long value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireVarint);
            WriteZigZag64(value);
        }

        internal void WriteFixed32Field(int fieldNumber, float value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireFixed32);
            EnsureCapacity(4);
            BinaryPrimitives.WriteSingleLittleEndian(_buffer.AsSpan(_pos, 4), value);
            _pos += 4;
        }

        internal void WriteFixed64Field(int fieldNumber, double value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireFixed64);
            EnsureCapacity(8);
            BinaryPrimitives.WriteDoubleLittleEndian(_buffer.AsSpan(_pos, 8), value);
            _pos += 8;
        }

        internal void WriteFixed64BitsField(int fieldNumber, ulong value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireFixed64);
            EnsureCapacity(8);
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.AsSpan(_pos, 8), value);
            _pos += 8;
        }

        internal void WriteBytesField(int fieldNumber, ReadOnlySpan<byte> bytes)
        {
            WriteTag(fieldNumber, ProtoSchema.WireLengthDelimited);
            WriteVarint32((uint)bytes.Length);
            EnsureCapacity(bytes.Length);
            bytes.CopyTo(_buffer.AsSpan(_pos, bytes.Length));
            _pos += bytes.Length;
        }

        internal void WriteStringField(int fieldNumber, string value)
        {
            WriteTag(fieldNumber, ProtoSchema.WireLengthDelimited);
            int byteCount = Encoding.UTF8.GetByteCount(value);
            WriteVarint32((uint)byteCount);
            EnsureCapacity(byteCount);
            Encoding.UTF8.GetBytes(value.AsSpan(), _buffer.AsSpan(_pos, byteCount));
            _pos += byteCount;
        }

        /// <summary>
        /// Begins a length-delimited submessage for <paramref name="fieldNumber"/>, reserving 1 byte for the length varint.
        /// </summary>
        internal int BeginSubMessage(int fieldNumber)
        {
            WriteTag(fieldNumber, ProtoSchema.WireLengthDelimited);
            EnsureCapacity(1);
            int lengthToken = _pos;
            _pos++;
            return lengthToken;
        }

        /// <summary>
        /// Completes a submessage started by <see cref="BeginSubMessage"/>, backpatching its varint byte length.
        /// </summary>
        internal void EndSubMessage(int lengthToken)
        {
            int payloadLength = _pos - (lengthToken + 1);
            if (payloadLength < 0x80)
            {
                _buffer[lengthToken] = (byte)payloadLength;
                return;
            }

            Span<byte> varintBuf = stackalloc byte[5];
            int varintSize = EncodeVarint32((uint)payloadLength, varintBuf);
            int extraBytes = varintSize - 1;
            EnsureCapacity(extraBytes);

            Buffer.BlockCopy(_buffer, lengthToken + 1, _buffer, lengthToken + varintSize, payloadLength);
            varintBuf.Slice(0, varintSize).CopyTo(_buffer.AsSpan(lengthToken, varintSize));
            _pos += extraBytes;
        }

        internal static int EncodeVarint32(uint value, Span<byte> destination)
        {
            int idx = 0;
            while (value >= 0x80u)
            {
                destination[idx++] = (byte)(value | 0x80u);
                value >>= 7;
            }

            destination[idx++] = (byte)value;
            return idx;
        }

        private void EnsureCapacity(int additionalBytes)
        {
            int required = _pos + additionalBytes;
            if (required <= _buffer.Length)
            {
                return;
            }

            int newSize = Math.Max(_buffer.Length * 2, required + 256);
            Array.Resize(ref _buffer, newSize);
        }
    }

    /// <summary>
    /// Zero-copy span-based Protocol Buffers v3 wire reader.
    /// </summary>
    internal ref struct ProtoSpanReader
    {
        private readonly ReadOnlySpan<byte> _buffer;
        private int _pos;

        internal ProtoSpanReader(ReadOnlySpan<byte> buffer)
        {
            _buffer = buffer;
            _pos = 0;
        }

        internal bool IsEOF => _pos >= _buffer.Length;

        internal bool TryReadTag(out int fieldNumber, out int wireType)
        {
            if (_pos >= _buffer.Length)
            {
                fieldNumber = 0;
                wireType = 0;
                return false;
            }

            uint tag = ReadVarint32();
            fieldNumber = (int)(tag >> 3);
            wireType = (int)(tag & 0x07u);
            return fieldNumber > 0;
        }

        internal uint ReadVarint32()
        {
            uint result = 0;
            int shift = 0;
            while (shift < 35)
            {
                if (_pos >= _buffer.Length)
                {
                    throw new FormatException("Unexpected EOF while reading Protobuf varint32.");
                }

                byte b = _buffer[_pos++];
                result |= (uint)(b & 0x7F) << shift;
                if ((b & 0x80) == 0)
                {
                    return result;
                }

                shift += 7;
            }

            throw new FormatException("Malformed Protobuf varint32.");
        }

        internal ulong ReadVarint64()
        {
            ulong result = 0;
            int shift = 0;
            while (shift < 70)
            {
                if (_pos >= _buffer.Length)
                {
                    throw new FormatException("Unexpected EOF while reading Protobuf varint64.");
                }

                byte b = _buffer[_pos++];
                result |= (ulong)(b & 0x7F) << shift;
                if ((b & 0x80) == 0)
                {
                    return result;
                }

                shift += 7;
            }

            throw new FormatException("Malformed Protobuf varint64.");
        }

        internal int ReadSInt32()
        {
            uint n = ReadVarint32();
            return (int)((n >> 1) ^ (uint)-(int)(n & 1));
        }

        internal long ReadSInt64()
        {
            ulong n = ReadVarint64();
            return (long)((n >> 1) ^ (ulong)-(long)(n & 1));
        }

        internal bool ReadBool()
        {
            return ReadVarint32() != 0;
        }

        internal float ReadFixed32Single()
        {
            if (_pos + 4 > _buffer.Length)
            {
                throw new FormatException("Unexpected EOF while reading Protobuf fixed32.");
            }

            float val = BinaryPrimitives.ReadSingleLittleEndian(_buffer.Slice(_pos, 4));
            _pos += 4;
            return val;
        }

        internal double ReadFixed64Double()
        {
            if (_pos + 8 > _buffer.Length)
            {
                throw new FormatException("Unexpected EOF while reading Protobuf fixed64.");
            }

            double val = BinaryPrimitives.ReadDoubleLittleEndian(_buffer.Slice(_pos, 8));
            _pos += 8;
            return val;
        }

        internal ulong ReadFixed64Bits()
        {
            if (_pos + 8 > _buffer.Length)
            {
                throw new FormatException("Unexpected EOF while reading Protobuf fixed64.");
            }

            ulong val = BinaryPrimitives.ReadUInt64LittleEndian(_buffer.Slice(_pos, 8));
            _pos += 8;
            return val;
        }

        internal ReadOnlySpan<byte> ReadLengthDelimitedSpan()
        {
            int len = checked((int)ReadVarint32());
            if (len < 0 || _pos + len > _buffer.Length)
            {
                throw new FormatException("Unexpected EOF while reading Protobuf length-delimited field.");
            }

            ReadOnlySpan<byte> slice = _buffer.Slice(_pos, len);
            _pos += len;
            return slice;
        }

        internal byte[] ReadBytes()
        {
            return ReadLengthDelimitedSpan().ToArray();
        }

        internal string ReadString()
        {
            ReadOnlySpan<byte> span = ReadLengthDelimitedSpan();
            return span.IsEmpty ? string.Empty : Encoding.UTF8.GetString(span);
        }

        internal void SkipField(int wireType)
        {
            switch (wireType)
            {
                case ProtoSchema.WireVarint:
                    _ = ReadVarint64();
                    break;
                case ProtoSchema.WireFixed64:
                    if (_pos + 8 > _buffer.Length)
                    {
                        throw new FormatException("Unexpected EOF skipping fixed64.");
                    }

                    _pos += 8;
                    break;
                case ProtoSchema.WireLengthDelimited:
                    _ = ReadLengthDelimitedSpan();
                    break;
                case ProtoSchema.WireFixed32:
                    if (_pos + 4 > _buffer.Length)
                    {
                        throw new FormatException("Unexpected EOF skipping fixed32.");
                    }

                    _pos += 4;
                    break;
                default:
                    throw new FormatException($"Unsupported Protobuf wire type: {wireType}");
            }
        }
    }

    /// <summary>
    /// Internal Protobuf serializer mirroring <see cref="InternalSerializer"/>.
    /// </summary>
    internal sealed class InternalProtoSerializer
    {
        private const int MaxDepthBelowTopLevel = 50;

        private readonly Stream _stream;
        private readonly ProtoBufferWriter _writer = new(1024);
        private readonly SerializationContext _context;
        private readonly ReferenceIdHandlerForSerializer<object> _objectRefIdHandler;
        private readonly ReferenceIdHandlerForSerializer<ConsolidatedString> _typeRefIdHandler;

        private TypeTable? _typeTable;
        private int _depthBelowTopLevel;
        private bool _isStopping;
        private bool? _canUseDefaultRunspaceInThreadSafeManner;
        private Collection<CollectionEntry<PSMemberInfo>>? _extendedMembersCollection;
        private Collection<CollectionEntry<PSPropertyInfo>>? _allPropertiesCollection;

        internal InternalProtoSerializer(Stream stream, SerializationContext context)
        {
            Dbg.Assert(stream is not null, "caller should validate the parameter");
            Dbg.Assert(context is not null, "caller should validate the parameter");
            _stream = stream;
            _context = context;

            IDictionary<object, ulong>? objectRefIdDictionary = null;
            if ((_context.options & SerializationOptions.NoObjectRefIds) == 0)
            {
                objectRefIdDictionary = new WeakReferenceDictionary<ulong>();
            }

            _objectRefIdHandler = new ReferenceIdHandlerForSerializer<object>(objectRefIdDictionary);
            _typeRefIdHandler = new ReferenceIdHandlerForSerializer<ConsolidatedString>(
                new Dictionary<ConsolidatedString, ulong>(ConsolidatedString.EqualityComparer));
        }

        internal TypeTable? TypeTable
        {
            get => _typeTable;
            set => _typeTable = value;
        }

        private bool CanUseDefaultRunspaceInThreadSafeManner
        {
            get
            {
                _canUseDefaultRunspaceInThreadSafeManner ??= Runspace.CanUseDefaultRunspace;
                return _canUseDefaultRunspaceInThreadSafeManner.Value;
            }
        }

        private Collection<CollectionEntry<PSMemberInfo>> ExtendedMembersCollection
        {
            get => _extendedMembersCollection ??= PSObject.GetMemberCollection(PSMemberViewTypes.Extended, _typeTable);
        }

        private Collection<CollectionEntry<PSPropertyInfo>> AllPropertiesCollection
        {
            get => _allPropertiesCollection ??= PSObject.GetPropertyCollection(PSMemberViewTypes.All, _typeTable);
        }

        internal void Start()
        {
            if ((_context.options & SerializationOptions.NoRootElement) != SerializationOptions.NoRootElement)
            {
                _stream.Write(ProtoSchema.StreamMagic);
            }
        }

        internal void End()
        {
            _stream.Flush();
        }

        internal void Stop()
        {
            _isStopping = true;
        }

        private void CheckIfStopping()
        {
            if (_isStopping)
            {
                throw PSTraceSource.NewInvalidOperationException(Serialization.Stopping);
            }
        }

        internal void WriteOneTopLevelObject(object? source, string? streamName)
        {
            Dbg.Assert(_depthBelowTopLevel == 0, "InternalProtoSerializer.depthBelowTopLevel should be 0 at top-level");
            _writer.Reset();
            WritePSValueBody(source, propertyName: null, streamName: streamName, _context.depth);
            _writer.WriteToStreamDelimited(_stream);
        }

        private void WriteNestedPSValueField(int fieldNumber, object? source, string? propertyName, int depth)
        {
            int token = _writer.BeginSubMessage(fieldNumber);
            WritePSValueBody(source, propertyName, streamName: null, depth);
            _writer.EndSubMessage(token);
        }

        private void WritePSValueBody(object? source, string? propertyName, string? streamName, int depth)
        {
            Dbg.Assert(depth >= 0, "depth should always be greater or equal to zero");
            CheckIfStopping();

            if (!string.IsNullOrEmpty(propertyName))
            {
                _writer.WriteStringField(ProtoSchema.Val_Name, propertyName);
            }

            if (streamName is not null)
            {
                _writer.WriteStringField(ProtoSchema.Val_StreamName, streamName);
            }

            if (source is null)
            {
                _writer.WriteBoolField(ProtoSchema.Val_Null, true);
                return;
            }

            try
            {
                _depthBelowTopLevel++;
                if (_depthBelowTopLevel == MaxDepthBelowTopLevel)
                {
                    _writer.WriteStringField(ProtoSchema.Val_String, Serialization.DeserializationTooDeep);
                    return;
                }

                depth = GetDepthOfSerialization(source, depth);

                if (HandlePrimitiveKnownTypeByConvertingToPSObject(source, depth))
                {
                    return;
                }

                string? refId = _objectRefIdHandler.GetRefId(source);
                if (refId is not null)
                {
                    _writer.WriteStringField(ProtoSchema.Val_ObjectRef, refId);
                    return;
                }

                if (HandlePrimitiveKnownTypePSObject(source, depth))
                {
                    return;
                }

                if (HandleKnownContainerTypes(source, depth))
                {
                    return;
                }

                PSObject mshSource = PSObject.AsPSObject(source);
                if (depth == 0 || SerializeAsString(mshSource))
                {
                    HandlePSObjectAsString(mshSource, depth);
                    return;
                }

                HandleComplexTypePSObject(source, depth);
            }
            finally
            {
                _depthBelowTopLevel--;
            }
        }

        private bool HandlePrimitiveKnownTypeByConvertingToPSObject(object source, int depth)
        {
            if (IsKnownProtoPrimitive(source.GetType()))
            {
                PSObject pktInfoPSObject = PSObject.AsPSObject(source);
                return HandlePrimitiveKnownTypePSObject(pktInfoPSObject, depth);
            }

            return false;
        }

        private bool HandlePrimitiveKnownTypePSObject(object source, int depth)
        {
            if (source is PSObject moSource && !moSource.ImmediateBaseObjectIsEmpty)
            {
                object baseObject = moSource.ImmediateBaseObject;
                if (IsKnownProtoPrimitive(baseObject.GetType()))
                {
                    WritePrimitiveTypePSObject(moSource, baseObject, depth);
                    return true;
                }
            }

            return false;
        }

        private static bool IsKnownProtoPrimitive(Type type)
        {
            return KnownTypes.GetTypeSerializationInfo(type) is not null;
        }

        private void WritePrimitiveTypePSObject(PSObject source, object? primitive, int depth)
        {
            string? toStringValue = SerializationUtilities.GetToStringForPrimitiveObject(source);
            bool hasModifiedTypesCollection = PSObjectHasModifiedTypesCollection(source);
            bool hasNotes = PSObjectHasNotes(source);
            bool hasModifiedToString = toStringValue is not null;

            if (hasNotes || hasModifiedTypesCollection || hasModifiedToString)
            {
                string? refId = _objectRefIdHandler.SetRefId(source);
                int objToken = _writer.BeginSubMessage(ProtoSchema.Val_Complex);

                WritePSObjectHeaderEntries(
                    source,
                    refId,
                    writeTypeNames: hasModifiedTypesCollection,
                    toStringValue: toStringValue);

                if (primitive is not null)
                {
                    int baseToken = _writer.BeginSubMessage(ProtoSchema.Obj_BaseValue);
                    WriteOnePrimitiveKnownType(primitive);
                    _writer.EndSubMessage(baseToken);
                }

                SerializeInstanceProperties(source, depth);

                _writer.EndSubMessage(objToken);
            }
            else if (primitive is not null)
            {
                WriteOnePrimitiveKnownType(primitive);
            }
            else
            {
                _writer.WriteBoolField(ProtoSchema.Val_Null, true);
            }
        }

        private void WriteOnePrimitiveKnownType(object source)
        {
            switch (source)
            {
                case bool b:
                    _writer.WriteBoolField(ProtoSchema.Val_Bool, b);
                    break;
                case int i32:
                    _writer.WriteSInt32Field(ProtoSchema.Val_Int32, i32);
                    break;
                case long i64:
                    _writer.WriteSInt64Field(ProtoSchema.Val_Int64, i64);
                    break;
                case ulong u64:
                    _writer.WriteUInt64Field(ProtoSchema.Val_UInt64, u64);
                    break;
                case byte u8:
                    _writer.WriteUInt32Field(ProtoSchema.Val_Byte, u8);
                    break;
                case sbyte i8:
                    _writer.WriteSInt32Field(ProtoSchema.Val_SByte, i8);
                    break;
                case short i16:
                    _writer.WriteSInt32Field(ProtoSchema.Val_Int16, i16);
                    break;
                case ushort u16:
                    _writer.WriteUInt32Field(ProtoSchema.Val_UInt16, u16);
                    break;
                case uint u32:
                    _writer.WriteUInt32Field(ProtoSchema.Val_UInt32, u32);
                    break;
                case float f32:
                    _writer.WriteFixed32Field(ProtoSchema.Val_Single, f32);
                    break;
                case double f64:
                    _writer.WriteFixed64Field(ProtoSchema.Val_Double, f64);
                    break;
                case decimal dec:
                    {
                        Span<int> bits = stackalloc int[4];
                        decimal.GetBits(dec, bits);
                        Span<byte> decBytes = stackalloc byte[16];
                        for (int i = 0; i < 4; i++)
                        {
                            BinaryPrimitives.WriteInt32LittleEndian(decBytes.Slice(i * 4, 4), bits[i]);
                        }

                        _writer.WriteBytesField(ProtoSchema.Val_Decimal, decBytes);
                        break;
                    }
                case char ch:
                    _writer.WriteUInt32Field(ProtoSchema.Val_Char, ch);
                    break;
                case string s:
                    _writer.WriteStringField(ProtoSchema.Val_String, s);
                    break;
                case byte[] bytes:
                    _writer.WriteBytesField(ProtoSchema.Val_Bytes, bytes);
                    break;
                case Guid guid:
                    {
                        Span<byte> guidBytes = stackalloc byte[16];
                        guid.TryWriteBytes(guidBytes, bigEndian: true, out _);
                        _writer.WriteBytesField(ProtoSchema.Val_Guid, guidBytes);
                        break;
                    }
                case DateTime dt:
                    {
                        ulong packed = ((ulong)dt.Ticks & 0x3FFFFFFFFFFFFFFFUL) | ((ulong)dt.Kind << 62);
                        _writer.WriteFixed64BitsField(ProtoSchema.Val_DateTime, packed);
                        break;
                    }
                case TimeSpan ts:
                    _writer.WriteSInt64Field(ProtoSchema.Val_TimeSpan, ts.Ticks);
                    break;
                case Uri uri:
                    _writer.WriteStringField(ProtoSchema.Val_Uri, uri.OriginalString);
                    break;
                case Version ver:
                    _writer.WriteStringField(ProtoSchema.Val_Version, ver.ToString());
                    break;
                case ScriptBlock sb:
                    _writer.WriteStringField(ProtoSchema.Val_ScriptBlock, sb.ToString());
                    break;
                case XmlDocument xmlDoc:
                    _writer.WriteStringField(ProtoSchema.Val_XmlDocument, xmlDoc.OuterXml);
                    break;
                case ProgressRecord pr:
                    {
                        int prToken = _writer.BeginSubMessage(ProtoSchema.Val_ProgressRecord);
                        _writer.WriteStringField(ProtoSchema.PR_Activity, pr.Activity ?? string.Empty);
                        _writer.WriteSInt32Field(ProtoSchema.PR_ActivityId, pr.ActivityId);
                        WriteNestedPSValueField(ProtoSchema.PR_CurrentOperation, pr.CurrentOperation, propertyName: null, depth: 1);
                        _writer.WriteSInt32Field(ProtoSchema.PR_ParentActivityId, pr.ParentActivityId);
                        _writer.WriteSInt32Field(ProtoSchema.PR_PercentComplete, pr.PercentComplete);
                        _writer.WriteUInt32Field(ProtoSchema.PR_RecordType, (uint)pr.RecordType);
                        _writer.WriteSInt32Field(ProtoSchema.PR_SecondsRemaining, pr.SecondsRemaining);
                        _writer.WriteStringField(ProtoSchema.PR_StatusDescription, pr.StatusDescription ?? string.Empty);
                        _writer.EndSubMessage(prToken);
                        break;
                    }
                case SecureString ss:
                    {
                        string encryptedString = _context.cryptoHelper is not null
                            ? _context.cryptoHelper.EncryptSecureString(ss)
                            : Microsoft.PowerShell.SecureStringHelper.Protect(ss);
                        _writer.WriteStringField(ProtoSchema.Val_SecureString, encryptedString);
                        break;
                    }
                default:
                    _writer.WriteStringField(ProtoSchema.Val_String, Convert.ToString(source, CultureInfo.InvariantCulture) ?? string.Empty);
                    break;
            }
        }

        private bool HandleKnownContainerTypes(object source, int depth)
        {
            PSObject? mshSource = source as PSObject;
            if (mshSource is not null && mshSource.ImmediateBaseObjectIsEmpty)
            {
                return false;
            }

            SerializationUtilities.GetKnownContainerTypeInfo(
                mshSource is not null ? mshSource.ImmediateBaseObject : source,
                out ContainerType ct,
                out IDictionary dictionary,
                out IEnumerable enumerable);

            if (ct == ContainerType.None)
            {
                return false;
            }

            string? refId = _objectRefIdHandler.SetRefId(source);
            int objToken = _writer.BeginSubMessage(ProtoSchema.Val_Complex);

            WritePSObjectHeaderEntries(
                mshSource ?? PSObject.AsPSObject(source),
                refId,
                writeTypeNames: true,
                toStringValue: null);

            _writer.WriteUInt32Field(ProtoSchema.Obj_ContainerKind, (uint)ct);

            if (ct == ContainerType.Dictionary)
            {
                WriteDictionary(dictionary, depth);
            }
            else
            {
                WriteEnumerable(enumerable, depth);
            }

            if (depth != 0)
            {
                if (ct == ContainerType.Enumerable || (mshSource is not null && mshSource.IsDeserialized))
                {
                    PSObject sourceAsPSObject = PSObject.AsPSObject(source);
                    PSMemberInfoInternalCollection<PSPropertyInfo> specificPropertiesToSerialize =
                        SerializationUtilities.GetSpecificPropertiesToSerialize(sourceAsPSObject, AllPropertiesCollection, _typeTable);
                    WritePSObjectProperties(sourceAsPSObject, depth, specificPropertiesToSerialize);
                    SerializeExtendedProperties(sourceAsPSObject, depth, specificPropertiesToSerialize);
                }
                else if (mshSource is not null)
                {
                    SerializeInstanceProperties(mshSource, depth);
                }
            }

            _writer.EndSubMessage(objToken);
            return true;
        }

        private void WriteEnumerable(IEnumerable enumerable, int depth)
        {
            IEnumerator? enumerator = null;
            try
            {
                enumerator = enumerable.GetEnumerator();
                try
                {
                    enumerator.Reset();
                }
                catch (NotSupportedException)
                {
                }
            }
            catch (Exception)
            {
                enumerator = null;
            }

            if (enumerator is not null)
            {
                while (true)
                {
                    object? item;
                    try
                    {
                        if (!enumerator.MoveNext())
                        {
                            break;
                        }

                        item = enumerator.Current;
                    }
                    catch (Exception)
                    {
                        break;
                    }

                    WriteNestedPSValueField(ProtoSchema.Obj_CollectionItems, item, propertyName: null, depth);
                }
            }
        }

        private void WriteDictionary(IDictionary dictionary, int depth)
        {
            IDictionaryEnumerator? dictionaryEnum = null;
            try
            {
                dictionaryEnum = dictionary.GetEnumerator();
            }
            catch (Exception)
            {
            }

            if (dictionaryEnum is not null)
            {
                while (true)
                {
                    object? key;
                    object? value;
                    try
                    {
                        if (!dictionaryEnum.MoveNext())
                        {
                            break;
                        }

                        key = dictionaryEnum.Key;
                        value = dictionaryEnum.Value;
                    }
                    catch (Exception)
                    {
                        break;
                    }

                    if (key is null)
                    {
                        break;
                    }

                    int entryToken = _writer.BeginSubMessage(ProtoSchema.Obj_DictEntries);
                    WriteNestedPSValueField(ProtoSchema.Dict_Key, key, propertyName: null, depth);
                    WriteNestedPSValueField(ProtoSchema.Dict_Value, value, propertyName: null, depth);
                    _writer.EndSubMessage(entryToken);
                }
            }
        }

        private void HandleComplexTypePSObject(object source, int depth)
        {
            PSObject mshSource = PSObject.AsPSObject(source);

            bool isErrorRecord = false;
            bool isInformationalRecord = false;
            bool isEnum = false;
            bool isPSObject = false;

            if (!mshSource.ImmediateBaseObjectIsEmpty)
            {
                if (mshSource.ImmediateBaseObject is ErrorRecord errorRecord)
                {
                    errorRecord.ToPSObjectForRemoting(mshSource);
                    isErrorRecord = true;
                }
                else if (mshSource.ImmediateBaseObject is InformationalRecord informationalRecord)
                {
                    informationalRecord.ToPSObjectForRemoting(mshSource);
                    isInformationalRecord = true;
                }
                else
                {
                    isEnum = mshSource.ImmediateBaseObject is Enum;
                    isPSObject = mshSource.ImmediateBaseObject is PSObject;
                }
            }

            bool writeToString = true;
            if (mshSource.ToStringFromDeserialization is null && mshSource.ImmediateBaseObjectIsEmpty)
            {
                writeToString = false;
            }

            string? refId = _objectRefIdHandler.SetRefId(source);
            int objToken = _writer.BeginSubMessage(ProtoSchema.Val_Complex);

            WritePSObjectHeaderEntries(
                mshSource,
                refId,
                writeTypeNames: true,
                toStringValue: writeToString ? SerializationUtilities.GetToString(mshSource) : null);

            PSMemberInfoInternalCollection<PSPropertyInfo> specificPropertiesToSerialize =
                SerializationUtilities.GetSpecificPropertiesToSerialize(mshSource, AllPropertiesCollection, _typeTable);

            if (isEnum)
            {
                object baseObject = mshSource.ImmediateBaseObject;
                object underlying = Convert.ChangeType(baseObject, Enum.GetUnderlyingType(baseObject.GetType()), CultureInfo.InvariantCulture);
                WriteNestedPSValueField(ProtoSchema.Obj_BaseValue, underlying, propertyName: null, depth);
            }
            else if (isPSObject)
            {
                WriteNestedPSValueField(ProtoSchema.Obj_BaseValue, mshSource.ImmediateBaseObject, propertyName: null, depth);
            }
            else if (!isErrorRecord && !isInformationalRecord)
            {
                WritePSObjectProperties(mshSource, depth, specificPropertiesToSerialize);
            }

            SerializeExtendedProperties(mshSource, depth, specificPropertiesToSerialize);

            _writer.EndSubMessage(objToken);
        }

        private void WritePSObjectHeaderEntries(
            PSObject mshObject,
            string? refId,
            bool writeTypeNames,
            string? toStringValue)
        {
            if (refId is not null)
            {
                _writer.WriteStringField(ProtoSchema.Obj_RefId, refId);
            }

            if (writeTypeNames)
            {
                ConsolidatedString typeNames = mshObject.InternalTypeNames;
                if (typeNames.Count > 0)
                {
                    string? typeNameHierarchyReferenceId = _typeRefIdHandler.GetRefId(typeNames);
                    if (typeNameHierarchyReferenceId is null)
                    {
                        string? tnRefId = _typeRefIdHandler.SetRefId(typeNames);
                        int tnToken = _writer.BeginSubMessage(ProtoSchema.Obj_TypeNames);
                        if (tnRefId is not null)
                        {
                            _writer.WriteStringField(ProtoSchema.TN_RefId, tnRefId);
                        }

                        foreach (string type in typeNames)
                        {
                            _writer.WriteStringField(ProtoSchema.TN_Types, type);
                        }

                        _writer.EndSubMessage(tnToken);
                    }
                    else
                    {
                        _writer.WriteStringField(ProtoSchema.Obj_TypeNamesRef, typeNameHierarchyReferenceId);
                    }
                }
            }

            if (toStringValue is not null)
            {
                _writer.WriteStringField(ProtoSchema.Obj_ToString, toStringValue);
            }
        }

        private void WritePSObjectProperties(
            PSObject source,
            int depth,
            IEnumerable<PSPropertyInfo>? specificPropertiesToSerialize)
        {
            --depth;

            if (specificPropertiesToSerialize is not null)
            {
                SerializeProperties(specificPropertiesToSerialize, depth);
            }
            else if (source.ShouldSerializeAdapter())
            {
                IEnumerable<PSPropertyInfo>? adapterCollection = source.GetAdaptedProperties();
                if (adapterCollection is not null)
                {
                    SerializeProperties(adapterCollection, depth);
                }
            }
        }

        private void SerializeProperties(IEnumerable<PSPropertyInfo> propertyCollection, int depth)
        {
            foreach (PSMemberInfo info in propertyCollection)
            {
                if (info is not PSProperty prop)
                {
                    continue;
                }

                object? value = SerializationUtilities.GetPropertyValueInThreadSafeManner(
                    prop,
                    CanUseDefaultRunspaceInThreadSafeManner,
                    out bool success);
                if (success)
                {
                    WriteNestedPSValueField(ProtoSchema.Obj_AdaptedProps, value, prop.Name, depth);
                }
            }
        }

        private void SerializeInstanceProperties(PSObject source, int depth)
        {
            PSMemberInfoCollection<PSMemberInfo>? instanceMembers = source.InstanceMembers;
            if (instanceMembers is not null)
            {
                WriteMemberInfoCollection(instanceMembers, depth, isTopLevelObjectMemberSet: true);
            }
        }

        private void SerializeExtendedProperties(
            PSObject source,
            int depth,
            IEnumerable<PSPropertyInfo>? specificPropertiesToSerialize)
        {
            IEnumerable<PSMemberInfo>? extendedMembersEnumerable;
            if (specificPropertiesToSerialize is null)
            {
                PSMemberInfoIntegratingCollection<PSMemberInfo> membersToSearch =
                    new(source, ExtendedMembersCollection);
                extendedMembersEnumerable = membersToSearch.Match(
                    "*",
                    PSMemberTypes.Properties | PSMemberTypes.PropertySet | PSMemberTypes.MemberSet,
                    MshMemberMatchOptions.IncludeHidden | MshMemberMatchOptions.OnlySerializable);
            }
            else
            {
                List<PSMemberInfo> extendedMembersList = new(source.InstanceMembers);
                extendedMembersEnumerable = extendedMembersList;

                foreach (PSMemberInfo member in specificPropertiesToSerialize)
                {
                    if (member.IsInstance || member is PSProperty)
                    {
                        continue;
                    }

                    extendedMembersList.Add(member);
                }
            }

            if (extendedMembersEnumerable is not null)
            {
                WriteMemberInfoCollection(extendedMembersEnumerable, depth, isTopLevelObjectMemberSet: true);
            }
        }

        private void WriteMemberInfoCollection(
            IEnumerable<PSMemberInfo> me,
            int depth,
            bool isTopLevelObjectMemberSet)
        {
            int notePropField = isTopLevelObjectMemberSet ? ProtoSchema.Obj_NoteProperties : ProtoSchema.MS_NoteProperties;
            int memberSetField = isTopLevelObjectMemberSet ? ProtoSchema.Obj_MemberSets : ProtoSchema.MS_NestedSets;

            foreach (PSMemberInfo info in me)
            {
                if (!info.ShouldSerialize)
                {
                    continue;
                }

                int depthOfMember = info.IsInstance ? depth : depth - 1;

                if (info.MemberType == (info.MemberType & PSMemberTypes.Properties))
                {
                    object? value = SerializationUtilities.GetPropertyValueInThreadSafeManner(
                        (PSPropertyInfo)info,
                        CanUseDefaultRunspaceInThreadSafeManner,
                        out bool gotValue);
                    if (gotValue)
                    {
                        WriteNestedPSValueField(notePropField, value, info.Name, depthOfMember);
                    }
                }
                else if (info.MemberType == PSMemberTypes.MemberSet)
                {
                    PSMemberSet set = (PSMemberSet)info;
                    if (!set.ShouldSerialize)
                    {
                        continue;
                    }

                    int setToken = _writer.BeginSubMessage(memberSetField);
                    _writer.WriteStringField(ProtoSchema.MS_Name, set.Name);
                    WriteMemberInfoCollection(set.Members, depthOfMember, isTopLevelObjectMemberSet: false);
                    _writer.EndSubMessage(setToken);
                }
            }
        }

        private void HandlePSObjectAsString(PSObject source, int depth)
        {
            string? value = GetSerializationString(source);
            WritePrimitiveTypePSObject(source, value, depth);
        }

        private string? GetSerializationString(PSObject source)
        {
            PSPropertyInfo? serializationProperty = null;
            try
            {
                serializationProperty = source.GetStringSerializationSource(_typeTable);
            }
            catch (ExtendedTypeSystemException)
            {
            }

            string? result = null;
            if (serializationProperty is not null)
            {
                object? val = SerializationUtilities.GetPropertyValueInThreadSafeManner(
                    serializationProperty,
                    CanUseDefaultRunspaceInThreadSafeManner,
                    out bool success);
                if (success && val is not null)
                {
                    result = SerializationUtilities.GetToString(val);
                }
            }
            else
            {
                result = SerializationUtilities.GetToString(source);
            }

            return result;
        }

        private bool SerializeAsString(PSObject source)
        {
            return source.GetSerializationMethod(_typeTable) == SerializationMethod.String;
        }

        private int GetDepthOfSerialization(object source, int depth)
        {
            PSObject? pso = PSObject.AsPSObject(source);
            if (pso is null)
            {
                return depth;
            }

            if (pso.BaseObject is PSCredential or SwitchParameter)
            {
                return 1;
            }

            if (pso.BaseObject is PSSenderInfo)
            {
                return 4;
            }

            if ((_context.options & SerializationOptions.UseDepthFromTypes) != 0)
            {
                int typesPs1xmlDepth = pso.GetSerializationDepth(_typeTable);
                if (typesPs1xmlDepth > 0 && typesPs1xmlDepth != depth)
                {
                    return typesPs1xmlDepth;
                }
            }

            if ((_context.options & SerializationOptions.PreserveSerializationSettingOfOriginal) != 0)
            {
                if (pso.IsDeserialized && depth <= 0)
                {
                    return 1;
                }
            }

            return depth;
        }

        private static bool PSObjectHasModifiedTypesCollection(PSObject pso)
        {
            ConsolidatedString currentTypes = pso.InternalTypeNames;
            Collection<string> originalTypes = pso.InternalAdapter.BaseGetTypeNameHierarchy(pso.ImmediateBaseObject);

            if (currentTypes.Count != originalTypes.Count)
            {
                return true;
            }

            using IEnumerator<string> currentEnumerator = currentTypes.GetEnumerator();
            using IEnumerator<string> originalEnumerator = originalTypes.GetEnumerator();
            while (currentEnumerator.MoveNext() && originalEnumerator.MoveNext())
            {
                if (!currentEnumerator.Current.Equals(originalEnumerator.Current, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool PSObjectHasNotes(PSObject source)
        {
            return source.InstanceMembers is not null && source.InstanceMembers.Count > 0;
        }
    }

    /// <summary>
    /// Internal Protobuf deserializer mirroring <see cref="InternalDeserializer"/>.
    /// </summary>
    internal sealed class InternalProtoDeserializer
    {
        private const int MaxDepthBelowTopLevel = 50;

        private readonly Stream _stream;
        private readonly DeserializationContext _context;
        private readonly ReferenceIdHandlerForDeserializer<object> _objectRefIdHandler = new();
        private readonly ReferenceIdHandlerForDeserializer<ConsolidatedString> _typeRefIdHandler = new();

        private byte[] _frameBuffer = new byte[1024];
        private TypeTable? _typeTable;
        private int _depthBelowTopLevel;
        private bool _isStopping;
        private bool _done;
        private int _peekedByte = -2; // -2 = not peeked, -1 = EOF

        internal InternalProtoDeserializer(Stream stream, DeserializationContext context)
        {
            Dbg.Assert(stream is not null, "caller should validate the parameter");
            Dbg.Assert(context is not null, "caller should validate the parameter");
            _stream = stream;
            _context = context;
        }

        internal TypeTable? TypeTable
        {
            get => _typeTable;
            set => _typeTable = value;
        }

        private bool DuplicateRefIdsAllowed => (_context.options & DeserializationOptions.NoRootElement) != 0;

        internal void Start()
        {
            if (PeekByte() == -1)
            {
                _done = true;
                return;
            }

            if ((_context.options & DeserializationOptions.NoRootElement) == DeserializationOptions.NoRootElement)
            {
                return;
            }

            // Verify and consume optional 4-byte "PSPB" magic header
            if (PeekByte() == 'P')
            {
                Span<byte> magic = stackalloc byte[4];
                int read = ReadBytesFromStream(magic);
                if (read < 4 || !magic.SequenceEqual(ProtoSchema.StreamMagic))
                {
                    throw new FormatException("Invalid PowerShell Protobuf (PSPB) stream header.");
                }
            }

            if (PeekByte() == -1)
            {
                _done = true;
            }
        }

        internal bool Done()
        {
            if (_done)
            {
                return true;
            }

            if (PeekByte() == -1)
            {
                _done = true;
            }

            return _done;
        }

        internal void Stop()
        {
            _isStopping = true;
        }

        private void CheckIfStopping()
        {
            if (_isStopping)
            {
                throw PSTraceSource.NewInvalidOperationException(Serialization.Stopping);
            }
        }

        internal bool TrySkipOneTopLevelObject()
        {
            CheckIfStopping();
            if (Done())
            {
                return false;
            }

            int frameLength = checked((int)ReadVarint32FromStream());
            if (frameLength == 0)
            {
                return true;
            }

            if (_stream.CanSeek)
            {
                _stream.Seek(frameLength, SeekOrigin.Current);
                return true;
            }

            EnsureFrameBuffer(Math.Min(frameLength, 4096));
            int remaining = frameLength;
            while (remaining > 0)
            {
                int chunk = Math.Min(remaining, _frameBuffer.Length);
                int read = ReadBytesFromStream(_frameBuffer.AsSpan(0, chunk));
                if (read < chunk)
                {
                    throw new FormatException("Unexpected EOF while skipping Protobuf frame.");
                }

                remaining -= read;
            }

            return true;
        }

        internal object? ReadOneTopLevelObject(out string? streamName)
        {
            CheckIfStopping();
            int frameLength = checked((int)ReadVarint32FromStream());
            EnsureFrameBuffer(frameLength);
            Span<byte> frameSpan = _frameBuffer.AsSpan(0, frameLength);
            if (frameLength > 0)
            {
                int read = ReadBytesFromStream(frameSpan);
                if (read < frameLength)
                {
                    throw new FormatException("Unexpected EOF while reading Protobuf frame.");
                }
            }

            return ReadPSValueFromSpan(frameSpan, out _, out streamName);
        }

        private object? ReadPSValueFromSpan(ReadOnlySpan<byte> span)
        {
            return ReadPSValueFromSpan(span, out _, out _);
        }

        private object? ReadPSValueFromSpan(ReadOnlySpan<byte> span, out string? name)
        {
            return ReadPSValueFromSpan(span, out name, out _);
        }

        private object? ReadPSValueFromSpan(ReadOnlySpan<byte> span, out string? name, out string? streamName)
        {
            CheckIfStopping();
            name = null;
            streamName = null;

            try
            {
                _depthBelowTopLevel++;
                if (_depthBelowTopLevel == MaxDepthBelowTopLevel)
                {
                    throw new FormatException(Serialization.DeserializationTooDeep);
                }

                object? result = ReadOneDeserializedValue(span, out name, out streamName, out bool isKnownPrimitiveType);
                if (result is null)
                {
                    return null;
                }

                if (!isKnownPrimitiveType)
                {
                    PSObject mshSource = PSObject.AsPSObject(result);

                    Type? targetType = mshSource.GetTargetTypeForDeserialization(_typeTable);
                    if (targetType is not null)
                    {
                        try
                        {
                            object? rehydratedResult = LanguagePrimitives.ConvertTo(
                                result,
                                targetType,
                                true,
                                CultureInfo.InvariantCulture,
                                _typeTable);
                            return rehydratedResult;
                        }
                        catch (InvalidCastException)
                        {
                        }
                        catch (ArgumentException)
                        {
                        }
                    }
                }

                return result;
            }
            finally
            {
                _depthBelowTopLevel--;
            }
        }

        private object? ReadOneDeserializedValue(
            ReadOnlySpan<byte> span,
            out string? name,
            out string? streamName,
            out bool isKnownPrimitiveType)
        {
            name = null;
            streamName = null;
            isKnownPrimitiveType = true;
            object? value = null;

            ProtoSpanReader reader = new(span);
            while (reader.TryReadTag(out int fieldNumber, out int wireType))
            {
                switch (fieldNumber)
                {
                    case ProtoSchema.Val_Name:
                        name = reader.ReadString();
                        break;

                    case ProtoSchema.Val_StreamName:
                        streamName = reader.ReadString();
                        break;

                    case ProtoSchema.Val_Null:
                        _ = reader.ReadBool();
                        value = null;
                        break;

                    case ProtoSchema.Val_Bool:
                        value = reader.ReadBool();
                        break;

                    case ProtoSchema.Val_String:
                        value = reader.ReadString();
                        break;

                    case ProtoSchema.Val_Int32:
                        value = reader.ReadSInt32();
                        break;

                    case ProtoSchema.Val_Int64:
                        value = reader.ReadSInt64();
                        break;

                    case ProtoSchema.Val_UInt64:
                        value = reader.ReadVarint64();
                        break;

                    case ProtoSchema.Val_Double:
                        value = reader.ReadFixed64Double();
                        break;

                    case ProtoSchema.Val_Bytes:
                        value = reader.ReadBytes();
                        break;

                    case ProtoSchema.Val_Guid:
                        {
                            ReadOnlySpan<byte> guidBytes = reader.ReadLengthDelimitedSpan();
                            value = new Guid(guidBytes, bigEndian: true);
                            break;
                        }

                    case ProtoSchema.Val_DateTime:
                        {
                            ulong packed = reader.ReadFixed64Bits();
                            long ticks = (long)(packed & 0x3FFFFFFFFFFFFFFFUL);
                            DateTimeKind kind = (DateTimeKind)(packed >> 62);
                            value = new DateTime(ticks, kind);
                            break;
                        }

                    case ProtoSchema.Val_TimeSpan:
                        value = new TimeSpan(reader.ReadSInt64());
                        break;

                    case ProtoSchema.Val_ObjectRef:
                        {
                            isKnownPrimitiveType = false;
                            string refId = reader.ReadString();
                            object? duplicate = _objectRefIdHandler.GetReferencedObject(refId);
                            if (duplicate is null)
                            {
                                throw new FormatException(StringUtil.Format(Serialization.InvalidReferenceId, refId));
                            }

                            value = duplicate;
                            break;
                        }

                    case ProtoSchema.Val_Complex:
                        {
                            isKnownPrimitiveType = false;
                            ReadOnlySpan<byte> complexSpan = reader.ReadLengthDelimitedSpan();
                            value = ReadPSObject(complexSpan);
                            break;
                        }

                    case ProtoSchema.Val_UInt32:
                        value = reader.ReadVarint32();
                        break;

                    case ProtoSchema.Val_Byte:
                        value = checked((byte)reader.ReadVarint32());
                        break;

                    case ProtoSchema.Val_SByte:
                        value = checked((sbyte)reader.ReadSInt32());
                        break;

                    case ProtoSchema.Val_Int16:
                        value = checked((short)reader.ReadSInt32());
                        break;

                    case ProtoSchema.Val_UInt16:
                        value = checked((ushort)reader.ReadVarint32());
                        break;

                    case ProtoSchema.Val_Char:
                        value = checked((char)reader.ReadVarint32());
                        break;

                    case ProtoSchema.Val_Single:
                        value = reader.ReadFixed32Single();
                        break;

                    case ProtoSchema.Val_Decimal:
                        {
                            ReadOnlySpan<byte> decBytes = reader.ReadLengthDelimitedSpan();
                            if (decBytes.Length != 16)
                            {
                                throw new FormatException("Invalid Protobuf decimal byte length.");
                            }

                            Span<int> bits = stackalloc int[4];
                            for (int i = 0; i < 4; i++)
                            {
                                bits[i] = BinaryPrimitives.ReadInt32LittleEndian(decBytes.Slice(i * 4, 4));
                            }

                            value = new decimal(bits);
                            break;
                        }

                    case ProtoSchema.Val_Uri:
                        value = new Uri(reader.ReadString(), UriKind.RelativeOrAbsolute);
                        break;

                    case ProtoSchema.Val_Version:
                        value = new Version(reader.ReadString());
                        break;

                    case ProtoSchema.Val_ScriptBlock:
                        {
                            string sbText = reader.ReadString();
                            if ((_context.options & DeserializationOptions.DeserializeScriptBlocks) == DeserializationOptions.DeserializeScriptBlocks)
                            {
                                value = ScriptBlock.Create(sbText);
                            }
                            else
                            {
                                value = sbText;
                            }

                            break;
                        }

                    case ProtoSchema.Val_XmlDocument:
                        {
                            string docAsString = reader.ReadString();
                            int? maxCharactersInDocument = null;
                            if (_context.MaximumAllowedMemory.HasValue)
                            {
                                maxCharactersInDocument = _context.MaximumAllowedMemory.Value / sizeof(char);
                            }

                            value = InternalDeserializer.LoadUnsafeXmlDocument(
                                docAsString,
                                preserveNonElements: true,
                                maxCharactersInDocument);
                            break;
                        }

                    case ProtoSchema.Val_ProgressRecord:
                        {
                            ReadOnlySpan<byte> prSpan = reader.ReadLengthDelimitedSpan();
                            value = ReadProgressRecord(prSpan);
                            break;
                        }

                    case ProtoSchema.Val_SecureString:
                        {
                            string encryptedString = reader.ReadString();
                            if (_context.cryptoHelper is not null)
                            {
                                value = _context.cryptoHelper.DecryptSecureString(encryptedString);
                            }
                            else
                            {
                                value = Microsoft.PowerShell.SecureStringHelper.Unprotect(encryptedString);
                            }

                            break;
                        }

                    default:
                        reader.SkipField(wireType);
                        break;
                }
            }

            return value;
        }

        private ProgressRecord ReadProgressRecord(ReadOnlySpan<byte> span)
        {
            string activity = string.Empty;
            int activityId = 0;
            string? currentOperation = null;
            int parentActivityId = -1;
            int percentComplete = -1;
            ProgressRecordType recordType = ProgressRecordType.Processing;
            int secondsRemaining = -1;
            string statusDescription = string.Empty;

            ProtoSpanReader reader = new(span);
            while (reader.TryReadTag(out int fieldNumber, out int wireType))
            {
                switch (fieldNumber)
                {
                    case ProtoSchema.PR_Activity:
                        activity = reader.ReadString();
                        break;
                    case ProtoSchema.PR_ActivityId:
                        activityId = reader.ReadSInt32();
                        break;
                    case ProtoSchema.PR_CurrentOperation:
                        {
                            ReadOnlySpan<byte> curOpSpan = reader.ReadLengthDelimitedSpan();
                            object? curOpObj = ReadPSValueFromSpan(curOpSpan);
                            currentOperation = curOpObj?.ToString();
                            break;
                        }
                    case ProtoSchema.PR_ParentActivityId:
                        parentActivityId = reader.ReadSInt32();
                        break;
                    case ProtoSchema.PR_PercentComplete:
                        percentComplete = reader.ReadSInt32();
                        break;
                    case ProtoSchema.PR_RecordType:
                        recordType = (ProgressRecordType)reader.ReadVarint32();
                        break;
                    case ProtoSchema.PR_SecondsRemaining:
                        secondsRemaining = reader.ReadSInt32();
                        break;
                    case ProtoSchema.PR_StatusDescription:
                        statusDescription = reader.ReadString();
                        break;
                    default:
                        reader.SkipField(wireType);
                        break;
                }
            }

            ProgressRecord record = new(activityId, activity, statusDescription);
            if (!string.IsNullOrEmpty(currentOperation))
            {
                record.CurrentOperation = currentOperation;
            }

            record.ParentActivityId = parentActivityId;
            record.PercentComplete = percentComplete;
            record.RecordType = recordType;
            record.SecondsRemaining = secondsRemaining;
            return record;
        }

        private PSObject ReadPSObject(ReadOnlySpan<byte> span)
        {
            PSObject dso = new();
            bool overrideTypeInfo = true;
            object? baseObject = null;
            bool hasBaseObject = false;

            ContainerType containerKind = ContainerType.None;
            ArrayList? collectionItems = null;
            ProtoPSDictionary? dictionary = null;
            bool adaptedPropsInitialized = false;

            ProtoSpanReader reader = new(span);
            while (reader.TryReadTag(out int fieldNumber, out int wireType))
            {
                switch (fieldNumber)
                {
                    case ProtoSchema.Obj_RefId:
                        {
                            string refId = reader.ReadString();
                            _objectRefIdHandler.SetRefId(dso, refId, DuplicateRefIdsAllowed);
                            break;
                        }

                    case ProtoSchema.Obj_TypeNames:
                        {
                            ReadOnlySpan<byte> tnSpan = reader.ReadLengthDelimitedSpan();
                            ReadTypeNames(dso, tnSpan);
                            overrideTypeInfo = false;
                            break;
                        }

                    case ProtoSchema.Obj_TypeNamesRef:
                        {
                            string refId = reader.ReadString();
                            ReadTypeNamesReference(dso, refId);
                            overrideTypeInfo = false;
                            break;
                        }

                    case ProtoSchema.Obj_ToString:
                        {
                            dso.ToStringFromDeserialization = reader.ReadString();
                            dso.InstanceMembers.Add(PSObject.DotNetInstanceAdapter.GetDotNetMethod<PSMemberInfo>(dso, "ToString"));
                            System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember("ToString");
                            dso.TokenText = dso.ToStringFromDeserialization;
                            break;
                        }

                    case ProtoSchema.Obj_BaseValue:
                        {
                            ReadOnlySpan<byte> baseSpan = reader.ReadLengthDelimitedSpan();
                            baseObject = ReadPSValueFromSpan(baseSpan);
                            hasBaseObject = true;
                            break;
                        }

                    case ProtoSchema.Obj_ContainerKind:
                        {
                            containerKind = (ContainerType)reader.ReadVarint32();
                            hasBaseObject = true;
                            if (containerKind == ContainerType.Dictionary)
                            {
                                bool isOrdered = dso.InternalTypeNames is not null
                                    && dso.InternalTypeNames.Count > 0
                                    && Deserializer.MaskDeserializationPrefix(dso.InternalTypeNames[0]) == typeof(OrderedDictionary).FullName;
                                dictionary = new ProtoPSDictionary(isOrdered);
                            }
                            else
                            {
                                collectionItems = new ArrayList();
                            }

                            break;
                        }

                    case ProtoSchema.Obj_CollectionItems:
                        {
                            ReadOnlySpan<byte> itemSpan = reader.ReadLengthDelimitedSpan();
                            collectionItems ??= new ArrayList();
                            collectionItems.Add(ReadPSValueFromSpan(itemSpan));
                            break;
                        }

                    case ProtoSchema.Obj_DictEntries:
                        {
                            ReadOnlySpan<byte> entrySpan = reader.ReadLengthDelimitedSpan();
                            if (dictionary is null)
                            {
                                bool isOrdered = dso.InternalTypeNames is not null
                                    && dso.InternalTypeNames.Count > 0
                                    && Deserializer.MaskDeserializationPrefix(dso.InternalTypeNames[0]) == typeof(OrderedDictionary).FullName;
                                dictionary = new ProtoPSDictionary(isOrdered);
                            }

                            ReadDictionaryEntry(entrySpan, dictionary);
                            break;
                        }

                    case ProtoSchema.Obj_AdaptedProps:
                        {
                            if (!adaptedPropsInitialized)
                            {
                                dso.IsDeserialized = true;
                                dso.AdaptedMembers = new PSMemberInfoInternalCollection<PSPropertyInfo>();
                                dso.InstanceMembers.Add(PSObject.DotNetInstanceAdapter.GetDotNetMethod<PSMemberInfo>(dso, "GetType"));
                                System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember("GetType");
                                dso.ClrMembers = new PSMemberInfoInternalCollection<PSPropertyInfo>();
                                adaptedPropsInitialized = true;
                            }

                            ReadOnlySpan<byte> propSpan = reader.ReadLengthDelimitedSpan();
                            object? propValue = ReadPSValueFromSpan(propSpan, out string? propName);
                            if (!string.IsNullOrEmpty(propName))
                            {
                                dso.AdaptedMembers.Add(SmaInternalBridge.CreateDeserializedPSProperty(propName, propValue));
                            }

                            break;
                        }

                    case ProtoSchema.Obj_NoteProperties:
                        {
                            ReadOnlySpan<byte> noteSpan = reader.ReadLengthDelimitedSpan();
                            object? noteVal = ReadPSValueFromSpan(noteSpan, out string? noteName);
                            if (!string.IsNullOrEmpty(noteName))
                            {
                                PSNoteProperty note = new(noteName, noteVal);
                                dso.InstanceMembers.Add(note);
                                System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember(note.Name);
                            }

                            break;
                        }

                    case ProtoSchema.Obj_MemberSets:
                        {
                            ReadOnlySpan<byte> setSpan = reader.ReadLengthDelimitedSpan();
                            ReadMemberSet(setSpan, dso.InstanceMembers);
                            break;
                        }

                    default:
                        reader.SkipField(wireType);
                        break;
                }
            }

            if (containerKind != ContainerType.None)
            {
                if (containerKind == ContainerType.Dictionary)
                {
                    dictionary ??= new ProtoPSDictionary(isOrdered: false);
                    baseObject = dictionary.DictionaryObject;
                }
                else
                {
                    collectionItems ??= new ArrayList();
                    if (containerKind == ContainerType.Stack)
                    {
                        collectionItems.Reverse();
                        baseObject = new Stack(collectionItems);
                    }
                    else if (containerKind == ContainerType.Queue)
                    {
                        baseObject = new Queue(collectionItems);
                    }
                    else
                    {
                        baseObject = collectionItems;
                    }
                }
            }

            if (hasBaseObject && baseObject is not null)
            {
                dso.SetCoreOnDeserialization(baseObject, overrideTypeInfo);
            }

            if (dso.ImmediateBaseObject is PSObject immediateBasePso)
            {
                PSObject.CopyDeserializerFields(source: immediateBasePso, target: dso);
            }

            return dso;
        }

        private void ReadTypeNames(PSObject dso, ReadOnlySpan<byte> span)
        {
            string? refId = null;
            Collection<string> typeNames = new();

            ProtoSpanReader reader = new(span);
            while (reader.TryReadTag(out int fieldNumber, out int wireType))
            {
                switch (fieldNumber)
                {
                    case ProtoSchema.TN_RefId:
                        refId = reader.ReadString();
                        break;
                    case ProtoSchema.TN_Types:
                        {
                            string item = reader.ReadString();
                            if (!string.IsNullOrEmpty(item))
                            {
                                Deserializer.AddDeserializationPrefix(ref item);
                                typeNames.Add(item);
                            }

                            break;
                        }
                    default:
                        reader.SkipField(wireType);
                        break;
                }
            }

            dso.InternalTypeNames = new ConsolidatedString(typeNames);
            if (refId is not null)
            {
                _typeRefIdHandler.SetRefId(dso.InternalTypeNames, refId, DuplicateRefIdsAllowed);
            }
        }

        private void ReadTypeNamesReference(PSObject dso, string refId)
        {
            ConsolidatedString? typeNames = _typeRefIdHandler.GetReferencedObject(refId);
            if (typeNames is null)
            {
                throw new FormatException(StringUtil.Format(Serialization.InvalidTypeHierarchyReferenceId, refId));
            }

            _context.LogExtraMemoryUsage((typeNames.Key.Length * sizeof(char)) - 8);
            dso.InternalTypeNames = new ConsolidatedString(typeNames);
        }

        private void ReadDictionaryEntry(ReadOnlySpan<byte> span, ProtoPSDictionary dictionary)
        {
            object? key = null;
            object? value = null;
            bool hasKey = false;

            ProtoSpanReader reader = new(span);
            while (reader.TryReadTag(out int fieldNumber, out int wireType))
            {
                switch (fieldNumber)
                {
                    case ProtoSchema.Dict_Key:
                        key = ReadPSValueFromSpan(reader.ReadLengthDelimitedSpan());
                        hasKey = true;
                        break;
                    case ProtoSchema.Dict_Value:
                        value = ReadPSValueFromSpan(reader.ReadLengthDelimitedSpan());
                        break;
                    default:
                        reader.SkipField(wireType);
                        break;
                }
            }

            if (!hasKey || key is null)
            {
                throw new FormatException(Serialization.NullAsDictionaryKey);
            }

            dictionary.Add(key, value);
        }

        private void ReadMemberSet(ReadOnlySpan<byte> span, PSMemberInfoCollection<PSMemberInfo> targetCollection)
        {
            string? setName = null;
            PSMemberSet? nestedSet = null;

            ProtoSpanReader reader = new(span);
            while (reader.TryReadTag(out int fieldNumber, out int wireType))
            {
                switch (fieldNumber)
                {
                    case ProtoSchema.MS_Name:
                        setName = reader.ReadString();
                        nestedSet = new PSMemberSet(setName);
                        targetCollection.Add(nestedSet);
                        System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember(setName);
                        break;

                    case ProtoSchema.MS_NoteProperties:
                        {
                            nestedSet ??= new PSMemberSet(setName ?? "PSMemberSet");
                            ReadOnlySpan<byte> noteSpan = reader.ReadLengthDelimitedSpan();
                            object? noteVal = ReadPSValueFromSpan(noteSpan, out string? noteName);
                            if (!string.IsNullOrEmpty(noteName))
                            {
                                PSNoteProperty note = new(noteName, noteVal);
                                nestedSet.Members.Add(note);
                                System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember(note.Name);
                            }

                            break;
                        }

                    case ProtoSchema.MS_NestedSets:
                        {
                            nestedSet ??= new PSMemberSet(setName ?? "PSMemberSet");
                            ReadOnlySpan<byte> childSetSpan = reader.ReadLengthDelimitedSpan();
                            ReadMemberSet(childSetSpan, nestedSet.Members);
                            break;
                        }

                    default:
                        reader.SkipField(wireType);
                        break;
                }
            }
        }

        #region Stream Helpers

        private int PeekByte()
        {
            if (_peekedByte == -2)
            {
                _peekedByte = _stream.ReadByte();
            }

            return _peekedByte;
        }

        private int ReadByteFromStream()
        {
            if (_peekedByte != -2)
            {
                int b = _peekedByte;
                _peekedByte = -2;
                return b;
            }

            return _stream.ReadByte();
        }

        private int ReadBytesFromStream(Span<byte> destination)
        {
            int totalRead = 0;
            if (_peekedByte >= 0 && !destination.IsEmpty)
            {
                destination[0] = (byte)_peekedByte;
                _peekedByte = -2;
                totalRead = 1;
            }

            while (totalRead < destination.Length)
            {
                int n = _stream.Read(destination.Slice(totalRead));
                if (n <= 0)
                {
                    break;
                }

                totalRead += n;
            }

            return totalRead;
        }

        private uint ReadVarint32FromStream()
        {
            uint result = 0;
            int shift = 0;
            while (shift < 35)
            {
                int b = ReadByteFromStream();
                if (b < 0)
                {
                    throw new FormatException("Unexpected EOF while reading Protobuf frame length varint.");
                }

                result |= (uint)(b & 0x7F) << shift;
                if ((b & 0x80) == 0)
                {
                    return result;
                }

                shift += 7;
            }

            throw new FormatException("Malformed Protobuf frame length varint.");
        }

        private void EnsureFrameBuffer(int length)
        {
            if (_frameBuffer.Length < length)
            {
                int newSize = Math.Max(_frameBuffer.Length * 2, length);
                _frameBuffer = new byte[newSize];
            }
        }

        #endregion Stream Helpers

        private sealed class ProtoPSDictionary
        {
            private readonly bool _isOrdered;
            private IDictionary _dict;
            private int _keyClashFoundIteration;

            internal ProtoPSDictionary(bool isOrdered)
            {
                _isOrdered = isOrdered;
                _dict = _isOrdered
                    ? new OrderedDictionary(StringComparer.CurrentCultureIgnoreCase)
                    : new Hashtable(StringComparer.CurrentCultureIgnoreCase);
            }

            internal object DictionaryObject => _dict;

            internal void Add(object key, object? value)
            {
                if (_keyClashFoundIteration == 0 && _dict.Contains(key))
                {
                    _keyClashFoundIteration++;
                    IDictionary newDict = _isOrdered ? new OrderedDictionary(_dict.Count) : new Hashtable(_dict.Count);
                    foreach (DictionaryEntry entry in _dict)
                    {
                        newDict.Add(entry.Key, entry.Value);
                    }

                    _dict = newDict;
                }

                if (_keyClashFoundIteration == 1 && _dict.Contains(key))
                {
                    _keyClashFoundIteration++;
                    IEqualityComparer equalityComparer = new ReferenceEqualityComparer();
                    IDictionary newDict = _isOrdered
                        ? new OrderedDictionary(_dict.Count, equalityComparer)
                        : new Hashtable(_dict.Count, equalityComparer);
                    foreach (DictionaryEntry entry in _dict)
                    {
                        newDict.Add(entry.Key, entry.Value);
                    }

                    _dict = newDict;
                }

                _dict.Add(key, value);
            }

            private sealed class ReferenceEqualityComparer : IEqualityComparer
            {
                bool IEqualityComparer.Equals(object? x, object? y) => ReferenceEquals(x, y);

                int IEqualityComparer.GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
            }
        }
    }
}
