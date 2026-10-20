// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

#nullable disable

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
    /// This class provides public functionality for serializing and deserializing a PSObject using CBOR (RFC 8949).
    /// </summary>
    public static class PSCborSerializer
    {
        private const int DefaultSerializationDepth = 1;

        /// <summary>
        /// Serializes an object into PowerShell CBOR format with a default depth of 1.
        /// </summary>
        /// <param name="source">The input object to serialize.</param>
        /// <returns>The serialized object as a CBOR byte array.</returns>
        public static byte[] Serialize(object source)
        {
            return Serialize(source, DefaultSerializationDepth);
        }

        /// <summary>
        /// Serializes an object into PowerShell CBOR format.
        /// </summary>
        /// <param name="source">The input object to serialize.</param>
        /// <param name="depth">The depth of the members to serialize.</param>
        /// <returns>The serialized object as a CBOR byte array.</returns>
        public static byte[] Serialize(object source, int depth)
        {
            return InternalCborSerializer.SerializeToBytes(
                source,
                new SerializationContext(depth, useDepthFromTypes: true));
        }

        /// <summary>
        /// Serializes an object into PowerShell CBOR format and writes it to the specified stream.
        /// </summary>
        /// <param name="source">The input object to serialize.</param>
        /// <param name="stream">The destination stream.</param>
        /// <param name="depth">The depth of the members to serialize.</param>
        public static void Serialize(object source, Stream stream, int depth)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            CborSerializer serializer = new(stream, depth, useDepthFromTypes: true);
            serializer.Serialize(source);
            serializer.Done();
        }

        /// <summary>
        /// Serializes a list of objects into PowerShell CBOR format.
        /// </summary>
        /// <param name="source">The input objects to serialize.</param>
        /// <param name="depth">The depth of the members to serialize.</param>
        /// <param name="enumerate">Enumerates input objects and serializes one at a time.</param>
        /// <returns>The serialized objects as a CBOR byte array.</returns>
        internal static byte[] Serialize(IList<object> source, int depth, bool enumerate)
        {
            using MemoryStream ms = new();
            CborSerializer serializer = new(ms, depth, useDepthFromTypes: true);

            if (enumerate)
            {
                foreach (object item in source)
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
        /// Deserializes PowerShell CBOR bytes into an object.
        /// </summary>
        /// <param name="source">The CBOR bytes representing the object to deserialize.</param>
        /// <returns>An object that represents the serialized content.</returns>
        public static object Deserialize(byte[] source)
        {
            object[] results = DeserializeAsList(source);

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
        /// Deserializes PowerShell CBOR from a stream into an object.
        /// </summary>
        /// <param name="stream">The stream containing CBOR data.</param>
        /// <returns>An object that represents the serialized content.</returns>
        public static object Deserialize(Stream stream)
        {
            object[] results = DeserializeAsList(stream);

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
        /// Deserializes PowerShell CBOR bytes into a list of objects.
        /// </summary>
        /// <param name="source">The CBOR bytes representing the objects to deserialize.</param>
        /// <returns>An object array representing the serialized content.</returns>
        public static object[] DeserializeAsList(byte[] source)
        {
            if (source is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(source));
            }

            List<object> results = new();
            CborDeserializer deserializer = new(new ReadOnlyMemory<byte>(source));
            while (!deserializer.Done())
            {
                object result = deserializer.Deserialize();
                results.Add(result);
            }

            return results.ToArray();
        }

        /// <summary>
        /// Deserializes PowerShell CBOR from a stream into a list of objects.
        /// </summary>
        /// <param name="stream">The stream containing CBOR data.</param>
        /// <returns>An object array representing the serialized content.</returns>
        public static object[] DeserializeAsList(Stream stream)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            List<object> results = new();
            CborDeserializer deserializer = new(stream);
            while (!deserializer.Done())
            {
                object result = deserializer.Deserialize();
                results.Add(result);
            }

            return results.ToArray();
        }
    }

    /// <summary>
    /// CBOR semantic tag definitions used by PowerShell CBOR serialization.
    /// </summary>
    internal static class CborTags
    {
        // Standard RFC 8949 / IANA tags
        internal const ulong StandardDateTimeString = 0;
        internal const ulong StandardEpochDateTime = 1;
        internal const ulong StandardUri = 32;
        internal const ulong StandardUuid = 37;
        internal const ulong SelfDescribedCbor = 55799;

        // PowerShell primitive type tags (1001..1019)
        internal const ulong DateTime = 1001;
        internal const ulong DateTimeOffset = 1002;
        internal const ulong TimeSpan = 1003;
        internal const ulong Decimal = 1004;
        internal const ulong Char = 1005;
        internal const ulong Byte = 1006;
        internal const ulong SByte = 1007;
        internal const ulong Int16 = 1008;
        internal const ulong UInt16 = 1009;
        internal const ulong UInt32 = 1010;
        internal const ulong Int64 = 1011;
        internal const ulong UInt64 = 1012;
        internal const ulong Version = 1013;
        internal const ulong SemanticVersion = 1014;
        internal const ulong ScriptBlock = 1015;
        internal const ulong XmlDocument = 1016;
        internal const ulong ProgressRecord = 1017;
        internal const ulong SecureString = 1018;

        // PowerShell PSObject and reference tags (1050..1052)
        internal const ulong PSObject = 1050;
        internal const ulong ObjectReference = 1051;
    }

    /// <summary>
    /// Map keys used inside a tagged PSObject CBOR map.
    /// </summary>
    internal static class CborObjectKeys
    {
        internal const string RefId = "r";
        internal const string TypeNames = "tn";
        internal const string TypeNamesRefId = "tnr";
        internal const string ToStringValue = "ts";
        internal const string BaseObject = "b";
        internal const string Container = "c";
        internal const string AdaptedProperties = "p";
        internal const string MemberSet = "ms";
        internal const string StreamName = "sn";
    }

    /// <summary>
    /// Provides functionality for serializing a PSObject into a CBOR stream.
    /// </summary>
    internal sealed class CborSerializer
    {
        private readonly InternalCborSerializer _serializer;

        internal CborSerializer(Stream stream)
            : this(stream, new SerializationContext())
        {
        }

        internal CborSerializer(Stream stream, int depth, bool useDepthFromTypes)
            : this(stream, new SerializationContext(depth, useDepthFromTypes))
        {
        }

        internal CborSerializer(Stream stream, SerializationContext context)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            if (context is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(context));
            }

            _serializer = new InternalCborSerializer(new CborStreamWriter(stream), context);
            _serializer.Start();
        }

        internal TypeTable TypeTable
        {
            get => _serializer.TypeTable;
            set => _serializer.TypeTable = value;
        }

        internal void Serialize(object source)
        {
            Serialize(source, null);
        }

        internal void Serialize(object source, string streamName)
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
    /// Provides functionality for deserializing a PSObject from a CBOR stream.
    /// </summary>
    internal sealed class CborDeserializer
    {
        private readonly InternalCborDeserializer _deserializer;

        internal CborDeserializer(Stream stream)
            : this(stream, new DeserializationContext())
        {
        }

        internal CborDeserializer(Stream stream, DeserializationContext context)
        {
            if (stream is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(stream));
            }

            if (context is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(context));
            }

            _deserializer = new InternalCborDeserializer(new CborStreamReader(stream), context);
            _deserializer.Start();
        }

        internal CborDeserializer(ReadOnlyMemory<byte> memory)
            : this(memory, new DeserializationContext())
        {
        }

        internal CborDeserializer(ReadOnlyMemory<byte> memory, DeserializationContext context)
        {
            if (context is null)
            {
                throw PSTraceSource.NewArgumentNullException(nameof(context));
            }

            _deserializer = new InternalCborDeserializer(new CborStreamReader(memory), context);
            _deserializer.Start();
        }

        internal TypeTable TypeTable
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

        internal object Deserialize()
        {
            return Deserialize(out _);
        }

        internal object Deserialize(out string streamName)
        {
            if (Done())
            {
                throw PSTraceSource.NewInvalidOperationException(Serialization.ReadCalledAfterDone);
            }

            return _deserializer.ReadOneObject(out streamName);
        }
    }

    /// <summary>
    /// Internal CBOR serializer mirroring <see cref="InternalSerializer"/>.
    /// </summary>
    internal sealed class InternalCborSerializer
    {
        private const int MaxDepthBelowTopLevel = 50;

        [ThreadStatic]
        private static CborStreamWriter s_threadStaticBufferWriter;

        private readonly CborStreamWriter _writer;
        private readonly SerializationContext _context;
        private readonly ReferenceIdHandlerForSerializer<object> _objectRefIdHandler;
        private readonly ReferenceIdHandlerForSerializer<ConsolidatedString> _typeRefIdHandler;

        private TypeTable _typeTable;
        private int _depthBelowTopLevel;
        private bool _isStopping;
        private bool? _canUseDefaultRunspaceInThreadSafeManner;
        private Collection<CollectionEntry<PSMemberInfo>> _extendedMembersCollection;
        private Collection<CollectionEntry<PSPropertyInfo>> _allPropertiesCollection;

        internal static byte[] SerializeToBytes(object source, SerializationContext context, string streamName = null)
        {
            CborStreamWriter writer = s_threadStaticBufferWriter;
            if (writer is not null)
            {
                s_threadStaticBufferWriter = null;
                writer.Reset();
            }
            else
            {
                writer = new CborStreamWriter(initialCapacity: 1024);
            }

            try
            {
                InternalCborSerializer serializer = new(writer, context);
                serializer.Start();
                serializer.WriteOneTopLevelObject(source, streamName);
                serializer.End();
                return writer.ToByteArray();
            }
            finally
            {
                s_threadStaticBufferWriter = writer;
            }
        }

        internal static Google.Protobuf.ByteString SerializeToByteString(object source, SerializationContext context, string streamName = null)
        {
            CborStreamWriter writer = s_threadStaticBufferWriter;
            if (writer is not null)
            {
                s_threadStaticBufferWriter = null;
                writer.Reset();
            }
            else
            {
                writer = new CborStreamWriter(initialCapacity: 1024);
            }

            try
            {
                InternalCborSerializer serializer = new(writer, context);
                serializer.Start();
                serializer.WriteOneTopLevelObject(source, streamName);
                serializer.End();
                return Google.Protobuf.ByteString.CopyFrom(writer.WrittenSpan);
            }
            finally
            {
                s_threadStaticBufferWriter = writer;
            }
        }

        internal InternalCborSerializer(CborStreamWriter writer, SerializationContext context)
        {
            Dbg.Assert(writer is not null, "caller should validate the parameter");
            Dbg.Assert(context is not null, "caller should validate the parameter");
            _writer = writer;
            _context = context;

            IDictionary<object, ulong> objectRefIdDictionary = null;
            if ((_context.options & SerializationOptions.NoObjectRefIds) == 0)
            {
                objectRefIdDictionary = new WeakReferenceDictionary<ulong>();
            }

            _objectRefIdHandler = new ReferenceIdHandlerForSerializer<object>(objectRefIdDictionary);
            _typeRefIdHandler = new ReferenceIdHandlerForSerializer<ConsolidatedString>(
                new Dictionary<ConsolidatedString, ulong>(ConsolidatedString.EqualityComparer));
        }

        internal TypeTable TypeTable
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
                _writer.WriteTag(CborTags.SelfDescribedCbor);
                _writer.WriteStartIndefiniteArray();
            }
        }

        internal void End()
        {
            if ((_context.options & SerializationOptions.NoRootElement) != SerializationOptions.NoRootElement)
            {
                _writer.WriteEndIndefinite();
            }

            _writer.Flush();
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

        internal void WriteOneTopLevelObject(object source, string streamName)
        {
            Dbg.Assert(_depthBelowTopLevel == 0, "InternalCborSerializer.depthBelowTopLevel should be 0 at top-level");
            WriteOneObject(source, streamName, _context.depth);
            _writer.Flush();
        }

        private void WriteOneObject(object source, string streamName, int depth)
        {
            Dbg.Assert(depth >= 0, "depth should always be greater or equal to zero");
            CheckIfStopping();

            if (source is null)
            {
                _writer.WriteNull();
                return;
            }

            try
            {
                _depthBelowTopLevel++;
                if (_depthBelowTopLevel == MaxDepthBelowTopLevel)
                {
                    _writer.WriteTextString(Serialization.DeserializationTooDeep);
                    return;
                }

                if (HandlePrimitiveKnownTypeByConvertingToPSObject(source, streamName, depth))
                {
                    return;
                }

                string refId = _objectRefIdHandler.GetRefId(source);
                if (refId is not null)
                {
                    WritePSObjectReference(refId);
                    return;
                }

                if (HandlePrimitiveKnownTypePSObject(source, streamName, depth))
                {
                    return;
                }

                depth = GetDepthOfSerialization(source, depth);

                if (HandleKnownContainerTypes(source, streamName, depth))
                {
                    return;
                }

                PSObject mshSource = PSObject.AsPSObject(source);
                if (depth == 0 || SerializeAsString(mshSource))
                {
                    HandlePSObjectAsString(mshSource, streamName, depth);
                    return;
                }

                HandleComplexTypePSObject(source, streamName, depth);
            }
            finally
            {
                _depthBelowTopLevel--;
            }
        }

        private bool HandlePrimitiveKnownTypeByConvertingToPSObject(object source, string streamName, int depth)
        {
            if (source is not PSObject && IsKnownCborPrimitive(source.GetType()))
            {
                if (streamName is null)
                {
                    WriteOnePrimitiveKnownType(source);
                    return true;
                }

                PSObject pktInfoPSObject = PSObject.AsPSObject(source);
                WritePrimitiveTypePSObject(pktInfoPSObject, source, streamName, depth);
                return true;
            }

            return false;
        }

        private bool HandlePrimitiveKnownTypePSObject(object source, string streamName, int depth)
        {
            if (source is PSObject moSource && !moSource.ImmediateBaseObjectIsEmpty)
            {
                object baseObject = moSource.ImmediateBaseObject;
                if (IsKnownCborPrimitive(baseObject.GetType()))
                {
                    WritePrimitiveTypePSObject(moSource, baseObject, streamName, depth);
                    return true;
                }
            }

            return false;
        }

        private static bool IsKnownCborPrimitive(Type type)
        {
            return KnownTypes.GetTypeSerializationInfo(type) is not null;
        }

        private void WritePrimitiveTypePSObject(PSObject source, object primitive, string streamName, int depth)
        {
            string toStringValue = SerializationUtilities.GetToStringForPrimitiveObject(source);
            bool hasModifiedTypesCollection = PSObjectHasModifiedTypesCollection(source);
            bool hasNotes = PSObjectHasNotes(source);
            bool hasModifiedToString = toStringValue is not null;

            if (hasNotes || hasModifiedTypesCollection || hasModifiedToString || streamName is not null)
            {
                depth = GetDepthOfSerialization(source, depth);
                string refId = _objectRefIdHandler.SetRefId(source);
                _writer.WriteTag(CborTags.PSObject);
                _writer.WriteStartIndefiniteMap();

                WritePSObjectHeaderEntries(
                    source,
                    streamName,
                    refId,
                    writeTypeNames: hasModifiedTypesCollection,
                    toStringValue: toStringValue);

                if (primitive is not null)
                {
                    _writer.WriteTextString(CborObjectKeys.BaseObject);
                    WriteOnePrimitiveKnownType(primitive);
                }

                SerializeInstanceProperties(source, depth);

                _writer.WriteEndIndefinite();
            }
            else if (primitive is not null)
            {
                WriteOnePrimitiveKnownType(primitive);
            }
            else
            {
                _writer.WriteNull();
            }
        }

        private void WriteOnePrimitiveKnownType(object source)
        {
            switch (source)
            {
                case bool b:
                    _writer.WriteBoolean(b);
                    break;
                case int i32:
                    _writer.WriteInt64(i32);
                    break;
                case long i64:
                    _writer.WriteTag(CborTags.Int64);
                    _writer.WriteInt64(i64);
                    break;
                case ulong u64:
                    _writer.WriteTag(CborTags.UInt64);
                    _writer.WriteUInt64(u64);
                    break;
                case byte u8:
                    _writer.WriteTag(CborTags.Byte);
                    _writer.WriteUInt64(u8);
                    break;
                case sbyte i8:
                    _writer.WriteTag(CborTags.SByte);
                    _writer.WriteInt64(i8);
                    break;
                case short i16:
                    _writer.WriteTag(CborTags.Int16);
                    _writer.WriteInt64(i16);
                    break;
                case ushort u16:
                    _writer.WriteTag(CborTags.UInt16);
                    _writer.WriteUInt64(u16);
                    break;
                case uint u32:
                    _writer.WriteTag(CborTags.UInt32);
                    _writer.WriteUInt64(u32);
                    break;
                case float f32:
                    _writer.WriteSingle(f32);
                    break;
                case double f64:
                    _writer.WriteDouble(f64);
                    break;
                case decimal dec:
                    {
                        int[] bits = decimal.GetBits(dec);
                        _writer.WriteTag(CborTags.Decimal);
                        _writer.WriteStartArray(4);
                        for (int i = 0; i < 4; i++)
                        {
                            _writer.WriteInt64(bits[i]);
                        }

                        break;
                    }
                case char ch:
                    _writer.WriteTag(CborTags.Char);
                    _writer.WriteUInt64(ch);
                    break;
                case string s:
                    _writer.WriteTextString(s);
                    break;
                case byte[] bytes:
                    _writer.WriteByteString(bytes);
                    break;
                case Guid guid:
                    {
                        Span<byte> guidBytes = stackalloc byte[16];
                        guid.TryWriteBytes(guidBytes, bigEndian: true, out _);
                        _writer.WriteTag(CborTags.StandardUuid);
                        _writer.WriteByteString(guidBytes);
                        break;
                    }
                case DateTime dt:
                    _writer.WriteTag(CborTags.DateTime);
                    _writer.WriteStartArray(2);
                    _writer.WriteInt64(dt.Ticks);
                    _writer.WriteInt64((int)dt.Kind);
                    break;
                case DateTimeOffset dto:
                    _writer.WriteTag(CborTags.DateTimeOffset);
                    _writer.WriteStartArray(2);
                    _writer.WriteInt64(dto.Ticks);
                    _writer.WriteInt64(dto.Offset.Ticks);
                    break;
                case TimeSpan ts:
                    _writer.WriteTag(CborTags.TimeSpan);
                    _writer.WriteInt64(ts.Ticks);
                    break;
                case Uri uri:
                    _writer.WriteTag(CborTags.StandardUri);
                    _writer.WriteTextString(uri.OriginalString);
                    break;
                case Version ver:
                    _writer.WriteTag(CborTags.Version);
                    _writer.WriteTextString(ver.ToString());
                    break;
                case SemanticVersion semVer:
                    _writer.WriteTag(CborTags.SemanticVersion);
                    _writer.WriteTextString(semVer.ToString());
                    break;
                case ScriptBlock sb:
                    _writer.WriteTag(CborTags.ScriptBlock);
                    _writer.WriteTextString(sb.ToString());
                    break;
                case XmlDocument xmlDoc:
                    _writer.WriteTag(CborTags.XmlDocument);
                    _writer.WriteTextString(xmlDoc.OuterXml);
                    break;
                case ProgressRecord pr:
                    _writer.WriteTag(CborTags.ProgressRecord);
                    _writer.WriteStartArray(8);
                    _writer.WriteTextString(pr.Activity ?? string.Empty);
                    _writer.WriteInt64(pr.ActivityId);
                    WriteOneObject(pr.CurrentOperation, null, 1);
                    _writer.WriteInt64(pr.ParentActivityId);
                    _writer.WriteInt64(pr.PercentComplete);
                    _writer.WriteInt64((int)pr.RecordType);
                    _writer.WriteInt64(pr.SecondsRemaining);
                    _writer.WriteTextString(pr.StatusDescription ?? string.Empty);
                    break;
                case SecureString ss:
                    {
                        string encryptedString = _context.cryptoHelper is not null
                            ? _context.cryptoHelper.EncryptSecureString(ss)
                            : Microsoft.PowerShell.SecureStringHelper.Protect(ss);
                        _writer.WriteTag(CborTags.SecureString);
                        _writer.WriteTextString(encryptedString);
                        break;
                    }
                default:
                    _writer.WriteTextString(Convert.ToString(source, CultureInfo.InvariantCulture) ?? string.Empty);
                    break;
            }
        }

        private bool HandleKnownContainerTypes(object source, string streamName, int depth)
        {
            PSObject mshSource = source as PSObject;
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

            string refId = _objectRefIdHandler.SetRefId(source);
            _writer.WriteTag(CborTags.PSObject);
            _writer.WriteStartIndefiniteMap();

            WritePSObjectHeaderEntries(
                mshSource ?? PSObject.AsPSObject(source),
                streamName,
                refId,
                writeTypeNames: true,
                toStringValue: null);

            _writer.WriteTextString(CborObjectKeys.Container);
            _writer.WriteStartArray(2);
            _writer.WriteInt64((int)ct);

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

            _writer.WriteEndIndefinite();
            return true;
        }

        private void WriteEnumerable(IEnumerable enumerable, int depth)
        {
            _writer.WriteStartIndefiniteArray();

            IEnumerator enumerator = null;
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
                    object item;
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

                    WriteOneObject(item, null, depth);
                }
            }

            _writer.WriteEndIndefinite();
        }

        private void WriteDictionary(IDictionary dictionary, int depth)
        {
            _writer.WriteStartIndefiniteArray();

            IDictionaryEnumerator dictionaryEnum = null;
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
                    object key;
                    object value;
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

                    _writer.WriteStartArray(2);
                    WriteOneObject(key, null, depth);
                    WriteOneObject(value, null, depth);
                }
            }

            _writer.WriteEndIndefinite();
        }

        private void WritePSObjectReference(string refId)
        {
            _writer.WriteTag(CborTags.ObjectReference);
            _writer.WriteTextString(refId);
        }

        private void HandleComplexTypePSObject(object source, string streamName, int depth)
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

            string refId = _objectRefIdHandler.SetRefId(source);
            _writer.WriteTag(CborTags.PSObject);
            _writer.WriteStartIndefiniteMap();

            WritePSObjectHeaderEntries(
                mshSource,
                streamName,
                refId,
                writeTypeNames: true,
                toStringValue: writeToString ? SerializationUtilities.GetToString(mshSource) : null);

            PSMemberInfoInternalCollection<PSPropertyInfo> specificPropertiesToSerialize =
                SerializationUtilities.GetSpecificPropertiesToSerialize(mshSource, AllPropertiesCollection, _typeTable);

            if (isEnum)
            {
                object baseObject = mshSource.ImmediateBaseObject;
                object underlying = Convert.ChangeType(baseObject, Enum.GetUnderlyingType(baseObject.GetType()), CultureInfo.InvariantCulture);
                _writer.WriteTextString(CborObjectKeys.BaseObject);
                WriteOneObject(underlying, null, depth);
            }
            else if (isPSObject)
            {
                _writer.WriteTextString(CborObjectKeys.BaseObject);
                WriteOneObject(mshSource.ImmediateBaseObject, null, depth);
            }
            else if (!isErrorRecord && !isInformationalRecord)
            {
                WritePSObjectProperties(mshSource, depth, specificPropertiesToSerialize);
            }

            SerializeExtendedProperties(mshSource, depth, specificPropertiesToSerialize);

            _writer.WriteEndIndefinite();
        }

        private void WritePSObjectHeaderEntries(
            PSObject mshObject,
            string streamName,
            string refId,
            bool writeTypeNames,
            string toStringValue)
        {
            if (refId is not null)
            {
                _writer.WriteTextString(CborObjectKeys.RefId);
                _writer.WriteTextString(refId);
            }

            if (streamName is not null)
            {
                _writer.WriteTextString(CborObjectKeys.StreamName);
                _writer.WriteTextString(streamName);
            }

            if (writeTypeNames)
            {
                ConsolidatedString typeNames = mshObject.InternalTypeNames;
                if (typeNames.Count > 0)
                {
                    string typeNameHierarchyReferenceId = _typeRefIdHandler.GetRefId(typeNames);
                    if (typeNameHierarchyReferenceId is null)
                    {
                        string tnRefId = _typeRefIdHandler.SetRefId(typeNames);
                        _writer.WriteTextString(CborObjectKeys.TypeNames);
                        _writer.WriteStartArray(2);
                        _writer.WriteTextString(tnRefId);
                        _writer.WriteStartArray(typeNames.Count);
                        foreach (string type in typeNames)
                        {
                            _writer.WriteTextString(type);
                        }
                    }
                    else
                    {
                        _writer.WriteTextString(CborObjectKeys.TypeNamesRefId);
                        _writer.WriteTextString(typeNameHierarchyReferenceId);
                    }
                }
            }

            if (toStringValue is not null)
            {
                _writer.WriteTextString(CborObjectKeys.ToStringValue);
                _writer.WriteTextString(toStringValue);
            }
        }

        private void WritePSObjectProperties(
            PSObject source,
            int depth,
            IEnumerable<PSPropertyInfo> specificPropertiesToSerialize)
        {
            --depth;

            if (specificPropertiesToSerialize is not null)
            {
                SerializeProperties(specificPropertiesToSerialize, depth);
            }
            else if (source.ShouldSerializeAdapter())
            {
                IEnumerable<PSPropertyInfo> adapterCollection = source.GetAdaptedProperties();
                if (adapterCollection is not null)
                {
                    SerializeProperties(adapterCollection, depth);
                }
            }
        }

        private void SerializeProperties(IEnumerable<PSPropertyInfo> propertyCollection, int depth)
        {
            bool startWritten = false;

            foreach (PSMemberInfo info in propertyCollection)
            {
                if (info is not PSProperty prop)
                {
                    continue;
                }

                object value = SerializationUtilities.GetPropertyValueInThreadSafeManner(
                    prop,
                    CanUseDefaultRunspaceInThreadSafeManner,
                    out bool success);
                if (success)
                {
                    if (!startWritten)
                    {
                        _writer.WriteTextString(CborObjectKeys.AdaptedProperties);
                        _writer.WriteStartIndefiniteMap();
                        startWritten = true;
                    }

                    _writer.WriteTextString(prop.Name);
                    WriteOneObject(value, null, depth);
                }
            }

            if (startWritten)
            {
                _writer.WriteEndIndefinite();
            }
        }

        private void SerializeInstanceProperties(PSObject source, int depth)
        {
            PSMemberInfoCollection<PSMemberInfo> instanceMembers = SmaInternalBridge.TryGetNonEmptyInstanceMembers(source);
            if (instanceMembers is not null)
            {
                WriteMemberInfoCollection(instanceMembers, depth, writeEnclosingKey: true);
            }
        }

        private void SerializeExtendedProperties(
            PSObject source,
            int depth,
            IEnumerable<PSPropertyInfo> specificPropertiesToSerialize)
        {
            IEnumerable<PSMemberInfo> extendedMembersEnumerable;
            if (specificPropertiesToSerialize is null)
            {
                if (!SmaInternalBridge.HasTypeTableMembers(source, _typeTable))
                {
                    PSMemberInfoCollection<PSMemberInfo> instanceMembers = SmaInternalBridge.TryGetNonEmptyInstanceMembers(source);
                    if (instanceMembers is not null)
                    {
                        WriteMemberInfoCollection(instanceMembers, depth, writeEnclosingKey: true);
                    }

                    return;
                }

                PSMemberInfoIntegratingCollection<PSMemberInfo> membersToSearch =
                    new(source, ExtendedMembersCollection);
                extendedMembersEnumerable = membersToSearch.Match(
                    "*",
                    PSMemberTypes.Properties | PSMemberTypes.PropertySet | PSMemberTypes.MemberSet,
                    MshMemberMatchOptions.IncludeHidden | MshMemberMatchOptions.OnlySerializable);
            }
            else
            {
                PSMemberInfoCollection<PSMemberInfo> instanceMembers = SmaInternalBridge.TryGetNonEmptyInstanceMembers(source);
                List<PSMemberInfo> extendedMembersList = instanceMembers is not null
                    ? new(instanceMembers)
                    : new();
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
                WriteMemberInfoCollection(extendedMembersEnumerable, depth, writeEnclosingKey: true);
            }
        }

        private void WriteMemberInfoCollection(
            IEnumerable<PSMemberInfo> me,
            int depth,
            bool writeEnclosingKey)
        {
            bool enclosingStarted = false;
            foreach (PSMemberInfo info in me)
            {
                if (!info.ShouldSerialize)
                {
                    continue;
                }

                int depthOfMember = info.IsInstance ? depth : depth - 1;

                if (info.MemberType == (info.MemberType & PSMemberTypes.Properties))
                {
                    object value = SerializationUtilities.GetPropertyValueInThreadSafeManner(
                        (PSPropertyInfo)info,
                        CanUseDefaultRunspaceInThreadSafeManner,
                        out bool gotValue);
                    if (gotValue)
                    {
                        if (!enclosingStarted)
                        {
                            if (writeEnclosingKey)
                            {
                                _writer.WriteTextString(CborObjectKeys.MemberSet);
                            }

                            _writer.WriteStartIndefiniteArray();
                            enclosingStarted = true;
                        }

                        // [0, name, value] represents a NoteProperty
                        _writer.WriteStartArray(3);
                        _writer.WriteInt64(0);
                        _writer.WriteTextString(info.Name);
                        WriteOneObject(value, null, depthOfMember);
                    }
                }
                else if (info.MemberType == PSMemberTypes.MemberSet)
                {
                    PSMemberSet set = (PSMemberSet)info;
                    if (!set.ShouldSerialize)
                    {
                        continue;
                    }

                    if (!enclosingStarted)
                    {
                        if (writeEnclosingKey)
                        {
                            _writer.WriteTextString(CborObjectKeys.MemberSet);
                        }

                        _writer.WriteStartIndefiniteArray();
                        enclosingStarted = true;
                    }

                    // [1, name, [nested entries...]] represents a MemberSet
                    _writer.WriteStartArray(3);
                    _writer.WriteInt64(1);
                    _writer.WriteTextString(set.Name);
                    WriteNestedMemberSetEntries(set.Members, depthOfMember);
                }
            }

            if (enclosingStarted)
            {
                _writer.WriteEndIndefinite();
            }
        }

        private void WriteNestedMemberSetEntries(IEnumerable<PSMemberInfo> me, int depth)
        {
            _writer.WriteStartIndefiniteArray();
            foreach (PSMemberInfo info in me)
            {
                if (!info.ShouldSerialize)
                {
                    continue;
                }

                int depthOfMember = info.IsInstance ? depth : depth - 1;

                if (info.MemberType == (info.MemberType & PSMemberTypes.Properties))
                {
                    object value = SerializationUtilities.GetPropertyValueInThreadSafeManner(
                        (PSPropertyInfo)info,
                        CanUseDefaultRunspaceInThreadSafeManner,
                        out bool gotValue);
                    if (gotValue)
                    {
                        _writer.WriteStartArray(3);
                        _writer.WriteInt64(0);
                        _writer.WriteTextString(info.Name);
                        WriteOneObject(value, null, depthOfMember);
                    }
                }
                else if (info.MemberType == PSMemberTypes.MemberSet)
                {
                    PSMemberSet set = (PSMemberSet)info;
                    if (set.ShouldSerialize)
                    {
                        _writer.WriteStartArray(3);
                        _writer.WriteInt64(1);
                        _writer.WriteTextString(set.Name);
                        WriteNestedMemberSetEntries(set.Members, depthOfMember);
                    }
                }
            }

            _writer.WriteEndIndefinite();
        }

        private void HandlePSObjectAsString(PSObject source, string streamName, int depth)
        {
            string value = GetSerializationString(source);
            WritePrimitiveTypePSObject(source, value, streamName, depth);
        }

        private string GetSerializationString(PSObject source)
        {
            PSPropertyInfo serializationProperty = null;
            try
            {
                serializationProperty = source.GetStringSerializationSource(_typeTable);
            }
            catch (ExtendedTypeSystemException)
            {
            }

            string result = null;
            if (serializationProperty is not null)
            {
                object val = SerializationUtilities.GetPropertyValueInThreadSafeManner(
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
            PSObject pso = PSObject.AsPSObject(source);
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
            if (!SmaInternalBridge.HasInstanceTypeName(pso))
            {
                return false;
            }

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
            return SmaInternalBridge.HasInstanceMembers(source);
        }
    }

    /// <summary>
    /// Internal CBOR deserializer mirroring <see cref="InternalDeserializer"/>.
    /// </summary>
    internal sealed class InternalCborDeserializer
    {
        private const int MaxDepthBelowTopLevel = 50;

        private readonly CborStreamReader _reader;
        private readonly DeserializationContext _context;
        private readonly ReferenceIdHandlerForDeserializer<object> _objectRefIdHandler = new();
        private readonly ReferenceIdHandlerForDeserializer<ConsolidatedString> _typeRefIdHandler = new();

        private TypeTable _typeTable;
        private int _depthBelowTopLevel;
        private bool _isStopping;
        private bool _done;
        private bool _rootIsIndefiniteArray;
        private ulong _remainingRootItems;
        private bool _hasRootArray;

        internal InternalCborDeserializer(CborStreamReader reader, DeserializationContext context)
        {
            Dbg.Assert(reader is not null, "caller should validate the parameter");
            Dbg.Assert(context is not null, "caller should validate the parameter");
            _reader = reader;
            _context = context;
        }

        internal TypeTable TypeTable
        {
            get => _typeTable;
            set => _typeTable = value;
        }

        private bool DuplicateRefIdsAllowed => (_context.options & DeserializationOptions.NoRootElement) != 0;

        internal void Start()
        {
            if (_reader.IsEOF())
            {
                _done = true;
                return;
            }

            if ((_context.options & DeserializationOptions.NoRootElement) == DeserializationOptions.NoRootElement)
            {
                _hasRootArray = false;
                return;
            }

            // Skip optional Self-Described CBOR tag (55799)
            if (_reader.PeekMajorType() == 6)
            {
                ulong tag = _reader.PeekTag();
                if (tag == CborTags.SelfDescribedCbor)
                {
                    _reader.ReadTag();
                }
            }

            if (_reader.IsEOF())
            {
                _done = true;
                return;
            }

            if (_reader.PeekMajorType() == 4)
            {
                _hasRootArray = true;
                ulong? len = _reader.ReadStartArray();
                if (len.HasValue)
                {
                    _rootIsIndefiniteArray = false;
                    _remainingRootItems = len.Value;
                    if (_remainingRootItems == 0)
                    {
                        _done = true;
                    }
                }
                else
                {
                    _rootIsIndefiniteArray = true;
                }
            }
            else
            {
                // Single root value (non-array CBOR document)
                _hasRootArray = false;
                _remainingRootItems = 1;
            }
        }

        internal bool Done()
        {
            if (_done)
            {
                return true;
            }

            if (!_hasRootArray)
            {
                if ((_context.options & DeserializationOptions.NoRootElement) == DeserializationOptions.NoRootElement)
                {
                    _done = _reader.IsEOF();
                }
                else if (_remainingRootItems == 0 || _reader.IsEOF())
                {
                    _done = true;
                }

                return _done;
            }

            if (_rootIsIndefiniteArray)
            {
                if (_reader.TryReadBreak())
                {
                    _done = true;
                }
            }
            else if (_remainingRootItems == 0)
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

        internal object ReadOneObject(out string streamName)
        {
            if (_hasRootArray && !_rootIsIndefiniteArray && _remainingRootItems > 0)
            {
                _remainingRootItems--;
            }
            else if (!_hasRootArray && _remainingRootItems > 0)
            {
                _remainingRootItems--;
            }

            return ReadOneObjectCore(out streamName);
        }

        private object ReadOneObjectCore(out string streamName)
        {
            CheckIfStopping();
            streamName = null;

            try
            {
                _depthBelowTopLevel++;
                if (_depthBelowTopLevel == MaxDepthBelowTopLevel)
                {
                    throw new FormatException(Serialization.DeserializationTooDeep);
                }

                object result = ReadOneDeserializedObject(out streamName, out bool isKnownPrimitiveType);
                if (result is null)
                {
                    return null;
                }

                if (!isKnownPrimitiveType)
                {
                    PSObject mshSource = PSObject.AsPSObject(result);

                    Type targetType = mshSource.GetTargetTypeForDeserialization(_typeTable);
                    if (targetType is not null)
                    {
                        try
                        {
                            object rehydratedResult = LanguagePrimitives.ConvertTo(
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

        private object ReadOneObjectCore()
        {
            return ReadOneObjectCore(out _);
        }

        private object ReadOneDeserializedObject(out string streamName, out bool isKnownPrimitiveType)
        {
            streamName = null;
            isKnownPrimitiveType = true;

            int majorType = _reader.PeekMajorType();
            switch (majorType)
            {
                case 0: // Unsigned integer
                    {
                        ulong val = _reader.ReadUInt64();
                        if (val <= int.MaxValue)
                        {
                            return (int)val;
                        }

                        if (val <= long.MaxValue)
                        {
                            return (long)val;
                        }

                        return val;
                    }
                case 1: // Negative integer
                    {
                        long val = _reader.ReadInt64();
                        if (val >= int.MinValue)
                        {
                            return (int)val;
                        }

                        return val;
                    }
                case 2: // Byte string
                    return _reader.ReadByteString();
                case 3: // Text string
                    return _reader.ReadTextString();
                case 4: // Untagged CBOR Array (external CBOR compatibility)
                    {
                        isKnownPrimitiveType = false;
                        ArrayList list = ReadUntaggedArray();
                        return PSObject.AsPSObject(list);
                    }
                case 5: // Untagged CBOR Map (external CBOR compatibility)
                    {
                        isKnownPrimitiveType = false;
                        return ReadUntaggedMapAsPSObject();
                    }
                case 6: // Tagged item
                    {
                        ulong tag = _reader.ReadTag();
                        return ReadTaggedObject(tag, out streamName, out isKnownPrimitiveType);
                    }
                case 7: // Simple / float
                    return _reader.ReadSimpleOrFloat();
                default:
                    throw new FormatException($"Unexpected CBOR major type: {majorType}");
            }
        }

        private object ReadTaggedObject(ulong tag, out string streamName, out bool isKnownPrimitiveType)
        {
            streamName = null;
            isKnownPrimitiveType = true;

            switch (tag)
            {
                case CborTags.SelfDescribedCbor:
                    return ReadOneDeserializedObject(out streamName, out isKnownPrimitiveType);

                case CborTags.PSObject:
                    isKnownPrimitiveType = false;
                    return ReadPSObject(out streamName);

                case CborTags.ObjectReference:
                    {
                        isKnownPrimitiveType = false;
                        string refId = _reader.ReadTextString();
                        object duplicate = _objectRefIdHandler.GetReferencedObject(refId);
                        if (duplicate is null)
                        {
                            throw new FormatException(StringUtil.Format(Serialization.InvalidReferenceId, refId));
                        }

                        return duplicate;
                    }

                case CborTags.Byte:
                    return checked((byte)_reader.ReadUInt64());

                case CborTags.SByte:
                    return checked((sbyte)_reader.ReadInt64());

                case CborTags.Int16:
                    return checked((short)_reader.ReadInt64());

                case CborTags.UInt16:
                    return checked((ushort)_reader.ReadUInt64());

                case CborTags.UInt32:
                    return checked((uint)_reader.ReadUInt64());

                case CborTags.Int64:
                    return _reader.ReadInt64();

                case CborTags.UInt64:
                    return _reader.ReadUInt64();

                case CborTags.Char:
                    return checked((char)_reader.ReadUInt64());

                case CborTags.Decimal:
                    {
                        ulong? len = _reader.ReadStartArray();
                        if (len != 4)
                        {
                            throw new FormatException("Invalid CBOR decimal array length.");
                        }

                        int[] bits = new int[4];
                        for (int i = 0; i < 4; i++)
                        {
                            bits[i] = checked((int)_reader.ReadInt64());
                        }

                        return new decimal(bits);
                    }

                case CborTags.DateTime:
                    {
                        ulong? len = _reader.ReadStartArray();
                        if (len != 2)
                        {
                            throw new FormatException("Invalid CBOR DateTime array length.");
                        }

                        long ticks = _reader.ReadInt64();
                        DateTimeKind kind = (DateTimeKind)_reader.ReadInt64();
                        return new DateTime(ticks, kind);
                    }

                case CborTags.StandardDateTimeString:
                    return XmlConvert.ToDateTime(_reader.ReadTextString(), XmlDateTimeSerializationMode.RoundtripKind);

                case CborTags.StandardEpochDateTime:
                    {
                        int mt = _reader.PeekMajorType();
                        if (mt is 0 or 1)
                        {
                            long seconds = _reader.ReadInt64();
                            return DateTimeOffset.FromUnixTimeSeconds(seconds).UtcDateTime;
                        }
                        else
                        {
                            double seconds = Convert.ToDouble(_reader.ReadSimpleOrFloat(), CultureInfo.InvariantCulture);
                            return DateTimeOffset.FromUnixTimeMilliseconds((long)(seconds * 1000.0)).UtcDateTime;
                        }
                    }

                case CborTags.DateTimeOffset:
                    {
                        ulong? len = _reader.ReadStartArray();
                        if (len != 2)
                        {
                            throw new FormatException("Invalid CBOR DateTimeOffset array length.");
                        }

                        long ticks = _reader.ReadInt64();
                        long offsetTicks = _reader.ReadInt64();
                        return new DateTimeOffset(ticks, new TimeSpan(offsetTicks));
                    }

                case CborTags.TimeSpan:
                    return new TimeSpan(_reader.ReadInt64());

                case CborTags.StandardUuid:
                    {
                        byte[] bytes = _reader.ReadByteString();
                        return new Guid(bytes, bigEndian: true);
                    }

                case CborTags.StandardUri:
                    return new Uri(_reader.ReadTextString(), UriKind.RelativeOrAbsolute);

                case CborTags.Version:
                    return new Version(_reader.ReadTextString());

                case CborTags.SemanticVersion:
                    return new SemanticVersion(_reader.ReadTextString());

                case CborTags.ScriptBlock:
                    {
                        string sbText = _reader.ReadTextString();
                        if ((_context.options & DeserializationOptions.DeserializeScriptBlocks) == DeserializationOptions.DeserializeScriptBlocks)
                        {
                            return ScriptBlock.Create(sbText);
                        }

                        return sbText;
                    }

                case CborTags.XmlDocument:
                    {
                        string docAsString = _reader.ReadTextString();
                        int? maxCharactersInDocument = null;
                        if (_context.MaximumAllowedMemory.HasValue)
                        {
                            maxCharactersInDocument = _context.MaximumAllowedMemory.Value / sizeof(char);
                        }

                        return InternalDeserializer.LoadUnsafeXmlDocument(
                            docAsString,
                            preserveNonElements: true,
                            maxCharactersInDocument);
                    }

                case CborTags.ProgressRecord:
                    {
                        ulong? len = _reader.ReadStartArray();
                        if (len != 8)
                        {
                            throw new FormatException("Invalid CBOR ProgressRecord array length.");
                        }

                        string activity = _reader.ReadTextString();
                        int activityId = checked((int)_reader.ReadInt64());
                        object currentOpObj = ReadOneObjectCore();
                        string currentOperation = currentOpObj?.ToString();
                        int parentActivityId = checked((int)_reader.ReadInt64());
                        int percentComplete = checked((int)_reader.ReadInt64());
                        ProgressRecordType recordType = (ProgressRecordType)_reader.ReadInt64();
                        int secondsRemaining = checked((int)_reader.ReadInt64());
                        string statusDescription = _reader.ReadTextString();

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

                case CborTags.SecureString:
                    {
                        string encryptedString = _reader.ReadTextString();
                        if (_context.cryptoHelper is not null)
                        {
                            return _context.cryptoHelper.DecryptSecureString(encryptedString);
                        }

                        return Microsoft.PowerShell.SecureStringHelper.Unprotect(encryptedString);
                    }

                default:
                    return ReadOneDeserializedObject(out streamName, out isKnownPrimitiveType);
            }
        }

        private PSObject ReadPSObject(out string streamName)
        {
            streamName = null;
            PSObject dso = new();
            bool overrideTypeInfo = true;
            object baseObject = null;
            bool hasBaseObject = false;

            ulong? mapLen = _reader.ReadStartMap();
            bool isIndefinite = !mapLen.HasValue;
            ulong remaining = mapLen ?? 0;

            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                string key = _reader.ReadTextString();
                switch (key)
                {
                    case CborObjectKeys.RefId:
                        {
                            string refId = _reader.ReadTextString();
                            _objectRefIdHandler.SetRefId(dso, refId, DuplicateRefIdsAllowed);
                            break;
                        }

                    case CborObjectKeys.StreamName:
                        streamName = _reader.ReadTextString();
                        break;

                    case CborObjectKeys.TypeNames:
                        {
                            ReadTypeNames(dso);
                            overrideTypeInfo = false;
                            break;
                        }

                    case CborObjectKeys.TypeNamesRefId:
                        {
                            ReadTypeNamesReference(dso);
                            overrideTypeInfo = false;
                            break;
                        }

                    case CborObjectKeys.ToStringValue:
                        {
                            dso.ToStringFromDeserialization = _reader.ReadTextString();
                            dso.InstanceMembers.Add(PSObject.DotNetInstanceAdapter.GetDotNetMethod<PSMemberInfo>(dso, "ToString"));
                            System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember("ToString");
                            dso.TokenText = dso.ToStringFromDeserialization;
                            break;
                        }

                    case CborObjectKeys.BaseObject:
                        {
                            baseObject = ReadOneObjectCore();
                            hasBaseObject = true;
                            break;
                        }

                    case CborObjectKeys.Container:
                        {
                            baseObject = ReadKnownContainer(dso.InternalTypeNames);
                            hasBaseObject = true;
                            break;
                        }

                    case CborObjectKeys.AdaptedProperties:
                        ReadProperties(dso);
                        break;

                    case CborObjectKeys.MemberSet:
                        ReadMemberSet(dso.InstanceMembers);
                        break;

                    default:
                        SkipCborValue();
                        break;
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

        private void ReadTypeNames(PSObject dso)
        {
            ulong? arrLen = _reader.ReadStartArray();
            if (arrLen != 2)
            {
                throw new FormatException("Invalid TypeNames tuple in CBOR PSObject.");
            }

            string refId = _reader.PeekMajorType() == 7 ? (string)_reader.ReadSimpleOrFloat() : _reader.ReadTextString();

            ulong? count = _reader.ReadStartArray();
            bool isIndefinite = !count.HasValue;
            ulong remaining = count ?? 0;

            Collection<string> typeNames = new();
            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                string item = _reader.ReadTextString();
                if (!string.IsNullOrEmpty(item))
                {
                    Deserializer.AddDeserializationPrefix(ref item);
                    typeNames.Add(item);
                }
            }

            dso.InternalTypeNames = new ConsolidatedString(typeNames);
            if (refId is not null)
            {
                _typeRefIdHandler.SetRefId(dso.InternalTypeNames, refId, DuplicateRefIdsAllowed);
            }
        }

        private void ReadTypeNamesReference(PSObject dso)
        {
            string refId = _reader.ReadTextString();
            ConsolidatedString typeNames = _typeRefIdHandler.GetReferencedObject(refId);
            if (typeNames is null)
            {
                throw new FormatException(StringUtil.Format(Serialization.InvalidTypeHierarchyReferenceId, refId));
            }

            _context.LogExtraMemoryUsage((typeNames.Key.Length * sizeof(char)) - 8);
            dso.InternalTypeNames = new ConsolidatedString(typeNames);
        }

        private void ReadProperties(PSObject dso)
        {
            dso.IsDeserialized = true;
            dso.AdaptedMembers = new PSMemberInfoInternalCollection<PSPropertyInfo>();
            dso.InstanceMembers.Add(PSObject.DotNetInstanceAdapter.GetDotNetMethod<PSMemberInfo>(dso, "GetType"));
            System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember("GetType");
            dso.ClrMembers = new PSMemberInfoInternalCollection<PSPropertyInfo>();

            ulong? mapLen = _reader.ReadStartMap();
            bool isIndefinite = !mapLen.HasValue;
            ulong remaining = mapLen ?? 0;

            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                string propName = _reader.ReadTextString();
                object propValue = ReadOneObjectCore();
                dso.AdaptedMembers.Add(SmaInternalBridge.CreateDeserializedPSProperty(propName, propValue));
            }
        }

        private void ReadMemberSet(PSMemberInfoCollection<PSMemberInfo> collection)
        {
            ulong? arrLen = _reader.ReadStartArray();
            bool isIndefinite = !arrLen.HasValue;
            ulong remaining = arrLen ?? 0;

            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                ulong? tupleLen = _reader.ReadStartArray();
                if (tupleLen != 3)
                {
                    throw new FormatException("Invalid MemberSet entry tuple in CBOR PSObject.");
                }

                long entryKind = _reader.ReadInt64();
                string name = _reader.ReadTextString();

                if (entryKind == 0)
                {
                    object value = ReadOneObjectCore();
                    PSNoteProperty note = new(name, value);
                    collection.Add(note);
                    System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember(note.Name);
                }
                else if (entryKind == 1)
                {
                    PSMemberSet nestedSet = new(name);
                    collection.Add(nestedSet);
                    ReadMemberSet(nestedSet.Members);
                    System.Management.Automation.Language.PSGetMemberBinder.SetHasInstanceMember(name);
                }
                else
                {
                    SkipCborValue();
                }
            }
        }

        private object ReadKnownContainer(ConsolidatedString internalTypeNames)
        {
            ulong? tupleLen = _reader.ReadStartArray();
            if (tupleLen != 2)
            {
                throw new FormatException("Invalid Container tuple in CBOR PSObject.");
            }

            ContainerType ct = (ContainerType)_reader.ReadInt64();
            switch (ct)
            {
                case ContainerType.Dictionary:
                    return ReadDictionary(internalTypeNames);
                case ContainerType.Enumerable:
                case ContainerType.List:
                case ContainerType.Queue:
                case ContainerType.Stack:
                    return ReadListContainer(ct);
                default:
                    throw new FormatException($"Unrecognized ContainerType: {ct}");
            }
        }

        private object ReadListContainer(ContainerType ct)
        {
            ArrayList list = new();
            ulong? arrLen = _reader.ReadStartArray();
            bool isIndefinite = !arrLen.HasValue;
            ulong remaining = arrLen ?? 0;

            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                list.Add(ReadOneObjectCore());
            }

            if (ct == ContainerType.Stack)
            {
                list.Reverse();
                return new Stack(list);
            }
            else if (ct == ContainerType.Queue)
            {
                return new Queue(list);
            }

            return list;
        }

        private object ReadDictionary(ConsolidatedString internalTypeNames)
        {
            bool isOrdered = internalTypeNames is not null
                && internalTypeNames.Count > 0
                && Deserializer.MaskDeserializationPrefix(internalTypeNames[0]) == typeof(OrderedDictionary).FullName;

            CborPSDictionary dictionary = new(isOrdered);

            ulong? arrLen = _reader.ReadStartArray();
            bool isIndefinite = !arrLen.HasValue;
            ulong remaining = arrLen ?? 0;

            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                ulong? pairLen = _reader.ReadStartArray();
                if (pairLen != 2)
                {
                    throw new FormatException("Invalid Dictionary entry pair in CBOR PSObject.");
                }

                object key = ReadOneObjectCore();
                if (key is null)
                {
                    throw new FormatException(Serialization.NullAsDictionaryKey);
                }

                object value = ReadOneObjectCore();
                dictionary.Add(key, value);
            }

            return dictionary.DictionaryObject;
        }

        private ArrayList ReadUntaggedArray()
        {
            ArrayList list = new();
            ulong? arrLen = _reader.ReadStartArray();
            bool isIndefinite = !arrLen.HasValue;
            ulong remaining = arrLen ?? 0;

            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                list.Add(ReadOneObjectCore());
            }

            return list;
        }

        private PSObject ReadUntaggedMapAsPSObject()
        {
            PSObject pso = new();
            ulong? mapLen = _reader.ReadStartMap();
            bool isIndefinite = !mapLen.HasValue;
            ulong remaining = mapLen ?? 0;

            while (isIndefinite ? !_reader.TryReadBreak() : remaining-- > 0)
            {
                object keyObj = ReadOneObjectCore();
                string key = keyObj?.ToString() ?? string.Empty;
                object val = ReadOneObjectCore();
                pso.Properties.Add(new PSNoteProperty(key, val));
            }

            return pso;
        }

        private void SkipCborValue()
        {
            int mt = _reader.PeekMajorType();
            switch (mt)
            {
                case 0:
                    _reader.ReadUInt64();
                    break;
                case 1:
                    _reader.ReadInt64();
                    break;
                case 2:
                    _reader.ReadByteString();
                    break;
                case 3:
                    _reader.ReadTextString();
                    break;
                case 4:
                    {
                        ulong? len = _reader.ReadStartArray();
                        bool isIndef = !len.HasValue;
                        ulong rem = len ?? 0;
                        while (isIndef ? !_reader.TryReadBreak() : rem-- > 0)
                        {
                            SkipCborValue();
                        }

                        break;
                    }
                case 5:
                    {
                        ulong? len = _reader.ReadStartMap();
                        bool isIndef = !len.HasValue;
                        ulong rem = len ?? 0;
                        while (isIndef ? !_reader.TryReadBreak() : rem-- > 0)
                        {
                            SkipCborValue();
                            SkipCborValue();
                        }

                        break;
                    }
                case 6:
                    _reader.ReadTag();
                    SkipCborValue();
                    break;
                case 7:
                    _reader.ReadSimpleOrFloat();
                    break;
            }
        }

        private sealed class CborPSDictionary
        {
            private readonly bool _isOrdered;
            private IDictionary _dict;
            private int _keyClashFoundIteration;

            internal CborPSDictionary(bool isOrdered)
            {
                _isOrdered = isOrdered;
                _dict = _isOrdered
                    ? new OrderedDictionary(StringComparer.CurrentCultureIgnoreCase)
                    : new Hashtable(StringComparer.CurrentCultureIgnoreCase);
            }

            internal object DictionaryObject => _dict;

            internal void Add(object key, object value)
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
                bool IEqualityComparer.Equals(object x, object y) => ReferenceEquals(x, y);

                int IEqualityComparer.GetHashCode(object obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
            }
        }
    }

    /// <summary>
    /// Streaming RFC 8949 CBOR writer over a <see cref="Stream"/> or internal buffer.
    /// </summary>
    internal sealed class CborStreamWriter
    {
        private readonly Stream _stream;
        private byte[] _buffer;
        private int _pos;

        internal CborStreamWriter(Stream stream)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        }

        internal CborStreamWriter(int initialCapacity = 1024)
        {
            _buffer = new byte[Math.Max(initialCapacity, 64)];
        }

        internal void Reset()
        {
            _pos = 0;
            if (_buffer is not null && _buffer.Length > 262144)
            {
                _buffer = new byte[4096];
            }
        }

        internal ReadOnlySpan<byte> WrittenSpan => _buffer.AsSpan(0, _pos);

        internal byte[] ToByteArray()
        {
            byte[] result = new byte[_pos];
            Buffer.BlockCopy(_buffer, 0, result, 0, _pos);
            return result;
        }

        private void EnsureCapacity(int additional)
        {
            int required = _pos + additional;
            if (required <= _buffer.Length)
            {
                return;
            }

            int newSize = Math.Max(_buffer.Length * 2, required);
            byte[] next = new byte[newSize];
            Buffer.BlockCopy(_buffer, 0, next, 0, _pos);
            _buffer = next;
        }

        private void WriteRawByte(byte b)
        {
            if (_stream is not null)
            {
                _stream.WriteByte(b);
            }
            else
            {
                EnsureCapacity(1);
                _buffer[_pos++] = b;
            }
        }

        private void WriteRawBytes(ReadOnlySpan<byte> bytes)
        {
            if (_stream is not null)
            {
                _stream.Write(bytes);
            }
            else
            {
                EnsureCapacity(bytes.Length);
                bytes.CopyTo(_buffer.AsSpan(_pos));
                _pos += bytes.Length;
            }
        }

        internal void Flush() => _stream?.Flush();

        internal void WriteUInt64(ulong value) => WriteTypeAndArgument(0, value);

        internal void WriteInt64(long value)
        {
            if (value >= 0)
            {
                WriteTypeAndArgument(0, (ulong)value);
            }
            else
            {
                WriteTypeAndArgument(1, (ulong)(-1 - value));
            }
        }

        internal void WriteByteString(ReadOnlySpan<byte> bytes)
        {
            WriteTypeAndArgument(2, (ulong)bytes.Length);
            WriteRawBytes(bytes);
        }

        internal void WriteTextString(string text)
        {
            if (text is null)
            {
                WriteNull();
                return;
            }

            if (text.Length == 0)
            {
                WriteRawByte(0x60);
                return;
            }

            if (_stream is null)
            {
                int byteCount = Encoding.UTF8.GetByteCount(text);
                WriteTypeAndArgument(3, (ulong)byteCount);
                EnsureCapacity(byteCount);
                int written = Encoding.UTF8.GetBytes(text.AsSpan(), _buffer.AsSpan(_pos, byteCount));
                _pos += written;
            }
            else if (text.Length <= 85)
            {
                Span<byte> utf8 = stackalloc byte[256];
                int written = Encoding.UTF8.GetBytes(text.AsSpan(), utf8);
                WriteTypeAndArgument(3, (ulong)written);
                _stream.Write(utf8.Slice(0, written));
            }
            else
            {
                byte[] utf8 = Encoding.UTF8.GetBytes(text);
                WriteTypeAndArgument(3, (ulong)utf8.Length);
                _stream.Write(utf8, 0, utf8.Length);
            }
        }

        internal void WriteStartArray(int count) => WriteTypeAndArgument(4, (ulong)count);

        internal void WriteStartIndefiniteArray() => WriteRawByte(0x9f);

        internal void WriteStartMap(int count) => WriteTypeAndArgument(5, (ulong)count);

        internal void WriteStartIndefiniteMap() => WriteRawByte(0xbf);

        internal void WriteEndIndefinite() => WriteRawByte(0xff);

        internal void WriteTag(ulong tag) => WriteTypeAndArgument(6, tag);

        internal void WriteBoolean(bool value) => WriteRawByte(value ? (byte)0xf5 : (byte)0xf4);

        internal void WriteNull() => WriteRawByte(0xf6);

        internal void WriteSingle(float value)
        {
            if (_stream is null)
            {
                EnsureCapacity(5);
                _buffer[_pos] = 0xfa;
                BinaryPrimitives.WriteSingleBigEndian(_buffer.AsSpan(_pos + 1, 4), value);
                _pos += 5;
            }
            else
            {
                Span<byte> buf = stackalloc byte[5];
                buf[0] = 0xfa;
                BinaryPrimitives.WriteSingleBigEndian(buf.Slice(1), value);
                _stream.Write(buf);
            }
        }

        internal void WriteDouble(double value)
        {
            if (_stream is null)
            {
                EnsureCapacity(9);
                _buffer[_pos] = 0xfb;
                BinaryPrimitives.WriteDoubleBigEndian(_buffer.AsSpan(_pos + 1, 8), value);
                _pos += 9;
            }
            else
            {
                Span<byte> buf = stackalloc byte[9];
                buf[0] = 0xfb;
                BinaryPrimitives.WriteDoubleBigEndian(buf.Slice(1), value);
                _stream.Write(buf);
            }
        }

        private void WriteTypeAndArgument(int majorType, ulong argument)
        {
            byte prefix = (byte)(majorType << 5);
            if (_stream is null)
            {
                if (argument < 24)
                {
                    EnsureCapacity(1);
                    _buffer[_pos++] = (byte)(prefix | (byte)argument);
                }
                else if (argument <= byte.MaxValue)
                {
                    EnsureCapacity(2);
                    _buffer[_pos] = (byte)(prefix | 24);
                    _buffer[_pos + 1] = (byte)argument;
                    _pos += 2;
                }
                else if (argument <= ushort.MaxValue)
                {
                    EnsureCapacity(3);
                    _buffer[_pos] = (byte)(prefix | 25);
                    BinaryPrimitives.WriteUInt16BigEndian(_buffer.AsSpan(_pos + 1, 2), (ushort)argument);
                    _pos += 3;
                }
                else if (argument <= uint.MaxValue)
                {
                    EnsureCapacity(5);
                    _buffer[_pos] = (byte)(prefix | 26);
                    BinaryPrimitives.WriteUInt32BigEndian(_buffer.AsSpan(_pos + 1, 4), (uint)argument);
                    _pos += 5;
                }
                else
                {
                    EnsureCapacity(9);
                    _buffer[_pos] = (byte)(prefix | 27);
                    BinaryPrimitives.WriteUInt64BigEndian(_buffer.AsSpan(_pos + 1, 8), argument);
                    _pos += 9;
                }

                return;
            }

            if (argument < 24)
            {
                _stream.WriteByte((byte)(prefix | (byte)argument));
            }
            else if (argument <= byte.MaxValue)
            {
                Span<byte> buf = stackalloc byte[2];
                buf[0] = (byte)(prefix | 24);
                buf[1] = (byte)argument;
                _stream.Write(buf);
            }
            else if (argument <= ushort.MaxValue)
            {
                Span<byte> buf = stackalloc byte[3];
                buf[0] = (byte)(prefix | 25);
                BinaryPrimitives.WriteUInt16BigEndian(buf.Slice(1), (ushort)argument);
                _stream.Write(buf);
            }
            else if (argument <= uint.MaxValue)
            {
                Span<byte> buf = stackalloc byte[5];
                buf[0] = (byte)(prefix | 26);
                BinaryPrimitives.WriteUInt32BigEndian(buf.Slice(1), (uint)argument);
                _stream.Write(buf);
            }
            else
            {
                Span<byte> buf = stackalloc byte[9];
                buf[0] = (byte)(prefix | 27);
                BinaryPrimitives.WriteUInt64BigEndian(buf.Slice(1), argument);
                _stream.Write(buf);
            }
        }
    }

    /// <summary>
    /// Streaming RFC 8949 CBOR reader over a <see cref="Stream"/> or <see cref="ReadOnlyMemory{T}"/>.
    /// </summary>
    internal sealed class CborStreamReader
    {
        private readonly Stream _stream;
        private readonly ReadOnlyMemory<byte> _memory;
        private readonly bool _useMemory;
        private int _pos;
        private int _peekedByte = -1;

        internal CborStreamReader(Stream stream)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        }

        internal CborStreamReader(ReadOnlyMemory<byte> memory)
        {
            _memory = memory;
            _useMemory = true;
        }

        internal bool IsEOF()
        {
            if (_useMemory)
            {
                return _pos >= _memory.Length;
            }

            if (_peekedByte >= 0)
            {
                return false;
            }

            _peekedByte = _stream.ReadByte();
            return _peekedByte < 0;
        }

        internal int PeekMajorType()
        {
            int b = PeekByte();
            return (b >> 5) & 0x07;
        }

        internal bool TryReadBreak()
        {
            int b = PeekByte();
            if (b == 0xff)
            {
                if (_useMemory)
                {
                    _pos++;
                }
                else
                {
                    _peekedByte = -1;
                }

                return true;
            }

            return false;
        }

        internal ulong PeekTag()
        {
            if (_useMemory)
            {
                int savedPos = _pos;
                ulong tag = ReadTag();
                _pos = savedPos;
                return tag;
            }

            long pos = _stream.CanSeek ? _stream.Position : -1;
            int savedPeek = _peekedByte;
            if (pos >= 0)
            {
                ulong tag = ReadTag();
                _stream.Position = pos;
                _peekedByte = savedPeek;
                return tag;
            }

            int b = PeekByte();
            if (((b >> 5) & 0x07) != 6)
            {
                throw new FormatException("Expected CBOR tag.");
            }

            int additional = b & 0x1f;
            if (additional < 24)
            {
                return (ulong)additional;
            }

            return 0;
        }

        internal ulong ReadUInt64()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            if (major != 0)
            {
                throw new FormatException($"Expected CBOR unsigned integer, got major type {major}.");
            }

            return ReadArgument(initial & 0x1f);
        }

        internal long ReadInt64()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            ulong arg = ReadArgument(initial & 0x1f);
            if (major == 0)
            {
                return checked((long)arg);
            }

            if (major == 1)
            {
                return checked(-1L - (long)arg);
            }

            throw new FormatException($"Expected CBOR integer, got major type {major}.");
        }

        internal byte[] ReadByteString()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            if (major != 2)
            {
                throw new FormatException($"Expected CBOR byte string, got major type {major}.");
            }

            int additional = initial & 0x1f;
            if (additional == 31)
            {
                using MemoryStream ms = new();
                while (!TryReadBreak())
                {
                    byte[] chunk = ReadByteString();
                    ms.Write(chunk, 0, chunk.Length);
                }

                return ms.ToArray();
            }

            ulong len = ReadArgument(additional);
            int byteLen = checked((int)len);
            byte[] buf = new byte[byteLen];
            ReadExact(buf);
            return buf;
        }

        internal string ReadTextString()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            if (major != 3)
            {
                throw new FormatException($"Expected CBOR text string, got major type {major}.");
            }

            int additional = initial & 0x1f;
            if (additional == 31)
            {
                StringBuilder sb = new();
                while (!TryReadBreak())
                {
                    sb.Append(ReadTextString());
                }

                return sb.ToString();
            }

            ulong len = ReadArgument(additional);
            if (len == 0)
            {
                return string.Empty;
            }

            int byteLen = checked((int)len);
            if (_useMemory)
            {
                if (_pos + byteLen > _memory.Length)
                {
                    throw new EndOfStreamException("Unexpected end of CBOR stream.");
                }

                ReadOnlySpan<byte> slice = _memory.Span.Slice(_pos, byteLen);
                _pos += byteLen;
                return DecodeShortOrUtf8String(slice);
            }

            if (byteLen <= 256)
            {
                Span<byte> stackBuf = stackalloc byte[byteLen];
                ReadExact(stackBuf);
                return DecodeShortOrUtf8String(stackBuf);
            }

            byte[] buf = new byte[byteLen];
            ReadExact(buf);
            return Encoding.UTF8.GetString(buf);
        }

        private static string DecodeShortOrUtf8String(ReadOnlySpan<byte> slice)
        {
            if (slice.Length == 1)
            {
                return slice[0] switch
                {
                    (byte)'r' => CborObjectKeys.RefId,
                    (byte)'b' => CborObjectKeys.BaseObject,
                    (byte)'c' => CborObjectKeys.Container,
                    (byte)'p' => CborObjectKeys.AdaptedProperties,
                    (byte)'0' => "0",
                    (byte)'1' => "1",
                    _ => Encoding.UTF8.GetString(slice)
                };
            }

            if (slice.Length == 2)
            {
                byte b0 = slice[0];
                byte b1 = slice[1];
                if (b0 == (byte)'t')
                {
                    if (b1 == (byte)'n')
                    {
                        return CborObjectKeys.TypeNames;
                    }

                    if (b1 == (byte)'s')
                    {
                        return CborObjectKeys.ToStringValue;
                    }
                }
                else if (b0 == (byte)'m' && b1 == (byte)'s')
                {
                    return CborObjectKeys.MemberSet;
                }
                else if (b0 == (byte)'s' && b1 == (byte)'n')
                {
                    return CborObjectKeys.StreamName;
                }
            }
            else if (slice.Length == 3 && slice[0] == (byte)'t' && slice[1] == (byte)'n' && slice[2] == (byte)'r')
            {
                return CborObjectKeys.TypeNamesRefId;
            }

            return Encoding.UTF8.GetString(slice);
        }

        internal ulong? ReadStartArray()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            if (major != 4)
            {
                throw new FormatException($"Expected CBOR array, got major type {major}.");
            }

            int additional = initial & 0x1f;
            if (additional == 31)
            {
                return null;
            }

            return ReadArgument(additional);
        }

        internal ulong? ReadStartMap()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            if (major != 5)
            {
                throw new FormatException($"Expected CBOR map, got major type {major}.");
            }

            int additional = initial & 0x1f;
            if (additional == 31)
            {
                return null;
            }

            return ReadArgument(additional);
        }

        internal ulong ReadTag()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            if (major != 6)
            {
                throw new FormatException($"Expected CBOR tag, got major type {major}.");
            }

            return ReadArgument(initial & 0x1f);
        }

        internal object ReadSimpleOrFloat()
        {
            byte initial = ReadByte();
            int major = (initial >> 5) & 0x07;
            if (major != 7)
            {
                throw new FormatException($"Expected CBOR simple/float value, got major type {major}.");
            }

            int additional = initial & 0x1f;
            switch (additional)
            {
                case 20:
                    return false;
                case 21:
                    return true;
                case 22:
                case 23:
                    return null;
                case 25:
                    {
                        Span<byte> buf = stackalloc byte[2];
                        ReadExact(buf);
                        Half h = BinaryPrimitives.ReadHalfBigEndian(buf);
                        return (float)h;
                    }
                case 26:
                    {
                        Span<byte> buf = stackalloc byte[4];
                        ReadExact(buf);
                        return BinaryPrimitives.ReadSingleBigEndian(buf);
                    }
                case 27:
                    {
                        Span<byte> buf = stackalloc byte[8];
                        ReadExact(buf);
                        return BinaryPrimitives.ReadDoubleBigEndian(buf);
                    }
                default:
                    throw new FormatException($"Unsupported CBOR simple value: {additional}");
            }
        }

        private ulong ReadArgument(int additional)
        {
            if (additional < 24)
            {
                return (ulong)additional;
            }

            if (_useMemory)
            {
                ReadOnlySpan<byte> span = _memory.Span;
                switch (additional)
                {
                    case 24:
                        return ReadByte();
                    case 25:
                        if (_pos + 2 > span.Length)
                        {
                            throw new EndOfStreamException("Unexpected end of CBOR stream.");
                        }

                        ushort u16 = BinaryPrimitives.ReadUInt16BigEndian(span.Slice(_pos, 2));
                        _pos += 2;
                        return u16;
                    case 26:
                        if (_pos + 4 > span.Length)
                        {
                            throw new EndOfStreamException("Unexpected end of CBOR stream.");
                        }

                        uint u32 = BinaryPrimitives.ReadUInt32BigEndian(span.Slice(_pos, 4));
                        _pos += 4;
                        return u32;
                    case 27:
                        if (_pos + 8 > span.Length)
                        {
                            throw new EndOfStreamException("Unexpected end of CBOR stream.");
                        }

                        ulong u64 = BinaryPrimitives.ReadUInt64BigEndian(span.Slice(_pos, 8));
                        _pos += 8;
                        return u64;
                    default:
                        throw new FormatException($"Invalid CBOR additional information: {additional}");
                }
            }

            switch (additional)
            {
                case 24:
                    return ReadByte();
                case 25:
                    {
                        Span<byte> buf = stackalloc byte[2];
                        ReadExact(buf);
                        return BinaryPrimitives.ReadUInt16BigEndian(buf);
                    }
                case 26:
                    {
                        Span<byte> buf = stackalloc byte[4];
                        ReadExact(buf);
                        return BinaryPrimitives.ReadUInt32BigEndian(buf);
                    }
                case 27:
                    {
                        Span<byte> buf = stackalloc byte[8];
                        ReadExact(buf);
                        return BinaryPrimitives.ReadUInt64BigEndian(buf);
                    }
                default:
                    throw new FormatException($"Invalid CBOR additional information: {additional}");
            }
        }

        private int PeekByte()
        {
            if (_useMemory)
            {
                if (_pos >= _memory.Length)
                {
                    throw new EndOfStreamException("Unexpected end of CBOR stream.");
                }

                return _memory.Span[_pos];
            }

            if (_peekedByte < 0)
            {
                _peekedByte = _stream.ReadByte();
                if (_peekedByte < 0)
                {
                    throw new EndOfStreamException("Unexpected end of CBOR stream.");
                }
            }

            return _peekedByte;
        }

        private byte ReadByte()
        {
            if (_useMemory)
            {
                if (_pos >= _memory.Length)
                {
                    throw new EndOfStreamException("Unexpected end of CBOR stream.");
                }

                return _memory.Span[_pos++];
            }

            int b = PeekByte();
            _peekedByte = -1;
            return (byte)b;
        }

        private void ReadExact(Span<byte> buffer)
        {
            if (_useMemory)
            {
                if (_pos + buffer.Length > _memory.Length)
                {
                    throw new EndOfStreamException("Unexpected end of CBOR stream.");
                }

                _memory.Span.Slice(_pos, buffer.Length).CopyTo(buffer);
                _pos += buffer.Length;
                return;
            }

            int offset = 0;
            if (_peekedByte >= 0 && buffer.Length > 0)
            {
                buffer[0] = (byte)_peekedByte;
                _peekedByte = -1;
                offset = 1;
            }

            while (offset < buffer.Length)
            {
                int read = _stream.Read(buffer.Slice(offset));
                if (read <= 0)
                {
                    throw new EndOfStreamException("Unexpected end of CBOR stream.");
                }

                offset += read;
            }
        }
    }
}
