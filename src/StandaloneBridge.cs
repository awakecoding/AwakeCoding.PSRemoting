// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq.Expressions;
using System.Management.Automation.Internal;
using System.Management.Automation.Language;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security;
using System.Xml;

namespace System.Management.Automation
{
    internal static class Diagnostics
    {
        internal static void Assert(bool condition, string? message = null)
        {
            System.Diagnostics.Debug.Assert(condition, message);
        }
    }

    [Flags]
    internal enum SerializationOptions
    {
        None = 0,
        UseDepthFromTypes = 1,
        NoRootElement = 2,
        NoNamespace = 4,
        NoObjectRefIds = 8,
        PreserveSerializationSettingOfOriginal = 16,
        RemotingOptions = UseDepthFromTypes | NoRootElement | NoNamespace | PreserveSerializationSettingOfOriginal,
    }

    [Flags]
    internal enum DeserializationOptions
    {
        None = 0,
        UseNewspace = 1,
        NoRootElement = 0x100,
        NoNamespace = 0x200,
        RemotingOptions = NoRootElement | NoNamespace,
        DeserializeScriptBlocks = 0x400,
    }

    internal enum ContainerType
    {
        None = 0,
        Dictionary = 1,
        Queue = 2,
        Stack = 3,
        List = 4,
        Enumerable = 5
    }

    internal enum SerializationMethod
    {
        AllPublicProperties = 0,
        String = 1,
        SpecificProperties = 2
    }

    [Flags]
    internal enum MshMemberMatchOptions
    {
        None = 0,
        IncludeHidden = 1,
        OnlySerializable = 2
    }

    internal sealed class SerializationContext
    {
        private const int DefaultSerializationDepth = 2;

        internal readonly int depth;
        internal readonly SerializationOptions options;
        internal readonly PSRemotingCryptoHelper? cryptoHelper;

        internal SerializationContext()
            : this(DefaultSerializationDepth, useDepthFromTypes: true)
        {
        }

        internal SerializationContext(int depth, bool useDepthFromTypes)
            : this(
                depth,
                (useDepthFromTypes ? SerializationOptions.UseDepthFromTypes : SerializationOptions.None) | SerializationOptions.PreserveSerializationSettingOfOriginal,
                cryptoHelper: null)
        {
        }

        internal SerializationContext(int depth, SerializationOptions options, PSRemotingCryptoHelper? cryptoHelper)
        {
            if (depth < 1)
            {
                throw new XmlException(Serialization.DepthTooSmall);
            }

            this.depth = depth;
            this.options = options;
            this.cryptoHelper = cryptoHelper;
        }
    }

    internal sealed class DeserializationContext
    {
        internal readonly DeserializationOptions options;
        internal readonly PSRemotingCryptoHelper? cryptoHelper;
        private int _totalDataProcessed;

        internal DeserializationContext()
            : this(DeserializationOptions.None, cryptoHelper: null)
        {
        }

        internal DeserializationContext(DeserializationOptions options, PSRemotingCryptoHelper? cryptoHelper)
        {
            this.options = options;
            this.cryptoHelper = cryptoHelper;
        }

        internal int? MaximumAllowedMemory { get; set; }

        internal void LogExtraMemoryUsage(int extraMemoryUsage)
        {
            if (!MaximumAllowedMemory.HasValue || extraMemoryUsage <= 0)
            {
                return;
            }

            _totalDataProcessed += extraMemoryUsage;
            if (_totalDataProcessed > MaximumAllowedMemory.Value)
            {
                throw new XmlException(Serialization.DeserializationMemoryQuota);
            }
        }
    }

    internal static class RefIdStrings
    {
        private static readonly string[] s_smallIds = CreateSmallIds();

        private static string[] CreateSmallIds()
        {
            var arr = new string[256];
            for (int i = 0; i < arr.Length; i++)
            {
                arr[i] = i.ToString(CultureInfo.InvariantCulture);
            }

            return arr;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static string Format(ulong id) =>
            id < (ulong)s_smallIds.Length ? s_smallIds[(int)id] : id.ToString(CultureInfo.InvariantCulture);
    }

    internal sealed class ReferenceIdHandlerForSerializer<T> where T : class
    {
        private readonly IDictionary<T, ulong>? _object2refId;
        private readonly IEqualityComparer<T>? _comparer;
        private T? _firstKey;
        private ulong _seed;

        internal ReferenceIdHandlerForSerializer(IDictionary<T, ulong>? dictionary)
        {
            _object2refId = dictionary;
            if (dictionary is Dictionary<T, ulong> concreteDict)
            {
                _comparer = concreteDict.Comparer;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal string? SetRefId(T t)
        {
            if (_object2refId is null)
            {
                return null;
            }

            ulong refId = _seed++;
            if (refId == 0)
            {
                _firstKey = t;
                return "0";
            }

            if (refId == 1 && _firstKey is not null)
            {
                _object2refId.Add(_firstKey, 0UL);
            }

            _object2refId.Add(t, refId);
            return RefIdStrings.Format(refId);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal string? GetRefId(T t)
        {
            if (_object2refId is null || _seed == 0)
            {
                return null;
            }

            if (_seed == 1)
            {
                bool isMatch = _comparer is not null
                    ? _comparer.Equals(_firstKey!, t)
                    : ReferenceEquals(_firstKey, t);
                return isMatch ? "0" : null;
            }

            if (_object2refId.TryGetValue(t, out ulong refId))
            {
                return RefIdStrings.Format(refId);
            }

            return null;
        }
    }

    internal sealed class ReferenceIdHandlerForDeserializer<T> where T : class
    {
        private Dictionary<string, T>? _refId2object;
        private string? _firstRefId;
        private T? _firstObject;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal void SetRefId(T o, string refId, bool duplicateRefIdsAllowed)
        {
            if (_refId2object is null)
            {
                if (_firstRefId is null)
                {
                    _firstRefId = refId;
                    _firstObject = o;
                    return;
                }

                if (string.Equals(_firstRefId, refId, StringComparison.Ordinal))
                {
                    if (!duplicateRefIdsAllowed)
                    {
                        throw new XmlException($"Duplicate RefId: {refId}");
                    }

                    _firstObject = o;
                    return;
                }

                _refId2object = new Dictionary<string, T>(4, StringComparer.Ordinal)
                {
                    [_firstRefId] = _firstObject!
                };
            }
            else if (_refId2object.ContainsKey(refId) && !duplicateRefIdsAllowed)
            {
                throw new XmlException($"Duplicate RefId: {refId}");
            }

            _refId2object[refId] = o;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal T? GetReferencedObject(string refId)
        {
            if (_refId2object is null)
            {
                return _firstRefId is not null && string.Equals(_firstRefId, refId, StringComparison.Ordinal)
                    ? _firstObject
                    : null;
            }

            _refId2object.TryGetValue(refId, out T? result);
            return result;
        }
    }

    internal sealed class WeakReferenceDictionary<T> : IDictionary<object, T>
    {
        private Dictionary<object, T>? _table;

        private Dictionary<object, T> Table => _table ??= new Dictionary<object, T>(ReferenceEqualityComparer.Instance);

        public void Add(object key, T value) => Table.Add(key, value);

        public bool ContainsKey(object key) => _table is not null && _table.ContainsKey(key);

        public ICollection<object> Keys => throw new NotSupportedException();

        public bool Remove(object key) => _table is not null && _table.Remove(key);

        public bool TryGetValue(object key, out T value)
        {
            if (_table is not null)
            {
                return _table.TryGetValue(key, out value!);
            }

            value = default!;
            return false;
        }

        public ICollection<T> Values => throw new NotSupportedException();

        public T this[object key]
        {
            get => TryGetValue(key, out T val) ? val : throw new KeyNotFoundException();
            set => Table[key] = value;
        }

        public void Add(KeyValuePair<object, T> item) => Add(item.Key, item.Value);
        public void Clear() => _table?.Clear();
        public bool Contains(KeyValuePair<object, T> item) => ContainsKey(item.Key);
        public void CopyTo(KeyValuePair<object, T>[] array, int arrayIndex) => throw new NotSupportedException();
        public int Count => _table?.Count ?? 0;
        public bool IsReadOnly => false;
        public bool Remove(KeyValuePair<object, T> item) => Remove(item.Key);
        public IEnumerator<KeyValuePair<object, T>> GetEnumerator() => throw new NotSupportedException();
        IEnumerator IEnumerable.GetEnumerator() => throw new NotSupportedException();
    }

    internal sealed class TypeSerializationInfo
    {
        internal TypeSerializationInfo(Type type)
        {
            Type = type;
        }

        internal Type Type { get; }
    }

    internal static class KnownTypes
    {
        private static readonly Dictionary<Type, TypeSerializationInfo> s_typeTable = new()
        {
            [typeof(bool)] = new TypeSerializationInfo(typeof(bool)),
            [typeof(byte)] = new TypeSerializationInfo(typeof(byte)),
            [typeof(char)] = new TypeSerializationInfo(typeof(char)),
            [typeof(DateTime)] = new TypeSerializationInfo(typeof(DateTime)),
            [typeof(decimal)] = new TypeSerializationInfo(typeof(decimal)),
            [typeof(double)] = new TypeSerializationInfo(typeof(double)),
            [typeof(Guid)] = new TypeSerializationInfo(typeof(Guid)),
            [typeof(short)] = new TypeSerializationInfo(typeof(short)),
            [typeof(int)] = new TypeSerializationInfo(typeof(int)),
            [typeof(long)] = new TypeSerializationInfo(typeof(long)),
            [typeof(sbyte)] = new TypeSerializationInfo(typeof(sbyte)),
            [typeof(float)] = new TypeSerializationInfo(typeof(float)),
            [typeof(string)] = new TypeSerializationInfo(typeof(string)),
            [typeof(TimeSpan)] = new TypeSerializationInfo(typeof(TimeSpan)),
            [typeof(ushort)] = new TypeSerializationInfo(typeof(ushort)),
            [typeof(uint)] = new TypeSerializationInfo(typeof(uint)),
            [typeof(ulong)] = new TypeSerializationInfo(typeof(ulong)),
            [typeof(Uri)] = new TypeSerializationInfo(typeof(Uri)),
            [typeof(byte[])] = new TypeSerializationInfo(typeof(byte[])),
            [typeof(Version)] = new TypeSerializationInfo(typeof(Version)),
            [typeof(ProgressRecord)] = new TypeSerializationInfo(typeof(ProgressRecord)),
            [typeof(SecureString)] = new TypeSerializationInfo(typeof(SecureString)),
            [typeof(ScriptBlock)] = new TypeSerializationInfo(typeof(ScriptBlock)),
            [typeof(XmlDocument)] = new TypeSerializationInfo(typeof(XmlDocument)),
        };

        internal static TypeSerializationInfo? GetTypeSerializationInfo(Type type)
        {
            if (!s_typeTable.TryGetValue(type, out TypeSerializationInfo? info) && typeof(XmlDocument).IsAssignableFrom(type))
            {
                info = s_typeTable[typeof(XmlDocument)];
            }

            return info;
        }
    }

    internal static class Deserializer
    {
        internal const string DeserializationTypeNamePrefix = "Deserialized.";

        internal static void AddDeserializationPrefix(ref string type)
        {
            if (!type.StartsWith(DeserializationTypeNamePrefix, StringComparison.OrdinalIgnoreCase))
            {
                type = string.Concat(DeserializationTypeNamePrefix, type);
            }
        }

        internal static string? MaskDeserializationPrefix(string type)
        {
            if (type.StartsWith(DeserializationTypeNamePrefix, StringComparison.OrdinalIgnoreCase))
            {
                return type.Substring(DeserializationTypeNamePrefix.Length);
            }

            return type;
        }
    }

    internal sealed class PSMemberInfoInternalCollection<T> : IEnumerable<T> where T : PSMemberInfo
    {
        internal PSMemberInfoCollection<T> InnerCollection { get; }

        internal PSMemberInfoInternalCollection()
        {
            InnerCollection = SmaInternalBridge.CreateSmaInternalMemberCollection<T>();
        }

        internal PSMemberInfoInternalCollection(PSMemberInfoCollection<T> innerCollection)
        {
            InnerCollection = innerCollection;
        }

        public T? this[string name]
        {
            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            get => InnerCollection[name];
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void Add(T item)
        {
            if (item is not null)
            {
                InnerCollection.Add(item);
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public void Add(T item, bool preValidated)
        {
            if (item is not null)
            {
                InnerCollection.Add(item, preValidated);
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public IEnumerator<T> GetEnumerator() => InnerCollection.GetEnumerator();

        IEnumerator IEnumerable.GetEnumerator() => InnerCollection.GetEnumerator();
    }

    internal readonly struct PSMemberInfoIntegratingCollection<T> where T : PSMemberInfo
    {
        private readonly PSObject _source;
        private readonly object? _smaCollections;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal PSMemberInfoIntegratingCollection(PSObject source, Collection<CollectionEntry<T>>? collections)
        {
            _source = source;
            _smaCollections = collections is { Count: > 0 } ? collections[0].SmaCollection : null;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal IEnumerable<T> Match(string name, PSMemberTypes memberTypes, MshMemberMatchOptions matchOptions)
        {
            return SmaInternalBridge.MatchExtendedMembers<T>(_source, _smaCollections, name, memberTypes, (int)matchOptions);
        }
    }

    internal sealed class ConsolidatedString : Collection<string>
    {
        internal static readonly ConsolidatedString Empty = new(Array.Empty<string>());
        internal static readonly IEqualityComparer<ConsolidatedString> EqualityComparer = new ConsolidatedStringEqualityComparer();

        private string? _key;
        private int _hashCode;
        private bool _hasHashCode;
        internal Collection<string>? SmaInstance;

        internal ConsolidatedString(IEnumerable<string> strings)
        {
            if (strings is Collection<string> smaCol && SmaInternalBridge.IsSmaConsolidatedString(smaCol))
            {
                SmaInstance = smaCol;
                for (int i = 0; i < smaCol.Count; i++)
                {
                    string s = smaCol[i];
                    if (!string.IsNullOrEmpty(s))
                    {
                        Items.Add(s);
                    }
                }
            }
            else if (strings is not null)
            {
                foreach (string s in strings)
                {
                    if (!string.IsNullOrEmpty(s))
                    {
                        Items.Add(s);
                    }
                }
            }
        }

        internal ConsolidatedString(ConsolidatedString other)
            : base(new List<string>(other.Items))
        {
            _key = other._key;
            _hashCode = other._hashCode;
            _hasHashCode = other._hasHashCode;
            SmaInstance = other.SmaInstance;
        }

        protected override void InsertItem(int index, string item)
        {
            base.InsertItem(index, item);
            _key = null;
            _hasHashCode = false;
            SmaInstance = null;
        }

        protected override void SetItem(int index, string item)
        {
            base.SetItem(index, item);
            _key = null;
            _hasHashCode = false;
            SmaInstance = null;
        }

        protected override void RemoveItem(int index)
        {
            base.RemoveItem(index);
            _key = null;
            _hasHashCode = false;
            SmaInstance = null;
        }

        protected override void ClearItems()
        {
            base.ClearItems();
            _key = null;
            _hasHashCode = false;
            SmaInstance = null;
        }

        internal string Key => _key ??= string.Join(" ", Items);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal int GetCachedHashCode()
        {
            if (!_hasHashCode)
            {
                _hashCode = Key.GetHashCode(StringComparison.Ordinal);
                _hasHashCode = true;
            }

            return _hashCode;
        }

        private sealed class ConsolidatedStringEqualityComparer : IEqualityComparer<ConsolidatedString>
        {
            public bool Equals(ConsolidatedString? x, ConsolidatedString? y)
            {
                if (ReferenceEquals(x, y))
                {
                    return true;
                }

                if (x is null || y is null || x.Count != y.Count)
                {
                    return false;
                }

                for (int i = 0; i < x.Count; i++)
                {
                    if (!string.Equals(x[i], y[i], StringComparison.Ordinal))
                    {
                        return false;
                    }
                }

                return true;
            }

            public int GetHashCode(ConsolidatedString obj) => obj.GetCachedHashCode();
        }
    }

    internal sealed class CollectionEntry<T> where T : PSMemberInfo
    {
        internal object? SmaCollection { get; }

        internal CollectionEntry(object? smaCollection)
        {
            SmaCollection = smaCollection;
        }
    }

    internal static class Serialization
    {
        internal const string DepthTooSmall = "Depth must be greater than or equal to 1.";
        internal const string DeserializationMemoryQuota = "Maximum allowed memory for deserialization has been exceeded.";
        internal const string DeserializationTooDeep = "Deserialization depth exceeded maximum allowed limit.";
        internal const string NullAsDictionaryKey = "Cannot serialize a dictionary with a null key.";
        internal const string XmlStopping = "The operation was stopped.";
        internal const string Stopping = "The operation was stopped.";
        internal const string ReadCalledAfterDone = "Read cannot be called after Done.";
        internal const string InvalidTypeHierarchyReferenceId = "Invalid TypeNames ReferenceId '{0}'.";
        internal const string InvalidReferenceId = "Invalid ReferenceId '{0}'.";
    }

    internal static class StringUtil
    {
        internal static string Format(string formatSpec, params object?[] args)
            => string.Format(CultureInfo.InvariantCulture, formatSpec, args);
    }

    internal static class InternalDeserializer
    {
        private static readonly Func<string, bool, int?, XmlDocument>? s_loadUnsafeXmlDocument = CompileLoadUnsafeXmlDocument();

        private static Func<string, bool, int?, XmlDocument>? CompileLoadUnsafeXmlDocument()
        {
            MethodInfo? method = typeof(PSObject).Assembly
                .GetType("System.Management.Automation.InternalDeserializer")
                ?.GetMethod(
                    "LoadUnsafeXmlDocument",
                    BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public,
                    null,
                    new[] { typeof(string), typeof(bool), typeof(int?) },
                    null);
            return method is not null
                ? (Func<string, bool, int?, XmlDocument>)Delegate.CreateDelegate(typeof(Func<string, bool, int?, XmlDocument>), method)
                : null;
        }

        internal static XmlDocument LoadUnsafeXmlDocument(
            string xmlText,
            bool preserveNonElements,
            int? maxCharactersInDocument)
        {
            if (s_loadUnsafeXmlDocument is not null)
            {
                return s_loadUnsafeXmlDocument(xmlText, preserveNonElements, maxCharactersInDocument);
            }

            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                MaxCharactersInDocument = maxCharactersInDocument ?? 0,
                IgnoreWhitespace = !preserveNonElements,
            };
            using var sr = new StringReader(xmlText);
            using var xr = XmlReader.Create(sr, settings);
            var doc = new XmlDocument { PreserveWhitespace = preserveNonElements };
            doc.Load(xr);
            return doc;
        }
    }

    internal readonly struct AdapterProxy
    {
        private readonly object? _adapter;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal AdapterProxy(object? adapter)
        {
            _adapter = adapter;
        }

        internal Collection<string> BaseGetTypeNameHierarchy(object? baseObject)
        {
            if (_adapter is not null && baseObject is not null)
            {
                Collection<string>? names = SmaInternalBridge.InvokeAdapterBaseGetTypeNameHierarchy(_adapter, baseObject);
                if (names is not null)
                {
                    return names;
                }
            }

            var fallback = new Collection<string>();
            for (Type? t = baseObject?.GetType(); t is not null; t = t.BaseType)
            {
                fallback.Add(t.FullName ?? t.Name);
            }

            return fallback;
        }
    }

    internal sealed class DotNetAdapterProxy
    {
        private readonly object? _dotNetAdapter;
        private readonly MethodInfo? _getDotNetMethodDefinition;

        internal DotNetAdapterProxy(object? dotNetAdapter)
        {
            _dotNetAdapter = dotNetAdapter;
            if (dotNetAdapter is not null)
            {
                foreach (MethodInfo m in dotNetAdapter.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    if (m.Name == "GetDotNetMethod" && m.IsGenericMethodDefinition)
                    {
                        _getDotNetMethodDefinition = m;
                        break;
                    }
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal T? GetDotNetMethod<T>(object obj, string methodName) where T : PSMemberInfo
        {
            if (_dotNetAdapter is not null && _getDotNetMethodDefinition is not null)
            {
                return DotNetMethodCache<T>.GetOrCreate(_dotNetAdapter.GetType(), _getDotNetMethodDefinition)(_dotNetAdapter, obj, methodName);
            }

            return null;
        }

        private static class DotNetMethodCache<T> where T : PSMemberInfo
        {
            private static Func<object, object, string, T?>? s_delegate;

            [MethodImpl(MethodImplOptions.AggressiveInlining)]
            internal static Func<object, object, string, T?> GetOrCreate(Type adapterType, MethodInfo openMethod)
            {
                return s_delegate ??= Compile(adapterType, openMethod);
            }

            private static Func<object, object, string, T?> Compile(Type adapterType, MethodInfo openMethod)
            {
                MethodInfo closed = openMethod.MakeGenericMethod(typeof(T));
                ParameterExpression adapter = Expression.Parameter(typeof(object), "adapter");
                ParameterExpression obj = Expression.Parameter(typeof(object), "obj");
                ParameterExpression methodName = Expression.Parameter(typeof(string), "methodName");
                MethodCallExpression call = Expression.Call(
                    Expression.Convert(adapter, adapterType),
                    closed,
                    obj,
                    methodName);
                return Expression.Lambda<Func<object, object, string, T?>>(call, adapter, obj, methodName).Compile();
            }
        }
    }

    internal static class SmaInternalBridge
    {
        private static readonly Assembly s_sma = typeof(PSObject).Assembly;
        private const BindingFlags AllInstance = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        private const BindingFlags AllStatic = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

        private static readonly Type? s_smaConsolidatedStringType =
            s_sma.GetType("System.Management.Automation.Runspaces.ConsolidatedString");
        private static readonly Type? s_smaAdapterType =
            s_sma.GetType("System.Management.Automation.Adapter");
        private static readonly Type? s_smaInternalCollectionGeneric =
            s_sma.GetType("System.Management.Automation.PSMemberInfoInternalCollection`1");
        private static readonly Type? s_smaIntegratingCollectionGeneric =
            s_sma.GetType("System.Management.Automation.PSMemberInfoIntegratingCollection`1");
        private static readonly Type? s_mshMemberMatchOptionsType =
            s_sma.GetType("System.Management.Automation.MshMemberMatchOptions");

        // Compiled property getters/setters on PSObject
        private static readonly Func<PSObject, bool>? s_getImmediateBaseObjectIsEmpty =
            CompileInstancePropertyGetter<PSObject, bool>("ImmediateBaseObjectIsEmpty");
        private static readonly Func<PSObject, string?>? s_getToStringFromDeserialization =
            CompileInstancePropertyGetter<PSObject, string?>("ToStringFromDeserialization");
        private static readonly Action<PSObject, string?>? s_setToStringFromDeserialization =
            CompileInstancePropertySetter<PSObject, string?>("ToStringFromDeserialization");
        private static readonly Func<PSObject, string?>? s_getTokenText =
            CompileInstancePropertyGetter<PSObject, string?>("TokenText");
        private static readonly Action<PSObject, string?>? s_setTokenText =
            CompileInstancePropertySetter<PSObject, string?>("TokenText");
        private static readonly Func<PSObject, bool>? s_getIsDeserialized =
            CompileInstancePropertyGetter<PSObject, bool>("IsDeserialized");
        private static readonly Action<PSObject, bool>? s_setIsDeserialized =
            CompileInstancePropertySetter<PSObject, bool>("IsDeserialized");
        private static readonly Func<PSObject, object?>? s_getInternalAdapter =
            CompileInstancePropertyGetterAsObject<PSObject>("InternalAdapter");
        private static readonly Func<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>? s_getAdaptedMembers =
            CompileInstancePropertyGetterCast<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>("AdaptedMembers");
        private static readonly Action<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>? s_setAdaptedMembers =
            CompileInstancePropertySetterCast<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>("AdaptedMembers");
        private static readonly Func<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>? s_getClrMembers =
            CompileInstancePropertyGetterCast<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>("ClrMembers");
        private static readonly Action<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>? s_setClrMembers =
            CompileInstancePropertySetterCast<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>("ClrMembers");
        private static readonly Func<PSObject, PSMemberInfoCollection<PSMemberInfo>?>? s_getInstanceMembers =
            CompileInstancePropertyGetterCast<PSObject, PSMemberInfoCollection<PSMemberInfo>?>("InstanceMembers");
        private static readonly Func<PSObject, Collection<string>?>? s_getInternalTypeNames =
            CompileInstancePropertyGetterCast<PSObject, Collection<string>?>("InternalTypeNames");
        private static readonly Action<PSObject, Collection<string>>? s_setInternalTypeNames =
            CompileInternalTypeNamesSetter();
        private static readonly Func<IEnumerable<string>, Collection<string>>? s_createSmaConsolidatedString =
            CompileCreateSmaConsolidatedString();

        // Fast non-allocating checks on PSObject
        private static readonly Func<PSObject, PSMemberInfoCollection<PSMemberInfo>?>? s_tryGetNonEmptyInstanceMembers =
            CompileTryGetNonEmptyInstanceMembers();
        private static readonly Func<PSObject, bool>? s_hasInstanceTypeNameCheck = CompileHasInstanceTypeNameCheck();
        private static readonly Func<PSObject, TypeTable?, bool>? s_hasTypeTableMembersCheck = CompileHasTypeTableMembersCheck();

        // Compiled methods on PSObject
        private static readonly Func<object, bool, PSObject>? s_asPSObjectPreserveToString =
            CompileStaticDelegate<Func<object, bool, PSObject>>(
                typeof(PSObject).GetMethod("AsPSObject", AllStatic, null, new[] { typeof(object), typeof(bool) }, null));
        private static readonly Action<PSObject, object?, bool>? s_setCoreOnDeserialization =
            CompileStaticOrInstanceDelegate<Action<PSObject, object?, bool>>(
                typeof(PSObject).GetMethod("SetCoreOnDeserialization", AllInstance));
        private static readonly Action<PSObject, PSObject>? s_copyDeserializerFields =
            CompileStaticDelegate<Action<PSObject, PSObject>>(
                typeof(PSObject).GetMethod("CopyDeserializerFields", AllStatic));
        private static readonly Func<PSObject, TypeTable?, int>? s_getSerializationMethod =
            CompileGetSerializationMethod();
        private static readonly Func<PSObject, TypeTable?, int>? s_getSerializationDepth =
            CompileStaticOrInstanceDelegate<Func<PSObject, TypeTable?, int>>(
                typeof(PSObject).GetMethod("GetSerializationDepth", AllInstance));
        private static readonly Func<PSObject, TypeTable?, PSPropertyInfo?>? s_getStringSerializationSource =
            CompileStaticOrInstanceDelegate<Func<PSObject, TypeTable?, PSPropertyInfo?>>(
                typeof(PSObject).GetMethod("GetStringSerializationSource", AllInstance));
        private static readonly Func<PSObject, bool>? s_shouldSerializeAdapter =
            CompileStaticOrInstanceDelegate<Func<PSObject, bool>>(
                typeof(PSObject).GetMethod("ShouldSerializeAdapter", AllInstance));
        private static readonly Func<PSObject, TypeTable?, Type?>? s_getTargetTypeForDeserialization =
            CompileStaticOrInstanceDelegate<Func<PSObject, TypeTable?, Type?>>(
                typeof(PSObject).GetMethod("GetTargetTypeForDeserialization", AllInstance));
        private static readonly Func<PSObject, TypeTable?, Collection<string>?>? s_getSpecificPropertiesToSerialize =
            CompileStaticOrInstanceDelegate<Func<PSObject, TypeTable?, Collection<string>?>>(
                typeof(PSObject).GetMethod("GetSpecificPropertiesToSerialize", AllInstance));
        private static readonly Func<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>? s_getAdaptedProperties =
            CompileGetAdaptedProperties();
        private static readonly Func<PSMemberViewTypes, TypeTable?, object?>? s_psoGetMemberCollection =
            CompileGetCollectionMethod("GetMemberCollection");
        private static readonly Func<PSMemberViewTypes, TypeTable?, object?>? s_psoGetPropertyCollection =
            CompileGetCollectionMethod("GetPropertyCollection");
        private static readonly Func<PSObject, object, TypeTable?, PSMemberInfoCollection<PSPropertyInfo>?>? s_smaGetSpecificPropertiesToSerialize =
            CompileSmaGetSpecificPropertiesToSerialize();

        // Cached default member/property collections when backupTypeTable is null
        private static readonly Collection<CollectionEntry<PSMemberInfo>> s_defaultExtendedMemberCollection =
            CreateMemberCollectionCore(PSMemberViewTypes.Extended, null);
        private static readonly Collection<CollectionEntry<PSPropertyInfo>> s_defaultAllPropertyCollection =
            CreatePropertyCollectionCore(PSMemberViewTypes.Extended | PSMemberViewTypes.Adapted | PSMemberViewTypes.Base, null);

        // Cache mapping SMA ConsolidatedString instances to our ConsolidatedString wrapper
        private static readonly ConditionalWeakTable<Collection<string>, ConsolidatedString> s_typeNamesCache = new();

        // Compiled getters/ctors for PSMemberInfo / PSAliasProperty / PSProperty / Adapter / etc.
        private static readonly Func<PSMemberInfo, bool>? s_getMemberShouldSerialize =
            CompileInstancePropertyGetter<PSMemberInfo, bool>("ShouldSerialize");
        private static readonly Func<PSAliasProperty, PSMemberInfo?>? s_getAliasReferencedMember =
            CompileInstancePropertyGetter<PSAliasProperty, PSMemberInfo?>("ReferencedMember");
        private static readonly Func<string, object?, PSProperty>? s_createPsProperty =
            CompileCreatePsProperty();
        private static readonly Func<object, object, Collection<string>?>? s_adapterBaseGetTypeNameHierarchy =
            CompileAdapterBaseGetTypeNameHierarchy();
        private static readonly Action<ErrorRecord, PSObject>? s_errorRecordToPSObjectForRemoting =
            CompileStaticOrInstanceDelegate<Action<ErrorRecord, PSObject>>(
                typeof(ErrorRecord).GetMethod("ToPSObjectForRemoting", AllInstance, null, new[] { typeof(PSObject) }, null));
        private static readonly Action<InformationalRecord, PSObject>? s_informationalRecordToPSObjectForRemoting =
            CompileStaticOrInstanceDelegate<Action<InformationalRecord, PSObject>>(
                typeof(InformationalRecord).GetMethod("ToPSObjectForRemoting", AllInstance, null, new[] { typeof(PSObject) }, null));
        private static readonly Func<string, bool, ScriptBlock>? s_scriptBlockCreateDelayParsed =
            CompileStaticDelegate<Func<string, bool, ScriptBlock>>(
                typeof(ScriptBlock).GetMethod("CreateDelayParsedScriptBlock", AllStatic, null, new[] { typeof(string), typeof(bool) }, null));
        private static readonly Action<string>? s_setHasInstanceMember =
            CompileStaticDelegate<Action<string>>(
                s_sma.GetType("System.Management.Automation.Language.PSGetMemberBinder")?.GetMethod("SetHasInstanceMember", AllStatic));
        private static readonly Func<object?, Type, bool, IFormatProvider, TypeTable?, object?>? s_convertTo5 =
            CompileStaticDelegate<Func<object?, Type, bool, IFormatProvider, TypeTable?, object?>>(
                typeof(LanguagePrimitives).GetMethod(
                    "ConvertTo",
                    AllStatic,
                    null,
                    new[] { typeof(object), typeof(Type), typeof(bool), typeof(IFormatProvider), typeof(TypeTable) },
                    null));
        private static readonly Func<PSRemotingCryptoHelper, SecureString, string>? s_cryptoEncrypt =
            CompileStaticOrInstanceDelegate<Func<PSRemotingCryptoHelper, SecureString, string>>(
                typeof(PSRemotingCryptoHelper).GetMethod("EncryptSecureString", AllInstance));
        private static readonly Func<PSRemotingCryptoHelper, string, SecureString>? s_cryptoDecrypt =
            CompileStaticOrInstanceDelegate<Func<PSRemotingCryptoHelper, string, SecureString>>(
                typeof(PSRemotingCryptoHelper).GetMethod("DecryptSecureString", AllInstance));

        private static readonly FieldInfo? s_psoDotNetInstanceAdapterField =
            typeof(PSObject).GetField("DotNetInstanceAdapter", AllStatic);
        private static readonly PropertyInfo? s_psoDotNetInstanceAdapterProp =
            typeof(PSObject).GetProperty("DotNetInstanceAdapter", AllStatic);

        private static readonly DotNetAdapterProxy s_dotNetAdapterProxy = new(
            s_psoDotNetInstanceAdapterField?.GetValue(null) ?? s_psoDotNetInstanceAdapterProp?.GetValue(null));

        // Cached GetType and ToString PSMemberInfo Templates or direct lookups
        internal static DotNetAdapterProxy DotNetInstanceAdapter => s_dotNetAdapterProxy;

        private static TDelegate? CompileStaticDelegate<TDelegate>(MethodInfo? method)
            where TDelegate : class, Delegate =>
            method is not null ? (TDelegate)Delegate.CreateDelegate(typeof(TDelegate), method) : null;

        private static TDelegate? CompileStaticOrInstanceDelegate<TDelegate>(MethodInfo? method)
            where TDelegate : class, Delegate =>
            method is not null ? (TDelegate)Delegate.CreateDelegate(typeof(TDelegate), method) : null;

        private static Func<TTarget, TValue>? CompileInstancePropertyGetter<TTarget, TValue>(string propertyName)
        {
            MethodInfo? getter = typeof(TTarget).GetProperty(propertyName, AllInstance)?.GetGetMethod(nonPublic: true);
            return getter is not null ? (Func<TTarget, TValue>)Delegate.CreateDelegate(typeof(Func<TTarget, TValue>), getter) : null;
        }

        private static Action<TTarget, TValue>? CompileInstancePropertySetter<TTarget, TValue>(string propertyName)
        {
            MethodInfo? setter = typeof(TTarget).GetProperty(propertyName, AllInstance)?.GetSetMethod(nonPublic: true);
            return setter is not null ? (Action<TTarget, TValue>)Delegate.CreateDelegate(typeof(Action<TTarget, TValue>), setter) : null;
        }

        private static Func<TTarget, object?>? CompileInstancePropertyGetterAsObject<TTarget>(string propertyName)
        {
            PropertyInfo? prop = typeof(TTarget).GetProperty(propertyName, AllInstance);
            if (prop is null)
            {
                return null;
            }

            ParameterExpression target = Expression.Parameter(typeof(TTarget), "target");
            UnaryExpression body = Expression.Convert(Expression.Property(target, prop), typeof(object));
            return Expression.Lambda<Func<TTarget, object?>>(body, target).Compile();
        }

        private static Func<TTarget, TResult>? CompileInstancePropertyGetterCast<TTarget, TResult>(string propertyName)
        {
            PropertyInfo? prop = typeof(TTarget).GetProperty(propertyName, AllInstance);
            if (prop is null)
            {
                return null;
            }

            ParameterExpression target = Expression.Parameter(typeof(TTarget), "target");
            UnaryExpression body = Expression.Convert(Expression.Property(target, prop), typeof(TResult));
            return Expression.Lambda<Func<TTarget, TResult>>(body, target).Compile();
        }

        private static Action<TTarget, TValue>? CompileInstancePropertySetterCast<TTarget, TValue>(string propertyName)
        {
            PropertyInfo? prop = typeof(TTarget).GetProperty(propertyName, AllInstance);
            if (prop is null)
            {
                return null;
            }

            ParameterExpression target = Expression.Parameter(typeof(TTarget), "target");
            ParameterExpression value = Expression.Parameter(typeof(TValue), "value");
            BinaryExpression body = Expression.Assign(
                Expression.Property(target, prop),
                Expression.Convert(value, prop.PropertyType));
            return Expression.Lambda<Action<TTarget, TValue>>(body, target, value).Compile();
        }

        private static Func<IEnumerable<string>, Collection<string>>? CompileCreateSmaConsolidatedString()
        {
            if (s_smaConsolidatedStringType is null)
            {
                return null;
            }

            ConstructorInfo? enumBoolCtor = s_smaConsolidatedStringType.GetConstructor(
                AllInstance,
                null,
                new[] { typeof(IEnumerable<string>), typeof(bool) },
                null);
            if (enumBoolCtor is not null)
            {
                ParameterExpression strings = Expression.Parameter(typeof(IEnumerable<string>), "strings");
                UnaryExpression body = Expression.Convert(
                    Expression.New(enumBoolCtor, strings, Expression.Constant(true)),
                    typeof(Collection<string>));
                return Expression.Lambda<Func<IEnumerable<string>, Collection<string>>>(body, strings).Compile();
            }

            ConstructorInfo? enumCtor = s_smaConsolidatedStringType.GetConstructor(
                AllInstance,
                null,
                new[] { typeof(IEnumerable<string>) },
                null);
            if (enumCtor is not null)
            {
                ParameterExpression strings = Expression.Parameter(typeof(IEnumerable<string>), "strings");
                UnaryExpression body = Expression.Convert(
                    Expression.New(enumCtor, strings),
                    typeof(Collection<string>));
                return Expression.Lambda<Func<IEnumerable<string>, Collection<string>>>(body, strings).Compile();
            }

            return null;
        }

        private static Action<PSObject, Collection<string>>? CompileInternalTypeNamesSetter()
        {
            PropertyInfo? prop = typeof(PSObject).GetProperty("InternalTypeNames", AllInstance);
            if (prop is null || s_smaConsolidatedStringType is null)
            {
                return null;
            }

            ConstructorInfo? copyCtor = s_smaConsolidatedStringType.GetConstructor(
                AllInstance,
                null,
                new[] { s_smaConsolidatedStringType },
                null);
            if (copyCtor is null)
            {
                return null;
            }

            ParameterExpression target = Expression.Parameter(typeof(PSObject), "target");
            ParameterExpression smaInstance = Expression.Parameter(typeof(Collection<string>), "smaInstance");
            NewExpression copyExpr = Expression.New(
                copyCtor,
                Expression.Convert(smaInstance, s_smaConsolidatedStringType));
            BinaryExpression assign = Expression.Assign(Expression.Property(target, prop), copyExpr);
            return Expression.Lambda<Action<PSObject, Collection<string>>>(assign, target, smaInstance).Compile();
        }

        private static Func<PSObject, PSMemberInfoCollection<PSMemberInfo>?>? CompileTryGetNonEmptyInstanceMembers()
        {
            if (s_smaInternalCollectionGeneric is null)
            {
                return null;
            }

            Type internalCollectionType = s_smaInternalCollectionGeneric.MakeGenericType(typeof(PSMemberInfo));
            MethodInfo? method = typeof(PSObject).GetMethod(
                "HasInstanceMembers",
                AllStatic,
                null,
                new[] { typeof(object), internalCollectionType.MakeByRefType() },
                null);
            if (method is null)
            {
                return null;
            }

            ParameterExpression pso = Expression.Parameter(typeof(PSObject), "pso");
            ParameterExpression outMembers = Expression.Variable(internalCollectionType, "members");
            ConditionalExpression body = Expression.Condition(
                Expression.Call(method, Expression.Convert(pso, typeof(object)), outMembers),
                Expression.Convert(outMembers, typeof(PSMemberInfoCollection<PSMemberInfo>)),
                Expression.Constant(null, typeof(PSMemberInfoCollection<PSMemberInfo>)));
            return Expression.Lambda<Func<PSObject, PSMemberInfoCollection<PSMemberInfo>?>>(
                Expression.Block(new[] { outMembers }, body),
                pso).Compile();
        }

        private static Func<PSObject, TypeTable?, bool>? CompileHasTypeTableMembersCheck()
        {
            MethodInfo? target1 = null;
            MethodInfo? target2 = null;
            foreach (MethodInfo m in typeof(PSObject).GetMethods(AllStatic))
            {
                if (m.Name == "TypeTableGetMembersDelegate" && m.IsGenericMethodDefinition)
                {
                    int paramCount = m.GetParameters().Length;
                    if (paramCount == 1)
                    {
                        target1 = m.MakeGenericMethod(typeof(PSMemberInfo));
                    }
                    else if (paramCount == 2)
                    {
                        target2 = m.MakeGenericMethod(typeof(PSMemberInfo));
                    }
                }
            }

            if (target2 is null)
            {
                return null;
            }

            PropertyInfo? countProp = target2.ReturnType.GetProperty("Count", AllInstance);
            if (countProp is null)
            {
                return null;
            }

            ParameterExpression pso = Expression.Parameter(typeof(PSObject), "pso");
            ParameterExpression typeTable = Expression.Parameter(typeof(TypeTable), "typeTable");
            ParameterExpression resultVar = Expression.Variable(target2.ReturnType, "members");
            Expression callExpr = target1 is not null
                ? Expression.Condition(
                    Expression.Equal(typeTable, Expression.Constant(null, typeof(TypeTable))),
                    Expression.Call(target1, pso),
                    Expression.Call(target2, pso, typeTable))
                : Expression.Call(target2, pso, typeTable);
            BlockExpression block = Expression.Block(
                new[] { resultVar },
                Expression.Assign(resultVar, callExpr),
                Expression.AndAlso(
                    Expression.NotEqual(resultVar, Expression.Constant(null, target2.ReturnType)),
                    Expression.GreaterThan(Expression.Property(resultVar, countProp), Expression.Constant(0))));
            return Expression.Lambda<Func<PSObject, TypeTable?, bool>>(block, pso, typeTable).Compile();
        }

        private static Func<PSObject, bool>? CompileHasInstanceTypeNameCheck()
        {
            if (s_smaConsolidatedStringType is null)
            {
                return null;
            }

            MethodInfo? method = typeof(PSObject).GetMethod(
                "HasInstanceTypeName",
                AllStatic,
                null,
                new[] { typeof(object), s_smaConsolidatedStringType.MakeByRefType() },
                null);
            if (method is null)
            {
                return null;
            }

            ParameterExpression pso = Expression.Parameter(typeof(PSObject), "pso");
            ParameterExpression dummyOut = Expression.Variable(s_smaConsolidatedStringType, "result");
            BlockExpression block = Expression.Block(
                new[] { dummyOut },
                Expression.Call(method, Expression.Convert(pso, typeof(object)), dummyOut));
            return Expression.Lambda<Func<PSObject, bool>>(block, pso).Compile();
        }

        private static Func<PSObject, TypeTable?, int>? CompileGetSerializationMethod()
        {
            MethodInfo? method = typeof(PSObject).GetMethod("GetSerializationMethod", AllInstance);
            if (method is null)
            {
                return null;
            }

            ParameterExpression pso = Expression.Parameter(typeof(PSObject), "pso");
            ParameterExpression typeTable = Expression.Parameter(typeof(TypeTable), "typeTable");
            UnaryExpression body = Expression.Convert(Expression.Call(pso, method, typeTable), typeof(int));
            return Expression.Lambda<Func<PSObject, TypeTable?, int>>(body, pso, typeTable).Compile();
        }

        private static Func<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>? CompileGetAdaptedProperties()
        {
            MethodInfo? method = typeof(PSObject).GetMethod("GetAdaptedProperties", AllInstance, null, Type.EmptyTypes, null);
            if (method is null)
            {
                return null;
            }

            ParameterExpression pso = Expression.Parameter(typeof(PSObject), "pso");
            UnaryExpression body = Expression.Convert(Expression.Call(pso, method), typeof(PSMemberInfoCollection<PSPropertyInfo>));
            return Expression.Lambda<Func<PSObject, PSMemberInfoCollection<PSPropertyInfo>?>>(body, pso).Compile();
        }

        private static Func<PSMemberViewTypes, TypeTable?, object?>? CompileGetCollectionMethod(string methodName)
        {
            MethodInfo? method = typeof(PSObject).GetMethod(
                methodName,
                AllStatic,
                null,
                new[] { typeof(PSMemberViewTypes), typeof(TypeTable) },
                null);
            if (method is null)
            {
                return null;
            }

            ParameterExpression viewType = Expression.Parameter(typeof(PSMemberViewTypes), "viewType");
            ParameterExpression typeTable = Expression.Parameter(typeof(TypeTable), "typeTable");
            UnaryExpression body = Expression.Convert(Expression.Call(method, viewType, typeTable), typeof(object));
            return Expression.Lambda<Func<PSMemberViewTypes, TypeTable?, object?>>(body, viewType, typeTable).Compile();
        }

        private static Func<PSObject, object, TypeTable?, PSMemberInfoCollection<PSPropertyInfo>?>? CompileSmaGetSpecificPropertiesToSerialize()
        {
            MethodInfo? method = s_sma.GetType("System.Management.Automation.SerializationUtilities")
                ?.GetMethod("GetSpecificPropertiesToSerialize", AllStatic);
            if (method is null)
            {
                return null;
            }

            ParameterInfo[] parameters = method.GetParameters();
            if (parameters.Length != 3)
            {
                return null;
            }

            ParameterExpression source = Expression.Parameter(typeof(PSObject), "source");
            ParameterExpression allProps = Expression.Parameter(typeof(object), "allProps");
            ParameterExpression typeTable = Expression.Parameter(typeof(TypeTable), "typeTable");
            MethodCallExpression call = Expression.Call(
                method,
                source,
                Expression.Convert(allProps, parameters[1].ParameterType),
                typeTable);
            return Expression.Lambda<Func<PSObject, object, TypeTable?, PSMemberInfoCollection<PSPropertyInfo>?>>(
                Expression.Convert(call, typeof(PSMemberInfoCollection<PSPropertyInfo>)),
                source,
                allProps,
                typeTable).Compile();
        }

        private static Func<string, object?, PSProperty>? CompileCreatePsProperty()
        {
            ConstructorInfo? ctor = typeof(PSProperty).GetConstructor(AllInstance, null, new[] { typeof(string), typeof(object) }, null);
            if (ctor is null)
            {
                return null;
            }

            ParameterExpression name = Expression.Parameter(typeof(string), "name");
            ParameterExpression value = Expression.Parameter(typeof(object), "value");
            return Expression.Lambda<Func<string, object?, PSProperty>>(Expression.New(ctor, name, value), name, value).Compile();
        }

        private static Func<object, object, Collection<string>?>? CompileAdapterBaseGetTypeNameHierarchy()
        {
            if (s_smaAdapterType is null)
            {
                return null;
            }

            MethodInfo? method = s_smaAdapterType.GetMethod("BaseGetTypeNameHierarchy", AllInstance, null, new[] { typeof(object) }, null);
            if (method is null)
            {
                return null;
            }

            ParameterExpression adapter = Expression.Parameter(typeof(object), "adapter");
            ParameterExpression baseObj = Expression.Parameter(typeof(object), "baseObj");
            MethodCallExpression call = Expression.Call(
                Expression.Convert(adapter, s_smaAdapterType),
                method,
                baseObj);
            return Expression.Lambda<Func<object, object, Collection<string>?>>(
                Expression.Convert(call, typeof(Collection<string>)),
                adapter,
                baseObj).Compile();
        }

        private static class InternalCollectionFactory<T> where T : PSMemberInfo
        {
            internal static readonly Func<PSMemberInfoCollection<T>> Create = CompileFactory();

            private static Func<PSMemberInfoCollection<T>> CompileFactory()
            {
                Type closedType = s_smaInternalCollectionGeneric!.MakeGenericType(typeof(T));
                ConstructorInfo ctor = closedType.GetConstructor(AllInstance, null, Type.EmptyTypes, null)!;
                return Expression.Lambda<Func<PSMemberInfoCollection<T>>>(
                    Expression.Convert(Expression.New(ctor), typeof(PSMemberInfoCollection<T>))).Compile();
            }
        }

        private static class ExtendedMemberMatcher<T> where T : PSMemberInfo
        {
            internal static readonly Func<PSObject, object, string, PSMemberTypes, int, IEnumerable<T>>? Match = CompileMatcher();

            private static Func<PSObject, object, string, PSMemberTypes, int, IEnumerable<T>>? CompileMatcher()
            {
                if (s_smaIntegratingCollectionGeneric is null || s_mshMemberMatchOptionsType is null)
                {
                    return null;
                }

                Type closedType = s_smaIntegratingCollectionGeneric.MakeGenericType(typeof(T));
                ConstructorInfo? ctor = Array.Find(closedType.GetConstructors(AllInstance), c => c.GetParameters().Length == 2);
                MethodInfo? matchMethod = closedType.GetMethod(
                    "Match",
                    AllInstance,
                    null,
                    new[] { typeof(string), typeof(PSMemberTypes), s_mshMemberMatchOptionsType },
                    null);
                if (ctor is null || matchMethod is null)
                {
                    return null;
                }

                ParameterInfo[] ctorParams = ctor.GetParameters();
                ParameterExpression source = Expression.Parameter(typeof(PSObject), "source");
                ParameterExpression collections = Expression.Parameter(typeof(object), "collections");
                ParameterExpression name = Expression.Parameter(typeof(string), "name");
                ParameterExpression memberTypes = Expression.Parameter(typeof(PSMemberTypes), "memberTypes");
                ParameterExpression matchOptions = Expression.Parameter(typeof(int), "matchOptions");

                NewExpression instance = Expression.New(
                    ctor,
                    source,
                    Expression.Convert(collections, ctorParams[1].ParameterType));
                MethodCallExpression call = Expression.Call(
                    instance,
                    matchMethod,
                    name,
                    memberTypes,
                    Expression.Convert(matchOptions, s_mshMemberMatchOptionsType));
                return Expression.Lambda<Func<PSObject, object, string, PSMemberTypes, int, IEnumerable<T>>>(
                    Expression.Convert(call, typeof(IEnumerable<T>)),
                    source,
                    collections,
                    name,
                    memberTypes,
                    matchOptions).Compile();
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool IsSmaConsolidatedString(Collection<string> col) =>
            s_smaConsolidatedStringType is not null && s_smaConsolidatedStringType.IsInstanceOfType(col);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static AdapterProxy GetInternalAdapter(PSObject pso)
            => new AdapterProxy(s_getInternalAdapter is not null ? s_getInternalAdapter(pso) : null);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static Collection<string>? InvokeAdapterBaseGetTypeNameHierarchy(object adapter, object baseObject) =>
            s_adapterBaseGetTypeNameHierarchy is not null ? s_adapterBaseGetTypeNameHierarchy(adapter, baseObject) : null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool GetImmediateBaseObjectIsEmpty(PSObject pso) =>
            s_getImmediateBaseObjectIsEmpty is not null
                ? s_getImmediateBaseObjectIsEmpty(pso)
                : pso.ImmediateBaseObject is PSCustomObject;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfoCollection<PSMemberInfo>? TryGetNonEmptyInstanceMembers(PSObject pso) =>
            s_tryGetNonEmptyInstanceMembers is not null
                ? s_tryGetNonEmptyInstanceMembers(pso)
                : (GetInstanceMembers(pso) is { Count: > 0 } members ? members : null);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool HasInstanceMembers(PSObject pso) =>
            TryGetNonEmptyInstanceMembers(pso) is not null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool HasInstanceTypeName(PSObject pso) =>
            s_hasInstanceTypeNameCheck is not null && s_hasInstanceTypeNameCheck(pso);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool HasTypeTableMembers(PSObject pso, TypeTable? typeTable) =>
            s_hasTypeTableMembersCheck is null || s_hasTypeTableMembersCheck(pso, typeTable);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static ConsolidatedString GetInternalTypeNames(PSObject pso)
        {
            if (s_getInternalTypeNames is not null)
            {
                Collection<string>? smaTypeNames = s_getInternalTypeNames(pso);
                if (smaTypeNames is not null)
                {
                    return s_typeNamesCache.GetValue(smaTypeNames, static tn => new ConsolidatedString(tn));
                }
            }

            return new ConsolidatedString(pso.TypeNames);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetInternalTypeNames(PSObject pso, ConsolidatedString? value)
        {
            if (value is null)
            {
                pso.TypeNames.Clear();
                return;
            }

            if (s_setInternalTypeNames is not null && s_createSmaConsolidatedString is not null)
            {
                Collection<string> smaInstance = value.SmaInstance ??= s_createSmaConsolidatedString(value);
                s_setInternalTypeNames(pso, smaInstance);
                return;
            }

            pso.TypeNames.Clear();
            foreach (string t in value)
            {
                pso.TypeNames.Add(t);
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static string? GetToStringFromDeserialization(PSObject pso) =>
            s_getToStringFromDeserialization is not null ? s_getToStringFromDeserialization(pso) : null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetToStringFromDeserialization(PSObject pso, string? value) =>
            s_setToStringFromDeserialization?.Invoke(pso, value);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static string? GetTokenText(PSObject pso) =>
            s_getTokenText is not null ? s_getTokenText(pso) : null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetTokenText(PSObject pso, string? value) =>
            s_setTokenText?.Invoke(pso, value);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool GetIsDeserialized(PSObject pso) =>
            s_getIsDeserialized is not null && s_getIsDeserialized(pso);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetIsDeserialized(PSObject pso, bool value) =>
            s_setIsDeserialized?.Invoke(pso, value);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetCoreOnDeserialization(PSObject pso, object? value, bool overrideTypeInfo) =>
            s_setCoreOnDeserialization?.Invoke(pso, value, overrideTypeInfo);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void CopyDeserializerFields(PSObject source, PSObject target) =>
            s_copyDeserializerFields?.Invoke(source, target);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSObject AsPSObjectWithPreserveToString(object obj, bool preserveToString) =>
            s_asPSObjectPreserveToString is not null
                ? s_asPSObjectPreserveToString(obj, preserveToString)
                : PSObject.AsPSObject(obj);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static SerializationMethod GetSerializationMethod(PSObject pso, TypeTable? typeTable) =>
            s_getSerializationMethod is not null
                ? (SerializationMethod)s_getSerializationMethod(pso, typeTable)
                : SerializationMethod.AllPublicProperties;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static int GetSerializationDepth(PSObject pso, TypeTable? typeTable) =>
            s_getSerializationDepth is not null ? s_getSerializationDepth(pso, typeTable) : 0;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSPropertyInfo? GetStringSerializationSource(PSObject pso, TypeTable? typeTable) =>
            s_getStringSerializationSource is not null ? s_getStringSerializationSource(pso, typeTable) : null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool ShouldSerializeAdapter(PSObject pso) =>
            s_shouldSerializeAdapter is null || s_shouldSerializeAdapter(pso);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static Type? GetTargetTypeForDeserialization(PSObject pso, TypeTable? typeTable) =>
            s_getTargetTypeForDeserialization is not null ? s_getTargetTypeForDeserialization(pso, typeTable) : null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static Collection<string> GetSpecificPropertiesToSerialize(PSObject pso, TypeTable? typeTable) =>
            (s_getSpecificPropertiesToSerialize is not null ? s_getSpecificPropertiesToSerialize(pso, typeTable) : null)
            ?? new Collection<string>();

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfoCollection<T> CreateSmaInternalMemberCollection<T>() where T : PSMemberInfo =>
            InternalCollectionFactory<T>.Create();

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetAdaptedProperties(PSObject pso)
        {
            PSMemberInfoCollection<PSPropertyInfo>? adapted = s_getAdaptedProperties is not null
                ? s_getAdaptedProperties(pso)
                : (s_getAdaptedMembers is not null ? s_getAdaptedMembers(pso) : null);
            return adapted is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(adapted) : null;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetAdaptedMembers(PSObject pso)
        {
            PSMemberInfoCollection<PSPropertyInfo>? adapted = s_getAdaptedMembers is not null ? s_getAdaptedMembers(pso) : null;
            return adapted is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(adapted) : null;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetAdaptedMembers(PSObject pso, PSMemberInfoInternalCollection<PSPropertyInfo>? value)
        {
            s_setAdaptedMembers?.Invoke(pso, value?.InnerCollection ?? CreateSmaInternalMemberCollection<PSPropertyInfo>());
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetClrMembers(PSObject pso)
        {
            PSMemberInfoCollection<PSPropertyInfo>? clr = s_getClrMembers is not null ? s_getClrMembers(pso) : null;
            return clr is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(clr) : null;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetClrMembers(PSObject pso, PSMemberInfoInternalCollection<PSPropertyInfo>? value)
        {
            s_setClrMembers?.Invoke(pso, value?.InnerCollection ?? CreateSmaInternalMemberCollection<PSPropertyInfo>());
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfoCollection<PSMemberInfo> GetInstanceMembers(PSObject pso)
        {
            if (s_getInstanceMembers is not null && s_getInstanceMembers(pso) is PSMemberInfoCollection<PSMemberInfo> instanceMembers)
            {
                return instanceMembers;
            }

            return pso.Members;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSProperty CreateDeserializedPSProperty(string name, object? serializedValue)
        {
            if (s_createPsProperty is not null)
            {
                return s_createPsProperty(name, serializedValue);
            }

            throw new NotSupportedException("PSProperty(string, object) constructor not found.");
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool GetMemberShouldSerialize(PSMemberInfo member) =>
            s_getMemberShouldSerialize is null || s_getMemberShouldSerialize(member);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfo? GetAliasReferencedMember(PSAliasProperty alias) =>
            s_getAliasReferencedMember is not null ? s_getAliasReferencedMember(alias) : null;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static void SetHasInstanceMember(string memberName) =>
            s_setHasInstanceMember?.Invoke(memberName);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static object? ConvertTo(object? valueToConvert, Type resultType, bool recursion, IFormatProvider formatProvider, TypeTable? backupTypeTable)
        {
            if (s_convertTo5 is not null)
            {
                return s_convertTo5(valueToConvert, resultType, recursion, formatProvider, backupTypeTable);
            }

            return LanguagePrimitives.ConvertTo(valueToConvert, resultType, formatProvider);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static string EncryptSecureString(PSRemotingCryptoHelper helper, SecureString ss) =>
            s_cryptoEncrypt is not null
                ? s_cryptoEncrypt(helper, ss)
                : Microsoft.PowerShell.SecureStringHelper.Protect(ss);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static SecureString DecryptSecureString(PSRemotingCryptoHelper helper, string encrypted) =>
            s_cryptoDecrypt is not null
                ? s_cryptoDecrypt(helper, encrypted)
                : Microsoft.PowerShell.SecureStringHelper.Unprotect(encrypted);

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static Collection<CollectionEntry<PSMemberInfo>> GetMemberCollection(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
        {
            if (backupTypeTable is null && viewType == PSMemberViewTypes.Extended)
            {
                return s_defaultExtendedMemberCollection;
            }

            return CreateMemberCollectionCore(viewType, backupTypeTable);
        }

        private static Collection<CollectionEntry<PSMemberInfo>> CreateMemberCollectionCore(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
        {
            object? smaCol = s_psoGetMemberCollection is not null ? s_psoGetMemberCollection(viewType, backupTypeTable) : null;
            return new Collection<CollectionEntry<PSMemberInfo>> { new CollectionEntry<PSMemberInfo>(smaCol) };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static Collection<CollectionEntry<PSPropertyInfo>> GetPropertyCollection(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
        {
            if (backupTypeTable is null && viewType == (PSMemberViewTypes.Extended | PSMemberViewTypes.Adapted | PSMemberViewTypes.Base))
            {
                return s_defaultAllPropertyCollection;
            }

            return CreatePropertyCollectionCore(viewType, backupTypeTable);
        }

        private static Collection<CollectionEntry<PSPropertyInfo>> CreatePropertyCollectionCore(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
        {
            object? smaCol = s_psoGetPropertyCollection is not null ? s_psoGetPropertyCollection(viewType, backupTypeTable) : null;
            return new Collection<CollectionEntry<PSPropertyInfo>> { new CollectionEntry<PSPropertyInfo>(smaCol) };
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetSpecificPropertiesToSerializeFromSma(
            PSObject source,
            Collection<CollectionEntry<PSPropertyInfo>>? allPropertiesCollection,
            TypeTable? typeTable)
        {
            object? smaPropsCol = allPropertiesCollection is { Count: > 0 } ? allPropertiesCollection[0].SmaCollection : null;
            if (s_smaGetSpecificPropertiesToSerialize is not null && smaPropsCol is not null)
            {
                PSMemberInfoCollection<PSPropertyInfo>? res = s_smaGetSpecificPropertiesToSerialize(source, smaPropsCol, typeTable);
                return res is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(res) : null;
            }

            return null;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static IEnumerable<T> MatchExtendedMembers<T>(PSObject source, object? smaCollections, string name, PSMemberTypes memberTypes, int matchOptions)
            where T : PSMemberInfo
        {
            Func<PSObject, object, string, PSMemberTypes, int, IEnumerable<T>>? matcher = ExtendedMemberMatcher<T>.Match;
            if (matcher is not null && smaCollections is not null)
            {
                return matcher(source, smaCollections, name, memberTypes, matchOptions);
            }

            var list = new List<T>();
            PSMemberInfoCollection<PSMemberInfo> instanceMembers = GetInstanceMembers(source);
            foreach (PSMemberInfo member in instanceMembers)
            {
                if (member is T typed && (member.MemberType & memberTypes) != 0 && GetMemberShouldSerialize(member))
                {
                    list.Add(typed);
                }
            }

            return new ReadOnlyCollection<T>(list);
        }

        internal static void ErrorRecordToPSObjectForRemoting(ErrorRecord errorRecord, PSObject dest) =>
            s_errorRecordToPSObjectForRemoting?.Invoke(errorRecord, dest);

        internal static void InformationalRecordToPSObjectForRemoting(InformationalRecord record, PSObject dest) =>
            s_informationalRecordToPSObjectForRemoting?.Invoke(record, dest);

        internal static ScriptBlock CreateScriptBlock(string scriptText, bool isProductCode) =>
            s_scriptBlockCreateDelayParsed is not null
                ? s_scriptBlockCreateDelayParsed(scriptText, isProductCode)
                : ScriptBlock.Create(scriptText);
    }

    internal static class StandaloneSerializationExtensions
    {
        extension(PSObject pso)
        {
            internal bool ImmediateBaseObjectIsEmpty => SmaInternalBridge.GetImmediateBaseObjectIsEmpty(pso);

            internal string? ToStringFromDeserialization
            {
                get => SmaInternalBridge.GetToStringFromDeserialization(pso);
                set => SmaInternalBridge.SetToStringFromDeserialization(pso, value);
            }

            internal string? TokenText
            {
                get => SmaInternalBridge.GetTokenText(pso);
                set => SmaInternalBridge.SetTokenText(pso, value);
            }

            internal bool IsDeserialized
            {
                get => SmaInternalBridge.GetIsDeserialized(pso);
                set => SmaInternalBridge.SetIsDeserialized(pso, value);
            }

            internal ConsolidatedString InternalTypeNames
            {
                get => SmaInternalBridge.GetInternalTypeNames(pso);
                set => SmaInternalBridge.SetInternalTypeNames(pso, value);
            }

            internal AdapterProxy InternalAdapter => SmaInternalBridge.GetInternalAdapter(pso);

            internal static DotNetAdapterProxy DotNetInstanceAdapter => SmaInternalBridge.DotNetInstanceAdapter;

            internal PSMemberInfoInternalCollection<PSPropertyInfo>? AdaptedMembers
            {
                get => SmaInternalBridge.GetAdaptedMembers(pso);
                set => SmaInternalBridge.SetAdaptedMembers(pso, value);
            }

            internal PSMemberInfoInternalCollection<PSPropertyInfo>? ClrMembers
            {
                get => SmaInternalBridge.GetClrMembers(pso);
                set => SmaInternalBridge.SetClrMembers(pso, value);
            }

            internal PSMemberInfoCollection<PSMemberInfo> InstanceMembers
                => SmaInternalBridge.GetInstanceMembers(pso);

            internal void SetCoreOnDeserialization(object? value, bool overrideTypeInfo)
                => SmaInternalBridge.SetCoreOnDeserialization(pso, value, overrideTypeInfo);

            internal static void CopyDeserializerFields(PSObject source, PSObject target)
                => SmaInternalBridge.CopyDeserializerFields(source, target);

            internal SerializationMethod GetSerializationMethod(TypeTable? typeTable)
                => SmaInternalBridge.GetSerializationMethod(pso, typeTable);

            internal int GetSerializationDepth(TypeTable? typeTable)
                => SmaInternalBridge.GetSerializationDepth(pso, typeTable);

            internal PSPropertyInfo? GetStringSerializationSource(TypeTable? typeTable)
                => SmaInternalBridge.GetStringSerializationSource(pso, typeTable);

            internal bool ShouldSerializeAdapter()
                => SmaInternalBridge.ShouldSerializeAdapter(pso);

            internal Type? GetTargetTypeForDeserialization(TypeTable? typeTable)
                => SmaInternalBridge.GetTargetTypeForDeserialization(pso, typeTable);

            internal Collection<string> GetSpecificPropertiesToSerialize(TypeTable? typeTable)
                => SmaInternalBridge.GetSpecificPropertiesToSerialize(pso, typeTable);

            internal PSMemberInfoInternalCollection<PSPropertyInfo>? GetAdaptedProperties()
                => SmaInternalBridge.GetAdaptedProperties(pso);

            internal static Collection<CollectionEntry<PSMemberInfo>> GetMemberCollection(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
                => SmaInternalBridge.GetMemberCollection(viewType, backupTypeTable);

            internal static Collection<CollectionEntry<PSPropertyInfo>> GetPropertyCollection(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
                => SmaInternalBridge.GetPropertyCollection(viewType, backupTypeTable);

            internal static PSObject AsPSObject(object obj, bool preserveToString)
                => SmaInternalBridge.AsPSObjectWithPreserveToString(obj, preserveToString);
        }

        extension(PSMemberInfo member)
        {
            internal bool ShouldSerialize => SmaInternalBridge.GetMemberShouldSerialize(member);
        }

        extension(PSAliasProperty alias)
        {
            internal PSMemberInfo? ReferencedMember => SmaInternalBridge.GetAliasReferencedMember(alias);
        }

        extension(ErrorRecord errorRecord)
        {
            internal void ToPSObjectForRemoting(PSObject dest)
                => SmaInternalBridge.ErrorRecordToPSObjectForRemoting(errorRecord, dest);
        }

        extension(InformationalRecord record)
        {
            internal void ToPSObjectForRemoting(PSObject dest)
                => SmaInternalBridge.InformationalRecordToPSObjectForRemoting(record, dest);
        }

        extension(ScriptBlock)
        {
            internal static ScriptBlock CreateDelayParsedScriptBlock(string script, bool isProductCode)
                => SmaInternalBridge.CreateScriptBlock(script, isProductCode);
        }

        extension(LanguagePrimitives)
        {
            internal static object? ConvertTo(object? valueToConvert, Type resultType, bool recursion, IFormatProvider formatProvider, TypeTable? backupTypeTable)
                => SmaInternalBridge.ConvertTo(valueToConvert, resultType, recursion, formatProvider, backupTypeTable);
        }

        extension(PSRemotingCryptoHelper cryptoHelper)
        {
            internal string EncryptSecureString(SecureString secureString)
                => SmaInternalBridge.EncryptSecureString(cryptoHelper, secureString);

            internal SecureString DecryptSecureString(string encryptedString)
                => SmaInternalBridge.DecryptSecureString(cryptoHelper, encryptedString);
        }

        extension(PSTraceSource)
        {
            internal static InvalidOperationException NewInvalidOperationException(string message)
                => new InvalidOperationException(message);

            internal static ArgumentNullException NewArgumentNullException(string paramName)
                => new ArgumentNullException(paramName);

            internal static ArgumentException NewArgumentException(string paramName)
                => new ArgumentException(paramName);
        }

        extension(PSMemberInfoCollection<PSMemberInfo> members)
        {
            internal int Count
            {
                get
                {
                    int count = 0;
                    foreach (PSMemberInfo _ in members)
                    {
                        count++;
                    }

                    return count;
                }
            }
        }
    }

    internal static class SerializationUtilities
    {
        internal static object? GetPropertyValue(PSObject psObject, string propertyName)
        {
            PSPropertyInfo? property = psObject.Properties[propertyName];
            return property?.Value;
        }

        internal static object? GetPsObjectPropertyBaseObject(PSObject psObject, string propertyName)
        {
            object? val = GetPropertyValue(psObject, propertyName);
            if (val is PSObject propertyPsObject)
            {
                return propertyPsObject.BaseObject;
            }

            return val;
        }

        internal static void GetKnownContainerTypeInfo(
            object source,
            out ContainerType ct,
            out IDictionary? dictionary,
            out IEnumerable? enumerable)
        {
            ct = ContainerType.None;
            dictionary = null;
            enumerable = null;

            dictionary = source as IDictionary;
            if (dictionary is not null)
            {
                ct = ContainerType.Dictionary;
            }
            else if (source is Stack)
            {
                ct = ContainerType.Stack;
                enumerable = LanguagePrimitives.GetEnumerable(source);
            }
            else if (source is Queue)
            {
                ct = ContainerType.Queue;
                enumerable = LanguagePrimitives.GetEnumerable(source);
            }
            else if (source is IList)
            {
                ct = ContainerType.List;
                enumerable = LanguagePrimitives.GetEnumerable(source);
            }
            else
            {
                Type gt = source.GetType();
                if (gt.IsGenericType)
                {
                    if (DerivesFromGenericType(gt, typeof(Stack<>)))
                    {
                        ct = ContainerType.Stack;
                        enumerable = LanguagePrimitives.GetEnumerable(source);
                    }
                    else if (DerivesFromGenericType(gt, typeof(Queue<>)))
                    {
                        ct = ContainerType.Queue;
                        enumerable = LanguagePrimitives.GetEnumerable(source);
                    }
                    else if (DerivesFromGenericType(gt, typeof(List<>)))
                    {
                        ct = ContainerType.List;
                        enumerable = LanguagePrimitives.GetEnumerable(source);
                    }
                }
            }

            if (ct == ContainerType.None)
            {
                try
                {
                    enumerable = LanguagePrimitives.GetEnumerable(source);
                    if (enumerable is not null)
                    {
                        ct = ContainerType.Enumerable;
                    }
                }
                catch
                {
                }
            }

            if (ct == ContainerType.None)
            {
                enumerable = source as IEnumerable;
                if (enumerable is not null)
                {
                    IEnumerator enumerator = enumerable.GetEnumerator();
                    if (enumerator is not null && enumerator.MoveNext())
                    {
                        ct = ContainerType.Enumerable;
                    }
                }
            }
        }

        private static bool DerivesFromGenericType(Type? derived, Type baseType)
        {
            while (derived is not null)
            {
                if (derived.IsGenericType)
                {
                    derived = derived.GetGenericTypeDefinition();
                }

                if (derived == baseType)
                {
                    return true;
                }

                derived = derived.BaseType;
            }

            return false;
        }

        internal static string? GetToString(object source)
        {
            try
            {
                return Convert.ToString(source, CultureInfo.InvariantCulture);
            }
            catch
            {
                return null;
        }
        }

        internal static string? GetToStringForPrimitiveObject(PSObject? pso)
        {
            if (pso is null)
            {
                return null;
            }

            if (pso.ToStringFromDeserialization is not null)
            {
                return pso.ToStringFromDeserialization;
            }

            string? token = pso.TokenText;
            if (token is not null)
            {
                string? originalToString = GetToString(pso.BaseObject);
                if (originalToString is null || !string.Equals(token, originalToString, StringComparison.Ordinal))
                {
                    return token;
                }
            }

            return null;
        }

        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetSpecificPropertiesToSerialize(
            PSObject? source,
            Collection<CollectionEntry<PSPropertyInfo>>? allPropertiesCollection,
            TypeTable? typeTable)
        {
            if (source is null || source.GetSerializationMethod(typeTable) != SerializationMethod.SpecificProperties)
            {
                return null;
            }

            PSMemberInfoInternalCollection<PSPropertyInfo>? fromSma =
                SmaInternalBridge.GetSpecificPropertiesToSerializeFromSma(source, allPropertiesCollection, typeTable);
            if (fromSma is not null)
            {
                return fromSma;
            }

            var specificProperties = new PSMemberInfoInternalCollection<PSPropertyInfo>();
            Collection<string> namesOfPropertiesToSerialize = source.GetSpecificPropertiesToSerialize(typeTable);
            foreach (string propertyName in namesOfPropertiesToSerialize)
            {
                PSPropertyInfo? property = source.Properties[propertyName];
                if (property is not null)
                {
                    specificProperties.Add(property);
                }
            }

            return specificProperties;
        }

        internal static object? GetPropertyValueInThreadSafeManner(
            PSPropertyInfo property,
            bool canUseDefaultRunspaceInThreadSafeManner,
            out bool success)
        {
            if (!property.IsGettable)
            {
                success = false;
                return null;
            }

            if (property is PSAliasProperty alias)
            {
                property = (alias.ReferencedMember as PSPropertyInfo)!;
                if (property is null)
                {
                    success = false;
                    return null;
                }
            }

            if (property is PSScriptProperty && !canUseDefaultRunspaceInThreadSafeManner)
            {
                success = false;
                return null;
            }

            try
            {
                object? value = property.Value;
                success = true;
                return value;
            }
            catch (ExtendedTypeSystemException)
            {
                success = false;
                return null;
            }
        }
    }
}

namespace System.Management.Automation.Language
{
    internal static class PSGetMemberBinder
    {
        internal static void SetHasInstanceMember(string memberName)
            => SmaInternalBridge.SetHasInstanceMember(memberName);
    }
}

namespace Microsoft.PowerShell
{
    using System.Management.Automation;

    internal static class SecureStringHelper
    {
        private static readonly Type? s_helperType =
            typeof(PSObject).Assembly.GetType("Microsoft.PowerShell.SecureStringHelper");
        private static readonly Func<SecureString, string>? s_protect = CompileProtect();
        private static readonly Func<string, SecureString>? s_unprotect = CompileUnprotect();

        private static Func<SecureString, string>? CompileProtect()
        {
            MethodInfo? m = s_helperType?.GetMethod("Protect", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return m is not null ? (Func<SecureString, string>)Delegate.CreateDelegate(typeof(Func<SecureString, string>), m) : null;
        }

        private static Func<string, SecureString>? CompileUnprotect()
        {
            MethodInfo? m = s_helperType?.GetMethod("Unprotect", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return m is not null ? (Func<string, SecureString>)Delegate.CreateDelegate(typeof(Func<string, SecureString>), m) : null;
        }

        internal static string Protect(SecureString input)
        {
            if (s_protect is not null)
            {
                return s_protect(input);
            }

            throw new PlatformNotSupportedException("SecureStringHelper.Protect unavailable.");
        }

        internal static SecureString Unprotect(string input)
        {
            if (s_unprotect is not null)
            {
                return s_unprotect(input);
            }

            throw new PlatformNotSupportedException("SecureStringHelper.Unprotect unavailable.");
        }
    }
}

namespace Microsoft.PowerShell.Commands
{
    using System.Management.Automation;

    internal static class EncodingConversion
    {
        internal const string Unicode = "unicode";
    }

    internal static class PathUtils
    {
        internal static string ResolveFilePath(string filePath, PSCmdlet command, bool isLiteralPath)
        {
            if (isLiteralPath)
            {
                return command.SessionState.Path.GetUnresolvedProviderPathFromPSPath(filePath);
            }

            Collection<PathInfo> resolvedPaths;
            try
            {
                resolvedPaths = command.SessionState.Path.GetResolvedPSPathFromPSPath(filePath);
            }
            catch (ItemNotFoundException)
            {
                return command.SessionState.Path.GetUnresolvedProviderPathFromPSPath(filePath);
            }

            if (resolvedPaths.Count > 1)
            {
                command.ThrowTerminatingError(
                    new ErrorRecord(
                        new IOException($"Path '{filePath}' resolved to multiple files."),
                        "PathResolvedToMultipleFiles",
                        ErrorCategory.InvalidArgument,
                        filePath));
            }

            if (resolvedPaths.Count == 1)
            {
                return resolvedPaths[0].ProviderPath;
            }

            return command.SessionState.Path.GetUnresolvedProviderPathFromPSPath(filePath);
        }

        internal static void MasterStreamOpen(
            PSCmdlet cmdlet,
            string filePath,
            string resolvedEncoding,
            bool defaultEncoding,
            bool Append,
            bool Force,
            bool NoClobber,
            out FileStream fileStream,
            out StreamWriter? streamWriter,
            out FileInfo? readOnlyFileInfo,
            bool isLiteralPath)
        {
            fileStream = null!;
            streamWriter = null;
            readOnlyFileInfo = null;

            string resolvedPath = ResolveFilePath(filePath, cmdlet, isLiteralPath);
            try
            {
                FileMode mode = Append ? FileMode.Append : (NoClobber ? FileMode.CreateNew : FileMode.Create);
                if (Force && File.Exists(resolvedPath) && !NoClobber)
                {
                    FileInfo fi = new(resolvedPath);
                    if ((fi.Attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                    {
                        readOnlyFileInfo = fi;
                        fi.Attributes &= ~FileAttributes.ReadOnly;
                    }
                }

                fileStream = new FileStream(resolvedPath, mode, FileAccess.Write, FileShare.Read);
            }
            catch (IOException e) when (NoClobber && File.Exists(resolvedPath))
            {
                cmdlet.ThrowTerminatingError(
                    new ErrorRecord(
                        e,
                        "NoClobber",
                        ErrorCategory.ResourceExists,
                        resolvedPath));
            }
            catch (Exception e)
            {
                cmdlet.ThrowTerminatingError(
                    new ErrorRecord(
                        e,
                        "FileOpenFailure",
                        ErrorCategory.OpenError,
                        resolvedPath));
            }
        }

        internal static FileStream OpenFileStream(string filePath, PSCmdlet command, bool isLiteralPath)
        {
            string resolvedPath = ResolveFilePath(filePath, command, isLiteralPath);
            try
            {
                return new FileStream(resolvedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
            catch (Exception e)
            {
                command.ThrowTerminatingError(
                    new ErrorRecord(
                        e,
                        "FileOpenFailure",
                        ErrorCategory.OpenError,
                        resolvedPath));
                return null!;
            }
        }
    }
}
