// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

#nullable enable

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
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

    internal sealed class ReferenceIdHandlerForSerializer<T> where T : class
    {
        private readonly IDictionary<T, ulong>? _object2refId;
        private ulong _seed;

        internal ReferenceIdHandlerForSerializer(IDictionary<T, ulong>? dictionary)
        {
            _object2refId = dictionary;
        }

        internal string? SetRefId(T t)
        {
            if (_object2refId is not null)
            {
                ulong refId = _seed++;
                _object2refId.Add(t, refId);
                return refId.ToString(CultureInfo.InvariantCulture);
            }

            return null;
        }

        internal string? GetRefId(T t)
        {
            if (_object2refId is not null && _object2refId.TryGetValue(t, out ulong refId))
            {
                return refId.ToString(CultureInfo.InvariantCulture);
            }

            return null;
        }
    }

    internal sealed class ReferenceIdHandlerForDeserializer<T> where T : class
    {
        private readonly Dictionary<string, T> _refId2object = new(StringComparer.Ordinal);

        internal void SetRefId(T o, string refId, bool duplicateRefIdsAllowed)
        {
            if (_refId2object.ContainsKey(refId) && !duplicateRefIdsAllowed)
            {
                throw new XmlException($"Duplicate RefId: {refId}");
            }

            _refId2object[refId] = o;
        }

        internal T? GetReferencedObject(string refId)
        {
            _refId2object.TryGetValue(refId, out T? result);
            return result;
        }
    }

    internal sealed class WeakReferenceDictionary<T> : IDictionary<object, T>
    {
        private readonly ConditionalWeakTable<object, StrongBox<T>> _table = new();

        public void Add(object key, T value)
        {
            _table.Add(key, new StrongBox<T>(value));
        }

        public bool ContainsKey(object key) => _table.TryGetValue(key, out _);

        public ICollection<object> Keys => throw new NotSupportedException();

        public bool Remove(object key) => _table.Remove(key);

        public bool TryGetValue(object key, out T value)
        {
            if (_table.TryGetValue(key, out StrongBox<T>? box) && box is not null)
            {
                value = box.Value!;
                return true;
            }

            value = default!;
            return false;
        }

        public ICollection<T> Values => throw new NotSupportedException();

        public T this[object key]
        {
            get => TryGetValue(key, out T val) ? val : throw new KeyNotFoundException();
            set
            {
                _table.Remove(key);
                _table.Add(key, new StrongBox<T>(value));
            }
        }

        public void Add(KeyValuePair<object, T> item) => Add(item.Key, item.Value);
        public void Clear() => _table.Clear();
        public bool Contains(KeyValuePair<object, T> item) => TryGetValue(item.Key, out _);
        public void CopyTo(KeyValuePair<object, T>[] array, int arrayIndex) => throw new NotSupportedException();
        public int Count => 0;
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

    internal sealed class PSMemberInfoInternalCollection<T> : Collection<T> where T : PSMemberInfo
    {
        private readonly Dictionary<string, T> _byName = new(StringComparer.OrdinalIgnoreCase);
        private readonly object? _smaBackingCollection;
        private readonly MethodInfo? _smaBackingAdd;

        internal PSMemberInfoInternalCollection()
        {
        }

        internal PSMemberInfoInternalCollection(object? smaBackingCollection)
        {
            _smaBackingCollection = smaBackingCollection;
            if (smaBackingCollection is not null)
            {
                _smaBackingAdd = smaBackingCollection.GetType().GetMethod("Add", new[] { typeof(T), typeof(bool) })
                    ?? smaBackingCollection.GetType().GetMethod("Add", new[] { typeof(T) });

                if (smaBackingCollection is IEnumerable enumerable)
                {
                    foreach (object? item in enumerable)
                    {
                        if (item is T typed)
                        {
                            base.InsertItem(Count, typed);
                            if (typed.Name is not null)
                            {
                                _byName[typed.Name] = typed;
                            }
                        }
                    }
                }
            }
        }

        public T? this[string name]
        {
            get
            {
                _byName.TryGetValue(name, out T? item);
                return item;
            }
        }

        protected override void InsertItem(int index, T item)
        {
            if (item is null)
            {
                return;
            }

            if (_smaBackingCollection is not null && _smaBackingAdd is not null)
            {
                if (_smaBackingAdd.GetParameters().Length == 2)
                {
                    _smaBackingAdd.Invoke(_smaBackingCollection, new object[] { item, false });
                }
                else
                {
                    _smaBackingAdd.Invoke(_smaBackingCollection, new object[] { item });
                }
            }

            base.InsertItem(index, item);
            if (item.Name is not null)
            {
                _byName[item.Name] = item;
            }
        }
    }

    internal sealed class PSMemberInfoIntegratingCollection<T> where T : PSMemberInfo
    {
        private readonly PSObject _source;
        private readonly object? _smaCollections;

        internal PSMemberInfoIntegratingCollection(PSObject source, Collection<CollectionEntry<T>>? collections)
        {
            _source = source;
            _smaCollections = collections is { Count: > 0 } ? collections[0].SmaCollection : null;
        }

        internal IEnumerable<T> Match(string name, PSMemberTypes memberTypes, MshMemberMatchOptions matchOptions)
        {
            return SmaInternalBridge.MatchExtendedMembers<T>(_source, _smaCollections, name, memberTypes, (int)matchOptions);
        }
    }

    internal sealed class ConsolidatedString : Collection<string>
    {
        internal static readonly ConsolidatedString Empty = new(Array.Empty<string>());
        internal static readonly IEqualityComparer<ConsolidatedString> EqualityComparer = new ConsolidatedStringEqualityComparer();

        internal ConsolidatedString(IEnumerable<string> strings)
        {
            if (strings is not null)
            {
                foreach (string s in strings)
                {
                    if (!string.IsNullOrEmpty(s))
                    {
                        Add(s);
                    }
                }
            }
        }

        internal ConsolidatedString(ConsolidatedString other)
            : this((IEnumerable<string>)other)
        {
        }

        internal string Key => string.Join(" ", this);

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

            public int GetHashCode(ConsolidatedString obj) => obj.Key.GetHashCode(StringComparison.Ordinal);
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
        private static readonly MethodInfo? s_loadUnsafeXmlDocument =
            typeof(PSObject).Assembly.GetType("System.Management.Automation.InternalDeserializer")
                ?.GetMethod("LoadUnsafeXmlDocument", BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public, null, new[] { typeof(string), typeof(bool), typeof(int?) }, null);

        internal static XmlDocument LoadUnsafeXmlDocument(
            string xmlText,
            bool preserveNonElements,
            int? maxCharactersInDocument)
        {
            if (s_loadUnsafeXmlDocument is not null)
            {
                return (XmlDocument)s_loadUnsafeXmlDocument.Invoke(null, new object?[] { xmlText, preserveNonElements, maxCharactersInDocument })!;
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

    internal sealed class AdapterProxy
    {
        private readonly object? _adapter;
        private readonly MethodInfo? _baseGetTypeNameHierarchy;

        internal AdapterProxy(object? adapter)
        {
            _adapter = adapter;
            _baseGetTypeNameHierarchy = adapter?.GetType().GetMethod(
                "BaseGetTypeNameHierarchy",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        }

        internal Collection<string> BaseGetTypeNameHierarchy(object? baseObject)
        {
            if (_adapter is not null && _baseGetTypeNameHierarchy is not null && baseObject is not null)
            {
                if (_baseGetTypeNameHierarchy.Invoke(_adapter, new[] { baseObject }) is IEnumerable<string> names)
                {
                    var list = new Collection<string>();
                    foreach (string s in names)
                    {
                        list.Add(s);
                    }

                    return list;
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
        private readonly MethodInfo? _getDotNetMethod;

        internal DotNetAdapterProxy(object? dotNetAdapter)
        {
            _dotNetAdapter = dotNetAdapter;
            if (dotNetAdapter is not null)
            {
                foreach (MethodInfo m in dotNetAdapter.GetType().GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    if (m.Name == "GetDotNetMethod" && m.IsGenericMethodDefinition)
                    {
                        _getDotNetMethod = m;
                        break;
                    }
                }
            }
        }

        internal T? GetDotNetMethod<T>(object obj, string methodName) where T : PSMemberInfo
        {
            if (_dotNetAdapter is not null && _getDotNetMethod is not null)
            {
                return _getDotNetMethod.MakeGenericMethod(typeof(T)).Invoke(_dotNetAdapter, new[] { obj, methodName }) as T;
            }

            return null;
        }
    }

    internal static class SmaInternalBridge
    {
        private static readonly Assembly s_sma = typeof(PSObject).Assembly;
        private const BindingFlags AllInstance = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        private const BindingFlags AllStatic = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

        private static readonly PropertyInfo? s_psoImmediateBaseObjectIsEmpty =
            typeof(PSObject).GetProperty("ImmediateBaseObjectIsEmpty", AllInstance);
        private static readonly PropertyInfo? s_psoToStringFromDeserialization =
            typeof(PSObject).GetProperty("ToStringFromDeserialization", AllInstance);
        private static readonly PropertyInfo? s_psoTokenText =
            typeof(PSObject).GetProperty("TokenText", AllInstance);
        private static readonly PropertyInfo? s_psoIsDeserialized =
            typeof(PSObject).GetProperty("IsDeserialized", AllInstance);
        private static readonly MethodInfo? s_psoAsPSObjectPreserveToString =
            typeof(PSObject).GetMethod("AsPSObject", AllStatic, null, new[] { typeof(object), typeof(bool) }, null);
        private static readonly MethodInfo? s_psoSetCoreOnDeserialization =
            typeof(PSObject).GetMethod("SetCoreOnDeserialization", AllInstance);
        private static readonly MethodInfo? s_psoGetSerializationMethod =
            typeof(PSObject).GetMethod("GetSerializationMethod", AllInstance);
        private static readonly MethodInfo? s_psoGetSerializationDepth =
            typeof(PSObject).GetMethod("GetSerializationDepth", AllInstance);
        private static readonly MethodInfo? s_psoGetStringSerializationSource =
            typeof(PSObject).GetMethod("GetStringSerializationSource", AllInstance);
        private static readonly MethodInfo? s_psoShouldSerializeAdapter =
            typeof(PSObject).GetMethod("ShouldSerializeAdapter", AllInstance);
        private static readonly MethodInfo? s_psoGetTargetTypeForDeserialization =
            typeof(PSObject).GetMethod("GetTargetTypeForDeserialization", AllInstance);
        private static readonly MethodInfo? s_psoGetSpecificPropertiesToSerialize =
            typeof(PSObject).GetMethod("GetSpecificPropertiesToSerialize", AllInstance);
        private static readonly MethodInfo? s_psoCopyDeserializerFields =
            typeof(PSObject).GetMethod("CopyDeserializerFields", AllStatic);
        private static readonly PropertyInfo? s_psoAdaptedMembers =
            typeof(PSObject).GetProperty("AdaptedMembers", AllInstance);
        private static readonly PropertyInfo? s_psoClrMembers =
            typeof(PSObject).GetProperty("ClrMembers", AllInstance);
        private static readonly PropertyInfo? s_psoInstanceMembers =
            typeof(PSObject).GetProperty("InstanceMembers", AllInstance);
        private static readonly PropertyInfo? s_psoInternalAdapter =
            typeof(PSObject).GetProperty("InternalAdapter", AllInstance);
        private static readonly FieldInfo? s_psoDotNetInstanceAdapterField =
            typeof(PSObject).GetField("DotNetInstanceAdapter", AllStatic);
        private static readonly PropertyInfo? s_psoDotNetInstanceAdapterProp =
            typeof(PSObject).GetProperty("DotNetInstanceAdapter", AllStatic);

        private static readonly PropertyInfo? s_memberShouldSerialize =
            typeof(PSMemberInfo).GetProperty("ShouldSerialize", AllInstance);
        private static readonly PropertyInfo? s_aliasReferencedMember =
            typeof(PSAliasProperty).GetProperty("ReferencedMember", AllInstance);
        private static readonly ConstructorInfo? s_psPropertyCtor =
            typeof(PSProperty).GetConstructor(AllInstance, null, new[] { typeof(string), typeof(object) }, null);

        private static readonly MethodInfo? s_errorRecordToPSObjectForRemoting =
            typeof(ErrorRecord).GetMethod("ToPSObjectForRemoting", AllInstance, null, new[] { typeof(PSObject) }, null);
        private static readonly MethodInfo? s_informationalRecordToPSObjectForRemoting =
            typeof(InformationalRecord).GetMethod("ToPSObjectForRemoting", AllInstance, null, new[] { typeof(PSObject) }, null);
        private static readonly MethodInfo? s_scriptBlockCreateDelayParsed =
            typeof(ScriptBlock).GetMethod("CreateDelayParsedScriptBlock", AllStatic, null, new[] { typeof(string), typeof(bool) }, null);

        private static readonly MethodInfo? s_setHasInstanceMember =
            s_sma.GetType("System.Management.Automation.Language.PSGetMemberBinder")
                ?.GetMethod("SetHasInstanceMember", AllStatic);

        private static readonly MethodInfo? s_convertTo5 =
            typeof(LanguagePrimitives).GetMethod(
                "ConvertTo",
                AllStatic,
                null,
                new[] { typeof(object), typeof(Type), typeof(bool), typeof(IFormatProvider), typeof(TypeTable) },
                null);

        private static readonly MethodInfo? s_cryptoEncrypt =
            typeof(PSRemotingCryptoHelper).GetMethod("EncryptSecureString", AllInstance);
        private static readonly MethodInfo? s_cryptoDecrypt =
            typeof(PSRemotingCryptoHelper).GetMethod("DecryptSecureString", AllInstance);

        private static readonly Type? s_smaInternalCollectionGeneric =
            s_sma.GetType("System.Management.Automation.PSMemberInfoInternalCollection`1");
        private static readonly Type? s_smaIntegratingCollectionGeneric =
            s_sma.GetType("System.Management.Automation.PSMemberInfoIntegratingCollection`1");
        private static readonly Type? s_mshMemberMatchOptionsType =
            s_sma.GetType("System.Management.Automation.MshMemberMatchOptions");
        private static readonly MethodInfo? s_psoGetMemberCollection =
            typeof(PSObject).GetMethod("GetMemberCollection", AllStatic, null, new[] { typeof(PSMemberViewTypes), typeof(TypeTable) }, null);
        private static readonly MethodInfo? s_psoGetPropertyCollection =
            typeof(PSObject).GetMethod("GetPropertyCollection", AllStatic, null, new[] { typeof(PSMemberViewTypes), typeof(TypeTable) }, null);
        private static readonly MethodInfo? s_smaGetSpecificPropertiesToSerialize =
            s_sma.GetType("System.Management.Automation.SerializationUtilities")
                ?.GetMethod("GetSpecificPropertiesToSerialize", AllStatic);

        private static readonly DotNetAdapterProxy s_dotNetAdapterProxy = new(
            s_psoDotNetInstanceAdapterField?.GetValue(null) ?? s_psoDotNetInstanceAdapterProp?.GetValue(null));

        internal static DotNetAdapterProxy DotNetInstanceAdapter => s_dotNetAdapterProxy;

        internal static AdapterProxy GetInternalAdapter(PSObject pso)
            => new AdapterProxy(s_psoInternalAdapter?.GetValue(pso));

        internal static bool GetImmediateBaseObjectIsEmpty(PSObject pso)
        {
            if (s_psoImmediateBaseObjectIsEmpty is not null)
            {
                return (bool)s_psoImmediateBaseObjectIsEmpty.GetValue(pso)!;
            }

            return pso.ImmediateBaseObject is PSCustomObject;
        }

        internal static string? GetToStringFromDeserialization(PSObject pso)
            => s_psoToStringFromDeserialization?.GetValue(pso) as string;

        internal static void SetToStringFromDeserialization(PSObject pso, string? value)
            => s_psoToStringFromDeserialization?.SetValue(pso, value);

        internal static string? GetTokenText(PSObject pso)
            => s_psoTokenText?.GetValue(pso) as string;

        internal static void SetTokenText(PSObject pso, string? value)
            => s_psoTokenText?.SetValue(pso, value);

        internal static bool GetIsDeserialized(PSObject pso)
            => s_psoIsDeserialized is not null && (bool)s_psoIsDeserialized.GetValue(pso)!;

        internal static void SetIsDeserialized(PSObject pso, bool value)
            => s_psoIsDeserialized?.SetValue(pso, value);

        internal static void SetCoreOnDeserialization(PSObject pso, object? value, bool overrideTypeInfo)
            => s_psoSetCoreOnDeserialization?.Invoke(pso, new object?[] { value, overrideTypeInfo });

        internal static void CopyDeserializerFields(PSObject source, PSObject target)
            => s_psoCopyDeserializerFields?.Invoke(null, new object[] { source, target });

        internal static PSObject AsPSObjectWithPreserveToString(object obj, bool preserveToString)
        {
            if (s_psoAsPSObjectPreserveToString is not null)
            {
                return (PSObject)s_psoAsPSObjectPreserveToString.Invoke(null, new object?[] { obj, preserveToString })!;
            }

            return PSObject.AsPSObject(obj);
        }

        internal static SerializationMethod GetSerializationMethod(PSObject pso, TypeTable? typeTable)
        {
            if (s_psoGetSerializationMethod is not null)
            {
                object? res = s_psoGetSerializationMethod.Invoke(pso, new object?[] { typeTable });
                if (res is not null)
                {
                    return (SerializationMethod)(int)res;
                }
            }

            return SerializationMethod.AllPublicProperties;
        }

        internal static int GetSerializationDepth(PSObject pso, TypeTable? typeTable)
        {
            if (s_psoGetSerializationDepth is not null)
            {
                object? res = s_psoGetSerializationDepth.Invoke(pso, new object?[] { typeTable });
                if (res is int depth)
                {
                    return depth;
                }
            }

            return 0;
        }

        internal static PSPropertyInfo? GetStringSerializationSource(PSObject pso, TypeTable? typeTable)
            => s_psoGetStringSerializationSource?.Invoke(pso, new object?[] { typeTable }) as PSPropertyInfo;

        internal static bool ShouldSerializeAdapter(PSObject pso)
        {
            if (s_psoShouldSerializeAdapter is not null)
            {
                object? res = s_psoShouldSerializeAdapter.Invoke(pso, Array.Empty<object>());
                if (res is bool b)
                {
                    return b;
                }
            }

            return true;
        }

        internal static Type? GetTargetTypeForDeserialization(PSObject pso, TypeTable? typeTable)
            => s_psoGetTargetTypeForDeserialization?.Invoke(pso, new object?[] { typeTable }) as Type;

        internal static Collection<string> GetSpecificPropertiesToSerialize(PSObject pso, TypeTable? typeTable)
        {
            if (s_psoGetSpecificPropertiesToSerialize is not null)
            {
                if (s_psoGetSpecificPropertiesToSerialize.Invoke(pso, new object?[] { typeTable }) is Collection<string> col)
                {
                    return col;
                }
            }

            return new Collection<string>();
        }

        internal static object? CreateSmaInternalPropertyCollection()
        {
            if (s_smaInternalCollectionGeneric is not null)
            {
                Type closed = s_smaInternalCollectionGeneric.MakeGenericType(typeof(PSPropertyInfo));
                return Activator.CreateInstance(closed, nonPublic: true);
            }

            return null;
        }

        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetAdaptedProperties(PSObject pso)
        {
            MethodInfo? m = typeof(PSObject).GetMethod("GetAdaptedProperties", AllInstance);
            object? adapted = m is not null ? m.Invoke(pso, Array.Empty<object>()) : s_psoAdaptedMembers?.GetValue(pso);
            return adapted is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(adapted) : null;
        }

        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetAdaptedMembers(PSObject pso)
        {
            object? adapted = s_psoAdaptedMembers?.GetValue(pso);
            return adapted is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(adapted) : null;
        }

        internal static void SetAdaptedMembers(PSObject pso, PSMemberInfoInternalCollection<PSPropertyInfo>? value)
        {
            object? backing = CreateSmaInternalPropertyCollection();
            s_psoAdaptedMembers?.SetValue(pso, backing);
        }

        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetClrMembers(PSObject pso)
        {
            object? clr = s_psoClrMembers?.GetValue(pso);
            return clr is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(clr) : null;
        }

        internal static void SetClrMembers(PSObject pso, PSMemberInfoInternalCollection<PSPropertyInfo>? value)
        {
            object? backing = CreateSmaInternalPropertyCollection();
            s_psoClrMembers?.SetValue(pso, backing);
        }

        internal static PSMemberInfoCollection<PSMemberInfo> GetInstanceMembers(PSObject pso)
        {
            if (s_psoInstanceMembers?.GetValue(pso) is PSMemberInfoCollection<PSMemberInfo> instanceMembers)
            {
                return instanceMembers;
            }

            return pso.Members;
        }

        internal static PSProperty CreateDeserializedPSProperty(string name, object? serializedValue)
        {
            if (s_psPropertyCtor is not null)
            {
                return (PSProperty)s_psPropertyCtor.Invoke(new object?[] { name, serializedValue });
            }

            throw new NotSupportedException("PSProperty(string, object) constructor not found.");
        }

        internal static bool GetMemberShouldSerialize(PSMemberInfo member)
        {
            if (s_memberShouldSerialize is not null && s_memberShouldSerialize.GetValue(member) is bool b)
            {
                return b;
            }

            return true;
        }

        internal static PSMemberInfo? GetAliasReferencedMember(PSAliasProperty alias)
            => s_aliasReferencedMember?.GetValue(alias) as PSMemberInfo;

        internal static void SetHasInstanceMember(string memberName)
            => s_setHasInstanceMember?.Invoke(null, new object[] { memberName });

        internal static object? ConvertTo(object? valueToConvert, Type resultType, bool recursion, IFormatProvider formatProvider, TypeTable? backupTypeTable)
        {
            if (s_convertTo5 is not null)
            {
                try
                {
                    return s_convertTo5.Invoke(null, new object?[] { valueToConvert, resultType, recursion, formatProvider, backupTypeTable });
                }
                catch (TargetInvocationException tie) when (tie.InnerException is not null)
                {
                    throw tie.InnerException;
                }
            }

            return LanguagePrimitives.ConvertTo(valueToConvert, resultType, formatProvider);
        }

        internal static string EncryptSecureString(PSRemotingCryptoHelper helper, SecureString ss)
        {
            if (s_cryptoEncrypt is not null)
            {
                return (string)s_cryptoEncrypt.Invoke(helper, new object[] { ss })!;
            }

            return Microsoft.PowerShell.SecureStringHelper.Protect(ss);
        }

        internal static SecureString DecryptSecureString(PSRemotingCryptoHelper helper, string encrypted)
        {
            if (s_cryptoDecrypt is not null)
            {
                return (SecureString)s_cryptoDecrypt.Invoke(helper, new object[] { encrypted })!;
            }

            return Microsoft.PowerShell.SecureStringHelper.Unprotect(encrypted);
        }

        internal static Collection<CollectionEntry<PSMemberInfo>> GetMemberCollection(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
        {
            object? smaCol = s_psoGetMemberCollection?.Invoke(null, new object?[] { viewType, backupTypeTable });
            return new Collection<CollectionEntry<PSMemberInfo>> { new CollectionEntry<PSMemberInfo>(smaCol) };
        }

        internal static Collection<CollectionEntry<PSPropertyInfo>> GetPropertyCollection(PSMemberViewTypes viewType, TypeTable? backupTypeTable)
        {
            object? smaCol = s_psoGetPropertyCollection?.Invoke(null, new object?[] { viewType, backupTypeTable });
            return new Collection<CollectionEntry<PSPropertyInfo>> { new CollectionEntry<PSPropertyInfo>(smaCol) };
        }

        internal static PSMemberInfoInternalCollection<PSPropertyInfo>? GetSpecificPropertiesToSerializeFromSma(
            PSObject source,
            Collection<CollectionEntry<PSPropertyInfo>>? allPropertiesCollection,
            TypeTable? typeTable)
        {
            object? smaPropsCol = allPropertiesCollection is { Count: > 0 } ? allPropertiesCollection[0].SmaCollection : null;
            if (s_smaGetSpecificPropertiesToSerialize is not null && smaPropsCol is not null)
            {
                object? res = s_smaGetSpecificPropertiesToSerialize.Invoke(null, new object?[] { source, smaPropsCol, typeTable });
                return res is not null ? new PSMemberInfoInternalCollection<PSPropertyInfo>(res) : null;
            }

            return null;
        }

        internal static IEnumerable<T> MatchExtendedMembers<T>(PSObject source, object? smaCollections, string name, PSMemberTypes memberTypes, int matchOptions)
            where T : PSMemberInfo
        {
            if (s_smaIntegratingCollectionGeneric is not null && smaCollections is not null && s_mshMemberMatchOptionsType is not null)
            {
                Type closedType = s_smaIntegratingCollectionGeneric.MakeGenericType(typeof(T));
                object? integrating = Activator.CreateInstance(
                    closedType,
                    AllInstance,
                    null,
                    new object?[] { source, smaCollections },
                    null);
                MethodInfo? matchMethod = closedType.GetMethod(
                    "Match",
                    AllInstance,
                    null,
                    new[] { typeof(string), typeof(PSMemberTypes), s_mshMemberMatchOptionsType },
                    null);
                if (integrating is not null && matchMethod is not null)
                {
                    object matchOptionsEnum = Enum.ToObject(s_mshMemberMatchOptionsType, matchOptions);
                    if (matchMethod.Invoke(integrating, new object[] { name, memberTypes, matchOptionsEnum }) is IEnumerable<T> matched)
                    {
                        return matched;
                    }
                }
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

        internal static ReadOnlyPSMemberInfoCollection<T> CreateReadOnlyCollection<T>(PSMemberInfoInternalCollection<T> members)
            where T : PSMemberInfo
        {
            ConstructorInfo? ctor = typeof(ReadOnlyPSMemberInfoCollection<T>).GetConstructors(AllInstance)[0];
            return (ReadOnlyPSMemberInfoCollection<T>)ctor.Invoke(new object?[] { null });
        }

        internal static void ErrorRecordToPSObjectForRemoting(ErrorRecord errorRecord, PSObject dest)
            => s_errorRecordToPSObjectForRemoting?.Invoke(errorRecord, new object[] { dest });

        internal static void InformationalRecordToPSObjectForRemoting(InformationalRecord record, PSObject dest)
            => s_informationalRecordToPSObjectForRemoting?.Invoke(record, new object[] { dest });

        internal static ScriptBlock CreateScriptBlock(string scriptText, bool isProductCode)
        {
            if (s_scriptBlockCreateDelayParsed is not null)
            {
                if (s_scriptBlockCreateDelayParsed.Invoke(null, new object[] { scriptText, isProductCode }) is ScriptBlock sb)
                {
                    return sb;
                }
            }

            return ScriptBlock.Create(scriptText);
        }
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
                get => new ConsolidatedString(pso.TypeNames);
                set
                {
                    pso.TypeNames.Clear();
                    if (value is not null)
                    {
                        foreach (string t in value)
                        {
                            pso.TypeNames.Add(t);
                        }
                    }
                }
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
            if (source is null)
            {
                return null;
            }

            PSMemberInfoInternalCollection<PSPropertyInfo>? fromSma =
                SmaInternalBridge.GetSpecificPropertiesToSerializeFromSma(source, allPropertiesCollection, typeTable);
            if (fromSma is not null)
            {
                return fromSma;
            }

            if (source.GetSerializationMethod(typeTable) == SerializationMethod.SpecificProperties)
            {
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

            return null;
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
        private static readonly MethodInfo? s_protect =
            s_helperType?.GetMethod("Protect", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
        private static readonly MethodInfo? s_unprotect =
            s_helperType?.GetMethod("Unprotect", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);

        internal static string Protect(SecureString input)
        {
            if (s_protect is not null)
            {
                return (string)s_protect.Invoke(null, new object[] { input })!;
            }

            throw new PlatformNotSupportedException("SecureStringHelper.Protect unavailable.");
        }

        internal static SecureString Unprotect(string input)
        {
            if (s_unprotect is not null)
            {
                return (SecureString)s_unprotect.Invoke(null, new object[] { input })!;
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
