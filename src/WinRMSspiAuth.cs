using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Management.Automation;
using System.Net;
using System.Runtime.InteropServices;

namespace AwakeCoding.PSRemoting.PowerShell
{
    internal interface IWinRMNegotiateAuthContext : IDisposable
    {
        string HttpAuthLabel { get; }
        bool Complete { get; }
        byte[]? Step(byte[]? incomingToken);
    }

    internal static class WinRMSspiProviderFactory
    {
        private static readonly object s_syncRoot = new();
        private static WinRMSspiProvider? s_devolutionsProvider;
        private static WinRMSspiProvider? s_windowsProvider;

        public static IWinRMNegotiateAuthContext CreateContext(WinRMClientInfo connectionInfo, Uri endpoint)
        {
            if (connectionInfo == null) throw new ArgumentNullException(nameof(connectionInfo));
            if (endpoint == null) throw new ArgumentNullException(nameof(endpoint));
            if (connectionInfo.Credential == null && !connectionInfo.UseImplicitCredential)
            {
                throw new InvalidOperationException("WinRM Negotiate/Kerberos requires explicit credentials unless -UseImplicitCredential is specified.");
            }

            WinRMSspiProvider provider = GetDefaultProvider();
            string package = connectionInfo.Authentication == WinRMAuthenticationMechanism.Kerberos ? "Kerberos" : "Negotiate";
            string headerLabel = connectionInfo.Authentication == WinRMAuthenticationMechanism.Kerberos ? "Kerberos" : "Negotiate";
            string targetName = $"host/{endpoint.Host}";

            return new WinRMSspiAuthContext(provider, package, headerLabel, targetName, connectionInfo.Credential);
        }

        private static WinRMSspiProvider GetDefaultProvider()
        {
            try
            {
                return GetOrLoadDevolutionsProvider();
            }
            catch (Exception ex) when (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] Devolutions.Sspi unavailable, falling back to Windows SSPI: {ex.Message}");
                return GetOrLoadWindowsProvider();
            }
        }

        private static WinRMSspiProvider GetOrLoadDevolutionsProvider()
        {
            lock (s_syncRoot)
            {
                s_devolutionsProvider ??= WinRMSspiProvider.Load(
                    "Devolutions.Sspi",
                    GetDevolutionsLibraryCandidates(),
                    required: true);

                return s_devolutionsProvider;
            }
        }

        private static WinRMSspiProvider GetOrLoadWindowsProvider()
        {
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                throw new PlatformNotSupportedException("The Windows SSPI backend is only available on Windows.");
            }

            lock (s_syncRoot)
            {
                s_windowsProvider ??= WinRMSspiProvider.Load(
                    "Windows SSPI",
                    new[] { "Secur32.dll" },
                    required: true);

                return s_windowsProvider;
            }
        }

        private static IEnumerable<string> GetDevolutionsLibraryCandidates()
        {
            string assemblyDirectory = Path.GetDirectoryName(typeof(WinRMSspiProviderFactory).Assembly.Location) ?? string.Empty;
            string osName;
            string extension;
            string prefix = string.Empty;

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                osName = "win";
                extension = "dll";
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            {
                osName = "osx";
                extension = "dylib";
                prefix = "lib";
            }
            else
            {
                osName = "linux";
                extension = "so";
                prefix = "lib";
            }

            string architecture = RuntimeInformation.ProcessArchitecture.ToString().ToLowerInvariant();
            string fileName = $"{prefix}DevolutionsSspi.{extension}";

            yield return Path.Combine(assemblyDirectory, "runtimes", $"{osName}-{architecture}", "native", fileName);
            yield return Path.Combine(assemblyDirectory, fileName);
            yield return fileName;
            yield return "DevolutionsSspi";
        }
    }

    internal sealed class WinRMSspiAuthContext : IWinRMNegotiateAuthContext
    {
        private readonly WinRMSspiProvider _provider;
        private readonly WinRMSspiCredentialHandle _credential;
        private readonly string _targetName;
        private WinRMSspiContextHandle? _context;

        public string HttpAuthLabel { get; }
        public bool Complete { get; private set; }

        public WinRMSspiAuthContext(
            WinRMSspiProvider provider,
            string package,
            string httpAuthLabel,
            string targetName,
            PSCredential? credential)
        {
            _provider = provider ?? throw new ArgumentNullException(nameof(provider));
            _credential = _provider.AcquireCredentials(package, credential);
            _targetName = targetName ?? throw new ArgumentNullException(nameof(targetName));
            HttpAuthLabel = httpAuthLabel ?? throw new ArgumentNullException(nameof(httpAuthLabel));
        }

        public byte[]? Step(byte[]? incomingToken)
        {
            var result = _provider.InitializeSecurityContext(
                _credential,
                _context,
                _targetName,
                CreateContextFlags(),
                incomingToken);

            _context = result.Context;
            Complete = !result.MoreNeeded;
            return result.OutputToken;
        }

        public void Dispose()
        {
            _context?.Dispose();
            _credential.Dispose();
        }

        private static InitiatorContextRequestFlags CreateContextFlags()
        {
            return InitiatorContextRequestFlags.ISC_REQ_CONNECTION
                | InitiatorContextRequestFlags.ISC_REQ_MUTUAL_AUTH
                | InitiatorContextRequestFlags.ISC_REQ_REPLAY_DETECT
                | InitiatorContextRequestFlags.ISC_REQ_SEQUENCE_DETECT
                | InitiatorContextRequestFlags.ISC_REQ_CONFIDENTIALITY
                | InitiatorContextRequestFlags.ISC_REQ_INTEGRITY;
        }
    }

    internal sealed class WinRMSspiProvider : IDisposable
    {
        private const int ContinueNeeded = unchecked((int)0x00090312);
        private readonly IntPtr _module;
        private readonly Dictionary<string, Delegate> _delegateCache = new(StringComparer.Ordinal);

        private WinRMSspiProvider(IntPtr module)
        {
            _module = module;
        }

        public static WinRMSspiProvider Load(string displayName, IEnumerable<string> candidates, bool required)
        {
            List<string> attempted = new();
            foreach (string candidate in candidates)
            {
                attempted.Add(candidate);
                if (NativeLibrary.TryLoad(candidate, out IntPtr module))
                {
                    return new WinRMSspiProvider(module);
                }
            }

            if (!required)
            {
                return new WinRMSspiProvider(IntPtr.Zero);
            }

            throw new DllNotFoundException(
                $"Failed to find required {displayName} library. Searched: '{string.Join("', '", attempted)}'. " +
                "Ensure the module's runtimes assets are present.");
        }

        public WinRMSspiCredentialHandle AcquireCredentials(string package, PSCredential? credential)
        {
            string? username = credential?.UserName;
            string? domain = null;
            string? password = credential?.GetNetworkCredential().Password;

            if (!string.IsNullOrEmpty(username) && username.Contains('\\'))
            {
                string[] parts = username.Split('\\', 2);
                domain = parts[0];
                username = parts[1];
            }

            IntPtr userPtr = IntPtr.Zero;
            IntPtr domainPtr = IntPtr.Zero;
            IntPtr passwordPtr = IntPtr.Zero;
            IntPtr authDataPtr = IntPtr.Zero;

            try
            {
                if (username != null || domain != null || password != null)
                {
                    userPtr = username != null ? Marshal.StringToHGlobalUni(username) : IntPtr.Zero;
                    domainPtr = domain != null ? Marshal.StringToHGlobalUni(domain) : IntPtr.Zero;
                    passwordPtr = password != null ? Marshal.StringToHGlobalUni(password) : IntPtr.Zero;
                    authDataPtr = Marshal.AllocHGlobal(Marshal.SizeOf<SEC_WINNT_AUTH_IDENTITY_W>());

                    var authData = new SEC_WINNT_AUTH_IDENTITY_W
                    {
                        User = userPtr,
                        UserLength = (uint)(username?.Length ?? 0),
                        Domain = domainPtr,
                        DomainLength = (uint)(domain?.Length ?? 0),
                        Password = passwordPtr,
                        PasswordLength = (uint)(password?.Length ?? 0),
                        Flags = WinNTAuthIdentityFlags.SEC_WINNT_AUTH_IDENTITY_UNICODE
                    };

                    Marshal.StructureToPtr(authData, authDataPtr, false);
                }

                var handle = new WinRMSspiCredentialHandle(this);
                int status = GetDelegate<AcquireCredentialsHandleDelegate>("AcquireCredentialsHandleW")(
                    null,
                    package,
                    (uint)CredentialUse.SECPKG_CRED_OUTBOUND,
                    IntPtr.Zero,
                    authDataPtr,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    handle.Pointer,
                    out _);

                if (status != 0)
                {
                    handle.Dispose();
                    throw CreateSspiException(status, "AcquireCredentialsHandle");
                }

                handle.MarkOwned();
                return handle;
            }
            finally
            {
                if (authDataPtr != IntPtr.Zero)
                {
                    Marshal.DestroyStructure<SEC_WINNT_AUTH_IDENTITY_W>(authDataPtr);
                    Marshal.FreeHGlobal(authDataPtr);
                }

                if (userPtr != IntPtr.Zero) Marshal.FreeHGlobal(userPtr);
                if (domainPtr != IntPtr.Zero) Marshal.FreeHGlobal(domainPtr);
                if (passwordPtr != IntPtr.Zero) Marshal.FreeHGlobal(passwordPtr);
            }
        }

        public WinRMSspiContextResult InitializeSecurityContext(
            WinRMSspiCredentialHandle credential,
            WinRMSspiContextHandle? context,
            string targetName,
            InitiatorContextRequestFlags flags,
            byte[]? inputToken)
        {
            flags |= InitiatorContextRequestFlags.ISC_REQ_ALLOCATE_MEMORY;

            IntPtr inputTokenPtr = IntPtr.Zero;
            IntPtr inputBufferPtr = IntPtr.Zero;
            IntPtr inputBufferDescPtr = IntPtr.Zero;
            IntPtr outputBufferPtr = IntPtr.Zero;
            IntPtr outputBufferDescPtr = IntPtr.Zero;
            WinRMSspiContextHandle? outputContext = null;

            try
            {
                if (inputToken != null && inputToken.Length > 0)
                {
                    inputTokenPtr = Marshal.AllocHGlobal(inputToken.Length);
                    Marshal.Copy(inputToken, 0, inputTokenPtr, inputToken.Length);

                    inputBufferPtr = Marshal.AllocHGlobal(Marshal.SizeOf<SecBuffer>());
                    Marshal.StructureToPtr(new SecBuffer
                    {
                        cbBuffer = (uint)inputToken.Length,
                        BufferType = (uint)SecBufferType.SECBUFFER_TOKEN,
                        pvBuffer = inputTokenPtr
                    }, inputBufferPtr, false);

                    inputBufferDescPtr = Marshal.AllocHGlobal(Marshal.SizeOf<SecBufferDesc>());
                    Marshal.StructureToPtr(new SecBufferDesc
                    {
                        ulVersion = 0,
                        cBuffers = 1,
                        pBuffers = inputBufferPtr
                    }, inputBufferDescPtr, false);
                }

                outputBufferPtr = Marshal.AllocHGlobal(Marshal.SizeOf<SecBuffer>());
                Marshal.StructureToPtr(new SecBuffer
                {
                    cbBuffer = 0,
                    BufferType = (uint)SecBufferType.SECBUFFER_TOKEN,
                    pvBuffer = IntPtr.Zero
                }, outputBufferPtr, false);

                outputBufferDescPtr = Marshal.AllocHGlobal(Marshal.SizeOf<SecBufferDesc>());
                Marshal.StructureToPtr(new SecBufferDesc
                {
                    ulVersion = 0,
                    cBuffers = 1,
                    pBuffers = outputBufferPtr
                }, outputBufferDescPtr, false);

                outputContext = context ?? new WinRMSspiContextHandle(this);
                IntPtr inputContext = context?.Pointer ?? IntPtr.Zero;
                IntPtr newContext = outputContext.Pointer;

                int status = GetDelegate<InitializeSecurityContextDelegate>("InitializeSecurityContextW")(
                    credential.Pointer,
                    inputContext,
                    targetName,
                    flags,
                    0,
                    (uint)TargetDataRep.SECURITY_NATIVE_DREP,
                    inputBufferDescPtr,
                    0,
                    newContext,
                    outputBufferDescPtr,
                    out _,
                    out _);

                if (status != 0 && status != ContinueNeeded)
                {
                    if (context == null)
                    {
                        outputContext.Dispose();
                    }

                    throw CreateSspiException(status, "InitializeSecurityContext");
                }

                outputContext.MarkOwned();
                var outputBuffer = Marshal.PtrToStructure<SecBuffer>(outputBufferPtr);
                byte[]? outputToken = null;
                if (outputBuffer.cbBuffer > 0 && outputBuffer.pvBuffer != IntPtr.Zero)
                {
                    outputToken = new byte[outputBuffer.cbBuffer];
                    Marshal.Copy(outputBuffer.pvBuffer, outputToken, 0, outputToken.Length);
                    GetDelegate<FreeContextBufferDelegate>("FreeContextBuffer")(outputBuffer.pvBuffer);
                }

                return new WinRMSspiContextResult(outputContext, outputToken, status == ContinueNeeded);
            }
            finally
            {
                if (inputBufferDescPtr != IntPtr.Zero)
                {
                    Marshal.DestroyStructure<SecBufferDesc>(inputBufferDescPtr);
                    Marshal.FreeHGlobal(inputBufferDescPtr);
                }

                if (inputBufferPtr != IntPtr.Zero)
                {
                    Marshal.DestroyStructure<SecBuffer>(inputBufferPtr);
                    Marshal.FreeHGlobal(inputBufferPtr);
                }

                if (inputTokenPtr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(inputTokenPtr);
                }

                if (outputBufferDescPtr != IntPtr.Zero)
                {
                    Marshal.DestroyStructure<SecBufferDesc>(outputBufferDescPtr);
                    Marshal.FreeHGlobal(outputBufferDescPtr);
                }

                if (outputBufferPtr != IntPtr.Zero)
                {
                    Marshal.DestroyStructure<SecBuffer>(outputBufferPtr);
                    Marshal.FreeHGlobal(outputBufferPtr);
                }
            }
        }

        public void ReleaseCredential(IntPtr handlePointer)
        {
            GetDelegate<FreeCredentialsHandleDelegate>("FreeCredentialsHandle")(handlePointer);
        }

        public void ReleaseContext(IntPtr handlePointer)
        {
            GetDelegate<DeleteSecurityContextDelegate>("DeleteSecurityContext")(handlePointer);
        }

        public void Dispose()
        {
            if (_module != IntPtr.Zero)
            {
                NativeLibrary.Free(_module);
            }
        }

        private T GetDelegate<T>(string exportName)
            where T : Delegate
        {
            if (_module == IntPtr.Zero)
            {
                throw new InvalidOperationException("The SSPI provider library is not loaded.");
            }

            if (!_delegateCache.TryGetValue(exportName, out Delegate? value))
            {
                IntPtr export = NativeLibrary.GetExport(_module, exportName);
                value = Marshal.GetDelegateForFunctionPointer(export, typeof(T));
                _delegateCache[exportName] = value;
            }

            return (T)value;
        }

        private static Exception CreateSspiException(int status, string operation)
        {
            return new InvalidOperationException(
                $"{operation} failed ({new Win32Exception(status).Message}, 0x{status:X8}).");
        }

        [UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Unicode)]
        private delegate int AcquireCredentialsHandleDelegate(
            string? principal,
            string package,
            uint credentialUse,
            IntPtr logonId,
            IntPtr authData,
            IntPtr getKeyFn,
            IntPtr getKeyArgument,
            IntPtr credentialHandle,
            out SECURITY_INTEGER expiry);

        [UnmanagedFunctionPointer(CallingConvention.Winapi, CharSet = CharSet.Unicode)]
        private delegate int InitializeSecurityContextDelegate(
            IntPtr credentialHandle,
            IntPtr contextHandle,
            string targetName,
            InitiatorContextRequestFlags contextReq,
            uint reserved1,
            uint targetDataRep,
            IntPtr input,
            uint reserved2,
            IntPtr newContext,
            IntPtr output,
            out InitiatorContextReturnFlags contextAttributes,
            out SECURITY_INTEGER expiry);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int FreeCredentialsHandleDelegate(IntPtr credentialHandle);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int DeleteSecurityContextDelegate(IntPtr contextHandle);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int FreeContextBufferDelegate(IntPtr contextBuffer);
    }

    internal sealed class WinRMSspiContextResult
    {
        public WinRMSspiContextResult(WinRMSspiContextHandle context, byte[]? outputToken, bool moreNeeded)
        {
            Context = context;
            OutputToken = outputToken;
            MoreNeeded = moreNeeded;
        }

        public WinRMSspiContextHandle Context { get; }
        public byte[]? OutputToken { get; }
        public bool MoreNeeded { get; }
    }

    internal sealed class WinRMSspiCredentialHandle : IDisposable
    {
        private readonly WinRMSspiProvider _provider;
        private bool _ownsHandle;

        public WinRMSspiCredentialHandle(WinRMSspiProvider provider)
        {
            _provider = provider;
            Pointer = Marshal.AllocHGlobal(Marshal.SizeOf<SecHandle>());
            Marshal.StructureToPtr(new SecHandle(), Pointer, false);
        }

        public IntPtr Pointer { get; private set; }

        public void MarkOwned() => _ownsHandle = true;

        public void Dispose()
        {
            if (Pointer != IntPtr.Zero)
            {
                if (_ownsHandle)
                {
                    _provider.ReleaseCredential(Pointer);
                }

                Marshal.DestroyStructure<SecHandle>(Pointer);
                Marshal.FreeHGlobal(Pointer);
                Pointer = IntPtr.Zero;
            }
        }
    }

    internal sealed class WinRMSspiContextHandle : IDisposable
    {
        private readonly WinRMSspiProvider _provider;
        private bool _ownsHandle;

        public WinRMSspiContextHandle(WinRMSspiProvider provider)
        {
            _provider = provider;
            Pointer = Marshal.AllocHGlobal(Marshal.SizeOf<SecHandle>());
            Marshal.StructureToPtr(new SecHandle(), Pointer, false);
        }

        public IntPtr Pointer { get; private set; }

        public void MarkOwned() => _ownsHandle = true;

        public void Dispose()
        {
            if (Pointer != IntPtr.Zero)
            {
                if (_ownsHandle)
                {
                    _provider.ReleaseContext(Pointer);
                }

                Marshal.DestroyStructure<SecHandle>(Pointer);
                Marshal.FreeHGlobal(Pointer);
                Pointer = IntPtr.Zero;
            }
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SEC_WINNT_AUTH_IDENTITY_W
    {
        public IntPtr User;
        public uint UserLength;
        public IntPtr Domain;
        public uint DomainLength;
        public IntPtr Password;
        public uint PasswordLength;
        public WinNTAuthIdentityFlags Flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SECURITY_INTEGER
    {
        public uint LowPart;
        public int HighPart;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SecBufferDesc
    {
        public uint ulVersion;
        public uint cBuffers;
        public IntPtr pBuffers;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SecBuffer
    {
        public uint cbBuffer;
        public uint BufferType;
        public IntPtr pvBuffer;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct SecHandle
    {
        public IntPtr dwLower;
        public IntPtr dwUpper;
    }

    internal enum CredentialUse : uint
    {
        SECPKG_CRED_OUTBOUND = 0x00000002
    }

    [Flags]
    internal enum InitiatorContextRequestFlags : uint
    {
        ISC_REQ_MUTUAL_AUTH = 0x00000002,
        ISC_REQ_REPLAY_DETECT = 0x00000004,
        ISC_REQ_SEQUENCE_DETECT = 0x00000008,
        ISC_REQ_CONFIDENTIALITY = 0x00000010,
        ISC_REQ_ALLOCATE_MEMORY = 0x00000100,
        ISC_REQ_CONNECTION = 0x00000800,
        ISC_REQ_INTEGRITY = 0x00010000
    }

    [Flags]
    internal enum InitiatorContextReturnFlags : uint
    {
        ISC_RET_MUTUAL_AUTH = 0x00000002
    }

    internal enum TargetDataRep : uint
    {
        SECURITY_NATIVE_DREP = 0x00000010
    }

    internal enum SecBufferType : uint
    {
        SECBUFFER_TOKEN = 2
    }

    internal enum WinNTAuthIdentityFlags : uint
    {
        SEC_WINNT_AUTH_IDENTITY_UNICODE = 2
    }
}
