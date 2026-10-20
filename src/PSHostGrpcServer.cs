using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Protobuf;
using Grpc.Core;
using AwakeCoding.PSRemoting.PowerShell.Grpc;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// gRPC-based PowerShell remoting server.
    /// </summary>
    public sealed class PSHostGrpcServer : PSHostServerBase
    {
        private const int BufferSize = 8192;

        private Server? _grpcServer;

        public bool UseSecureConnection { get; private set; }

        public bool AllowUnencrypted { get; private set; }

        public PSHostGrpcEncoding Encoding { get; private set; }

        public string GrpcUri
        {
            get
            {
                string scheme = UseSecureConnection ? "grpcs" : "grpc";
                return $"{scheme}://{ListenAddress}:{Port}";
            }
        }

        public PSHostGrpcServer(
            string name,
            int port,
            string listenAddress,
            bool useSecureConnection,
            bool allowUnencrypted,
            int maxConnections,
            int drainTimeout,
            PSHostGrpcEncoding encoding = PSHostGrpcEncoding.Protobuf)
            : base(name, port, listenAddress, maxConnections, drainTimeout)
        {
            UseSecureConnection = useSecureConnection;
            AllowUnencrypted = allowUnencrypted;
            Encoding = encoding;
        }

        public override void StartListenerAsync()
        {
            PSHostGrpcPlatform.EnsureSupported();

            if (State == ServerState.Running || State == ServerState.Starting)
            {
                throw new InvalidOperationException($"Server '{Name}' is already {State}");
            }

            if (UseSecureConnection)
            {
                throw new NotSupportedException("gRPC TLS server credentials are not implemented yet. Use loopback plaintext gRPC for this module transport.");
            }

            if (!AllowUnencrypted && !IsLoopbackAddress(ListenAddress))
            {
                throw new InvalidOperationException("Plaintext gRPC is allowed only on loopback by default. Specify -AllowUnencrypted to bind plaintext gRPC to a non-loopback address.");
            }

            State = ServerState.Starting;

            try
            {
                _serverInstance.CancellationTokenSource = new CancellationTokenSource();

                if (Port == 0)
                {
                    Port = GetAvailableTcpPort(ListenAddress);
                }

                var service = new PSHostGrpcTransportService(this, _serverInstance.CancellationTokenSource.Token);
                var serverPort = new ServerPort(
                    NormalizeListenAddress(ListenAddress),
                    Port,
                    ServerCredentials.Insecure);

                _grpcServer = new Server
                {
                    Services = { PSHostGrpcTransport.BindService(service) },
                    Ports = { serverPort }
                };

                _grpcServer.Start();
                State = ServerState.Running;
            }
            catch (Exception ex)
            {
                State = ServerState.Failed;
                LastError = ex;
                throw;
            }
        }

        public override void StopListenerAsync(bool force)
        {
            if (!TryBeginStopping())
            {
                return;
            }

            if (State != ServerState.Running && State != ServerState.Starting)
            {
                ResetStoppingFlag();
                return;
            }

            State = ServerState.Stopping;

            try
            {
                _serverInstance.CancellationTokenSource?.Cancel();

                if (!force && ConnectionCount > 0)
                {
                    var drainStart = DateTime.UtcNow;
                    var timeout = TimeSpan.FromSeconds(DrainTimeout);

                    while (ConnectionCount > 0 && (DateTime.UtcNow - drainStart) < timeout)
                    {
                        Thread.Sleep(100);
                    }
                }

                var connections = _serverInstance.ActiveConnections.Values.ToArray();
                bool hadRemainingConnections = connections.Length > 0;

                _serverInstance.ActiveConnections.Clear();

                if (_grpcServer != null)
                {
                    if (force || hadRemainingConnections)
                    {
                        _grpcServer.KillAsync().Wait(ListenerThreadJoinTimeoutMs);
                    }
                    else
                    {
                        _grpcServer.ShutdownAsync().Wait(ListenerThreadJoinTimeoutMs);
                    }
                }

                Unregister();
                State = ServerState.Stopped;
            }
            catch (Exception ex)
            {
                Unregister();
                State = ServerState.Failed;
                LastError = ex;
                throw;
            }
            finally
            {
                _serverInstance.CancellationTokenSource?.Dispose();
                _serverInstance.CancellationTokenSource = null;
                _grpcServer = null;
                ResetStoppingFlag();
            }
        }

        protected override void AcceptConnectionAsync()
        {
            throw new NotImplementedException("gRPC connections are accepted by the gRPC server runtime.");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _grpcServer != null)
            {
                try
                {
                    _grpcServer.KillAsync().Wait(ListenerThreadJoinTimeoutMs);
                }
                catch { }
            }

            base.Dispose(disposing);
        }

        private static bool IsLoopbackAddress(string listenAddress)
        {
            if (listenAddress.Equals("localhost", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IPAddress.TryParse(listenAddress, out var address) && IPAddress.IsLoopback(address);
        }

        private static string NormalizeListenAddress(string listenAddress)
        {
            if (listenAddress == "*" || listenAddress == "+")
            {
                return "0.0.0.0";
            }

            return listenAddress;
        }

        private static int GetAvailableTcpPort(string listenAddress)
        {
            IPAddress address;
            if (listenAddress == "*" || listenAddress == "+" || listenAddress == "0.0.0.0")
            {
                address = IPAddress.Any;
            }
            else if (listenAddress.Equals("localhost", StringComparison.OrdinalIgnoreCase))
            {
                address = IPAddress.Loopback;
            }
            else if (!IPAddress.TryParse(listenAddress, out address!))
            {
                address = IPAddress.Loopback;
            }

            var listener = new TcpListener(address, 0);
            listener.Start();
            try
            {
                return ((IPEndPoint)listener.LocalEndpoint).Port;
            }
            finally
            {
                listener.Stop();
            }
        }

        private sealed class PSHostGrpcTransportService : PSHostGrpcTransport.PSHostGrpcTransportBase
        {
            private static readonly BindingFlags s_anyFlags = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

            private static readonly Type? s_outOfProcessTextWriterType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.OutOfProcessTextWriter");
            private static readonly Type? s_outOfProcessServerSessionTmType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.Server.OutOfProcessServerSessionTransportManager");
            private static readonly Type? s_abstractServerSessionTmType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.Server.AbstractServerSessionTransportManager");
            private static readonly Type? s_abstractServerTmType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.Server.AbstractServerTransportManager");
            private static readonly Type? s_serverRemoteSessionType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.ServerRemoteSession");
            private static readonly Type? s_serverRunspacePoolDriverType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.ServerRunspacePoolDriver");
            private static readonly Type? s_serverRunspacePoolDsHandlerType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.ServerRunspacePoolDataStructureHandler");
            private static readonly Type? s_serverPowerShellDriverType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.ServerPowerShellDriver");
            private static readonly Type? s_serverPowerShellDsHandlerType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.ServerPowerShellDataStructureHandler");
            private static readonly Type? s_serverRemoteHostType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.ServerRemoteHost");
            private static readonly Type? s_serverMethodExecutorType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.ServerMethodExecutor");
            private static readonly Type? s_serverRemoteDebuggerType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.ServerRemoteDebugger");
            private static readonly Type? s_remotingEncoderType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.RemotingEncoder");
            private static readonly Type? s_remotingDecoderType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.RemotingDecoder");
            private static readonly Type? s_remoteStreamOptionsType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.RemoteStreamOptions");
            private static readonly Type? s_rdoPsObjectType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.RemoteDataObject`1")?.MakeGenericType(typeof(PSObject));
            private static readonly Type? s_remoteDataEventArgsRdoPsObjectType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.RemoteDataEventArgs`1") != null && s_rdoPsObjectType != null
                ? typeof(PSObject).Assembly.GetType("System.Management.Automation.RemoteDataEventArgs`1")!.MakeGenericType(s_rdoPsObjectType)
                : null;

            private static readonly Type? s_baseTransportManagerType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.BaseTransportManager");
            private static readonly Type? s_psIdentityType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.PSIdentity");
            private static readonly Type? s_psPrincipalType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.PSPrincipal");
            private static readonly Type? s_psSenderInfoType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.PSSenderInfo");
            private static readonly Type? s_psCertificateDetailsType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Remoting.PSCertificateDetails");

            private static readonly Type? s_psRemotingCryptoHelperServerType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.Internal.PSRemotingCryptoHelperServer");
            private static readonly ConstructorInfo? s_psRemotingCryptoHelperServerCtor = s_psRemotingCryptoHelperServerType?.GetConstructor(s_anyFlags, null, Type.EmptyTypes, null);
            private static readonly ConstructorInfo? s_outOfProcessSessionTmCtor = s_outOfProcessServerSessionTmType?.GetConstructors(s_anyFlags).FirstOrDefault();
            private static readonly ConstructorInfo? s_psIdentityCtor = s_psIdentityType != null && s_psCertificateDetailsType != null
                ? s_psIdentityType.GetConstructor(s_anyFlags, null, new[] { typeof(string), typeof(bool), typeof(string), s_psCertificateDetailsType }, null)
                : null;
            private static readonly ConstructorInfo? s_psPrincipalCtor = s_psPrincipalType?.GetConstructors(s_anyFlags).FirstOrDefault();
            private static readonly ConstructorInfo? s_psSenderInfoCtor = s_psSenderInfoType != null && s_psPrincipalType != null
                ? s_psSenderInfoType.GetConstructor(s_anyFlags, null, new[] { s_psPrincipalType, typeof(string) }, null)
                : null;
            private static readonly MethodInfo? s_createServerRemoteSessionMethod = s_serverRemoteSessionType?.GetMethod(
                "CreateServerRemoteSession",
                s_anyFlags);

            private static readonly MethodInfo? s_createCommandTransportManagerMethod = s_outOfProcessServerSessionTmType?.GetMethod("CreateCommandTransportManager", s_anyFlags);
            private static readonly MethodInfo? s_getCommandTransportManagerMethod = s_outOfProcessServerSessionTmType?.GetMethod("GetCommandTransportManager", s_anyFlags);
            private static readonly MethodInfo? s_removeCommandTransportManagerMethod = s_outOfProcessServerSessionTmType?.GetMethod("RemoveCommandTransportManager", s_anyFlags);
            private static readonly MethodInfo? s_closeTransportManagerMethod = s_abstractServerTmType?.GetMethod("Close", s_anyFlags);
            private static readonly MethodInfo? s_onDataAvailableCallbackMethod = s_baseTransportManagerType?.GetMethod("OnDataAvailableCallback", s_anyFlags);

            private static readonly FieldInfo? s_runspacePoolDriverField = s_serverRemoteSessionType?.GetField("_runspacePoolDriver", s_anyFlags);
            private static readonly FieldInfo? s_rsToUseForSteppablePipelineField = s_serverRunspacePoolDriverType?.GetField("_rsToUseForSteppablePipeline", s_anyFlags);
            private static readonly PropertyInfo? s_rpDriverDsHandlerProp = s_serverRunspacePoolDriverType?.GetProperty("DataStructureHandler", s_anyFlags);
            private static readonly PropertyInfo? s_rpDriverInstanceIdProp = s_serverRunspacePoolDriverType?.GetProperty("InstanceId", s_anyFlags);
            private static readonly PropertyInfo? s_rpDriverServerRemoteHostProp = s_serverRunspacePoolDriverType?.GetProperty("ServerRemoteHost", s_anyFlags);
            private static readonly FieldInfo? s_rpDriverServerRemoteDebuggerField = s_serverRunspacePoolDriverType?.GetField("_serverRemoteDebugger", s_anyFlags);
            private static readonly FieldInfo? s_rpDriverInputCollectionField = s_serverRunspacePoolDriverType?.GetField("_inputCollection", s_anyFlags);
            private static readonly FieldInfo? s_createAndInvokePowerShellField = s_serverRunspacePoolDsHandlerType?.GetField("CreateAndInvokePowerShell", s_anyFlags);

            private static readonly PropertyInfo? s_remoteHostIsRunspacePushedProp = s_serverRemoteHostType?.GetProperty("IsRunspacePushed", s_anyFlags);
            private static readonly PropertyInfo? s_remoteHostServerMethodExecutorProp = s_serverRemoteHostType?.GetProperty("ServerMethodExecutor", s_anyFlags);
            private static readonly MethodInfo? s_abortAllCallsMethod = s_serverMethodExecutorType?.GetMethod("AbortAllCalls", s_anyFlags);
            private static readonly PropertyInfo? s_debuggerInBreakpointProp = s_serverRemoteDebuggerType?.GetProperty("InBreakpoint", s_anyFlags);

            private static readonly ConstructorInfo? s_serverPowerShellDriverCtor = s_serverPowerShellDriverType?.GetConstructors(s_anyFlags)
                .FirstOrDefault(c => c.GetParameters().Length == 11);
            private static readonly FieldInfo? s_psDriverLocalOutputField = s_serverPowerShellDriverType?.GetField("_localPowerShellOutput", s_anyFlags);
            private static readonly FieldInfo? s_psDriverSyncObjectField = s_serverPowerShellDriverType?.GetField("_syncObject", s_anyFlags);
            private static readonly PropertyInfo? s_psDriverRemoteStreamOptionsProp = s_serverPowerShellDriverType?.GetProperty("RemoteStreamOptions", s_anyFlags);
            private static readonly FieldInfo? s_psDriverRemoteHostField = s_serverPowerShellDriverType?.GetField("_remoteHost", s_anyFlags);
            private static readonly PropertyInfo? s_psDriverLocalPowerShellProp = s_serverPowerShellDriverType?.GetProperty("LocalPowerShell", s_anyFlags);
            private static readonly PropertyInfo? s_psDriverInputCollectionProp = s_serverPowerShellDriverType?.GetProperty("InputCollection", s_anyFlags);
            private static readonly PropertyInfo? s_psDriverDsHandlerProp = s_serverPowerShellDriverType?.GetProperty("DataStructureHandler", s_anyFlags);
            private static readonly MethodInfo? s_psDriverUnregisterOutputHandlersMethod = s_serverPowerShellDriverType?.GetMethod("UnregisterPipelineOutputEventHandlers", s_anyFlags);
            private static readonly MethodInfo? s_psDriverUnregisterPsHandlersMethod = s_serverPowerShellDriverType?.GetMethod("UnregisterPowerShellEventHandlers", s_anyFlags);
            private static readonly MethodInfo? s_psDriverUnregisterDsHandlersMethod = s_serverPowerShellDriverType?.GetMethod("UnregisterDataStructureHandlerEventHandlers", s_anyFlags);
            private static readonly MethodInfo? s_psDriverStartMethod = s_serverPowerShellDriverType?.GetMethod("Start", s_anyFlags, null, Type.EmptyTypes, null);
            private static readonly MethodInfo? s_psDsHandlerRaiseRemoveAssocMethod = s_serverPowerShellDsHandlerType?.GetMethod("RaiseRemoveAssociationEvent", s_anyFlags);

            private static readonly PropertyInfo? s_psRunningExtraCommandsProp = typeof(System.Management.Automation.PowerShell).GetProperty("RunningExtraCommands", s_anyFlags);
            private static readonly MethodInfo? s_getHostInfoMethod = s_remotingDecoderType?.GetMethod("GetHostInfo", s_anyFlags);
            private static readonly MethodInfo? s_getPowerShellMethod = s_remotingDecoderType?.GetMethod("GetPowerShell", s_anyFlags);
            private static readonly MethodInfo? s_getNoInputMethod = s_remotingDecoderType?.GetMethod("GetNoInput", s_anyFlags);
            private static readonly MethodInfo? s_getApartmentStateMethod = s_remotingDecoderType?.GetMethod("GetApartmentState", s_anyFlags);
            private static readonly MethodInfo? s_getRemoteStreamOptionsMethod = s_remotingDecoderType?.GetMethod("GetRemoteStreamOptions", s_anyFlags);
            private static readonly MethodInfo? s_getAddToHistoryMethod = s_remotingDecoderType?.GetMethod("GetAddToHistory", s_anyFlags);
            private static readonly MethodInfo? s_getIsNestedMethod = s_remotingDecoderType?.GetMethod("GetIsNested", s_anyFlags);

            private static readonly Type? s_remotingDataType = typeof(PSObject).Assembly.GetType(
                "System.Management.Automation.RemotingDataType");
            private static readonly object? s_warningDataTypeEnum = s_remotingDataType != null ? Enum.ToObject(s_remotingDataType, 0x41009) : null;
            private static readonly object? s_verboseDataTypeEnum = s_remotingDataType != null ? Enum.ToObject(s_remotingDataType, 0x41008) : null;
            private static readonly object? s_debugDataTypeEnum = s_remotingDataType != null ? Enum.ToObject(s_remotingDataType, 0x41007) : null;
            private const uint PowerShellOutputDataType = 0x41004u;

            private static readonly MethodInfo? s_genErrorRecordMethod = s_remotingEncoderType?.GetMethod("GeneratePowerShellError", s_anyFlags);
            private static readonly MethodInfo? s_genInformationalMethod = s_remotingDataType != null
                ? s_remotingEncoderType?.GetMethod("GeneratePowerShellInformational", s_anyFlags, null, new[] { typeof(object), typeof(Guid), typeof(Guid), s_remotingDataType }, null)
                : null;
            private static readonly MethodInfo? s_genProgressMethod = s_remotingEncoderType?.GetMethod(
                "GeneratePowerShellInformational", s_anyFlags, null, new[] { typeof(ProgressRecord), typeof(Guid), typeof(Guid) }, null);
            private static readonly MethodInfo? s_genInformationMethod = s_remotingEncoderType?.GetMethod(
                "GeneratePowerShellInformational", s_anyFlags, null, new[] { typeof(InformationRecord), typeof(Guid), typeof(Guid) }, null);
            private static readonly MethodInfo? s_genStateInfoMethod = s_remotingEncoderType?.GetMethod("GeneratePowerShellStateInfo", s_anyFlags);

            private static readonly PropertyInfo? s_preserveInvocationInfoOnceProp = typeof(ErrorRecord).GetProperty(
                "PreserveInvocationInfoOnce", s_anyFlags);
            private static readonly PropertyInfo? s_serializeExtendedInfoProp = typeof(InformationalRecord).GetProperty(
                "SerializeExtendedInfo", s_anyFlags);

            private static readonly PropertyInfo? s_rdoEventArgsDataProp = s_remoteDataEventArgsRdoPsObjectType?.GetProperty("Data", s_anyFlags);
            private static readonly PropertyInfo? s_rdoPsObjectDataProp = s_rdoPsObjectType?.GetProperty("Data", s_anyFlags);
            private static readonly PropertyInfo? s_rdoPsObjectPowerShellIdProp = s_rdoPsObjectType?.GetProperty("PowerShellId", s_anyFlags);

            private readonly PSHostGrpcServer _owner;
            private readonly CancellationToken _serverToken;

            public PSHostGrpcTransportService(PSHostGrpcServer owner, CancellationToken serverToken)
            {
                _owner = owner;
                _serverToken = serverToken;
            }

            public override async Task Connect(
                IAsyncStreamReader<PSHostGrpcFrame> requestStream,
                IServerStreamWriter<PSHostGrpcFrame> responseStream,
                ServerCallContext context)
            {
                if (_owner.IsMaxConnectionsReached())
                {
                    throw new RpcException(new Status(StatusCode.ResourceExhausted, "Server at maximum capacity"));
                }

                var connectionId = Guid.NewGuid().ToString();
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(_serverToken, context.CancellationToken);
                using var writeLock = new SemaphoreSlim(1, 1);

                var clientAddress = context.Peer ?? "grpc-client";
                _owner.AddConnection(new ConnectionDetails(connectionId, clientAddress, Environment.ProcessId));

                PSHostGrpcEncoding connectionEncoding = ResolveConnectionEncoding(context, _owner.Encoding);
                bool clientRequestedEncoding = HasExplicitEncodingHeader(context);
                var outboundCodec = new PSHostGrpcPsrpCodec(connectionEncoding);
                var inboundCodec = new PSHostGrpcPsrpCodec(connectionEncoding);

                object? sessionTM = null;
                try
                {
                    var stdOutWriter = new GrpcServerTextWriter(responseStream, writeLock, outboundCodec, isError: false, linkedCts.Token);
                    var stdErrWriter = new GrpcServerTextWriter(responseStream, writeLock, outboundCodec, isError: true, linkedCts.Token);

                    object stdOutWrapper = CreateOutOfProcessTextWriter(stdOutWriter);
                    object stdErrWrapper = CreateOutOfProcessTextWriter(stdErrWriter);
                    object? cryptoHelper = s_psRemotingCryptoHelperServerCtor?.Invoke(Array.Empty<object>());

                    sessionTM = s_outOfProcessSessionTmCtor!.Invoke(new object?[] { stdOutWrapper, stdErrWrapper, cryptoHelper });
                    object senderInfo = CreateLocalSenderInfo();
                    object serverSession = s_createServerRemoteSessionMethod!.Invoke(
                        null,
                        new object?[] { senderInfo, "Microsoft.PowerShell", string.Empty, sessionTM, null, null, null, null })!;
                    bool hookedRunspacePoolDriver = false;

                    while (await requestStream.MoveNext(linkedCts.Token).ConfigureAwait(false))
                    {
                        var frame = requestStream.Current;
                        if (frame.Kind == PSHostGrpcFrameKind.Close)
                        {
                            break;
                        }

                        switch (frame.Kind)
                        {
                            case PSHostGrpcFrameKind.Data:
                                {
                                    if (!clientRequestedEncoding)
                                    {
                                        outboundCodec.Encoding = PSHostGrpcPsrpCodec.FromWireEncoding(frame.PayloadEncoding);
                                    }

                                    bool hasObject = inboundCodec.TryDecodeDataFrameToRemoteDataObject(
                                        frame,
                                        out Guid psGuid,
                                        out object? remoteDataObject);

                                    if (hasObject && remoteDataObject != null)
                                    {
                                        object? targetTM = psGuid == Guid.Empty
                                            ? sessionTM
                                            : s_getCommandTransportManagerMethod?.Invoke(sessionTM, new object[] { psGuid });

                                        targetTM ??= sessionTM;
                                        s_onDataAvailableCallbackMethod?.Invoke(targetTM, new object[] { remoteDataObject });

                                        if (!hookedRunspacePoolDriver && psGuid == Guid.Empty)
                                        {
                                            hookedRunspacePoolDriver = TryHookRunspacePoolDriver(
                                                serverSession,
                                                sessionTM,
                                                responseStream,
                                                writeLock,
                                                outboundCodec,
                                                linkedCts.Token);
                                        }
                                    }

                                    string psGuidString = string.IsNullOrEmpty(frame.PsGuid) ? Guid.Empty.ToString() : frame.PsGuid;
                                    await WriteFrameAsync(
                                        responseStream,
                                        writeLock,
                                        new PSHostGrpcFrame
                                        {
                                            Kind = PSHostGrpcFrameKind.DataAck,
                                            PsGuid = psGuidString
                                        },
                                        linkedCts.Token).ConfigureAwait(false);
                                }
                                break;

                            case PSHostGrpcFrameKind.Command:
                                {
                                    Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                    s_createCommandTransportManagerMethod?.Invoke(sessionTM, new object[] { psGuid });
                                }
                                break;

                            case PSHostGrpcFrameKind.Signal:
                                {
                                    Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                    string psGuidString = psGuid.ToString();
                                    if (psGuid != Guid.Empty)
                                    {
                                        object? cmdTM = s_getCommandTransportManagerMethod?.Invoke(sessionTM, new object[] { psGuid });
                                        if (cmdTM != null)
                                        {
                                            s_closeTransportManagerMethod?.Invoke(cmdTM, new object?[] { null });
                                        }
                                    }

                                    await WriteFrameAsync(
                                        responseStream,
                                        writeLock,
                                        new PSHostGrpcFrame
                                        {
                                            Kind = PSHostGrpcFrameKind.SignalAck,
                                            PsGuid = psGuidString
                                        },
                                        linkedCts.Token).ConfigureAwait(false);
                                }
                                break;

                            case PSHostGrpcFrameKind.ClosePacket:
                                {
                                    Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                    string psGuidString = psGuid.ToString();
                                    if (psGuid == Guid.Empty)
                                    {
                                        s_closeTransportManagerMethod?.Invoke(sessionTM, new object?[] { null });
                                    }
                                    else
                                    {
                                        object? cmdTM = s_getCommandTransportManagerMethod?.Invoke(sessionTM, new object[] { psGuid });
                                        if (cmdTM != null)
                                        {
                                            s_closeTransportManagerMethod?.Invoke(cmdTM, new object?[] { null });
                                        }

                                        s_removeCommandTransportManagerMethod?.Invoke(sessionTM, new object[] { psGuid });
                                    }

                                    await WriteFrameAsync(
                                        responseStream,
                                        writeLock,
                                        new PSHostGrpcFrame
                                        {
                                            Kind = PSHostGrpcFrameKind.CloseAck,
                                            PsGuid = psGuidString
                                        },
                                        linkedCts.Token).ConfigureAwait(false);
                                }
                                break;
                        }
                    }
                }
                catch (OperationCanceledException)
                {
                }
                catch (IOException)
                {
                }
                catch (ObjectDisposedException)
                {
                }
                catch (RpcException ex) when (ex.StatusCode == StatusCode.Cancelled)
                {
                }
                finally
                {
                    _owner.RemoveConnection(connectionId);

                    if (sessionTM != null)
                    {
                        try
                        {
                            s_closeTransportManagerMethod?.Invoke(sessionTM, new object?[] { null });
                        }
                        catch { }
                    }
                }
            }

            private static object CreateOutOfProcessTextWriter(TextWriter innerWriter)
            {
                var ctor = s_outOfProcessTextWriterType?.GetConstructor(
                    s_anyFlags,
                    binder: null,
                    types: new[] { typeof(TextWriter) },
                    modifiers: null)
                    ?? throw new InvalidOperationException("Could not locate OutOfProcessTextWriter(TextWriter) constructor.");

                return ctor.Invoke(new object[] { innerWriter });
            }

            private static object CreateLocalSenderInfo()
            {
                object identity = s_psIdentityCtor!.Invoke(new object?[] { string.Empty, true, string.Empty, null });
                object principal = s_psPrincipalCtor!.Invoke(new object?[] { identity, null });
                return s_psSenderInfoCtor!.Invoke(new object?[] { principal, "http://localhost" });
            }

            private static bool TryHookRunspacePoolDriver(
                object serverSession,
                object sessionTM,
                IServerStreamWriter<PSHostGrpcFrame> responseStream,
                SemaphoreSlim writeLock,
                PSHostGrpcPsrpCodec outboundCodec,
                CancellationToken cancellationToken)
            {
                object? rpDriver = s_runspacePoolDriverField?.GetValue(serverSession);
                if (rpDriver == null)
                {
                    return false;
                }

                // Clear _rsToUseForSteppablePipeline so ServerPowerShellDriver uses RunspacePool
                // and RunspacePool.Close() does not double-close the runspace.
                s_rsToUseForSteppablePipelineField?.SetValue(rpDriver, null);

                object? dsHandler = s_rpDriverDsHandlerProp?.GetValue(rpDriver);
                if (dsHandler == null || s_createAndInvokePowerShellField == null || s_remoteDataEventArgsRdoPsObjectType == null)
                {
                    return true;
                }

                if (s_createAndInvokePowerShellField.GetValue(dsHandler) is not Delegate originalHandler)
                {
                    return true;
                }

                var hook = new DirectPowerShellDriverHook(
                    rpDriver,
                    sessionTM,
                    originalHandler,
                    responseStream,
                    writeLock,
                    outboundCodec,
                    cancellationToken);

                Delegate replacement = hook.CreateEventHandlerDelegate(s_remoteDataEventArgsRdoPsObjectType);
                s_createAndInvokePowerShellField.SetValue(dsHandler, replacement);
                return true;
            }

            private static async Task WriteFrameAsync(
                IServerStreamWriter<PSHostGrpcFrame> responseStream,
                SemaphoreSlim writeLock,
                PSHostGrpcFrame frame,
                CancellationToken cancellationToken)
            {
                await writeLock.WaitAsync(cancellationToken).ConfigureAwait(false);
                try
                {
                    await responseStream.WriteAsync(frame).ConfigureAwait(false);
                }
                finally
                {
                    writeLock.Release();
                }
            }

            private static void WriteFramesSync(
                IServerStreamWriter<PSHostGrpcFrame> responseStream,
                SemaphoreSlim writeLock,
                IEnumerable<PSHostGrpcFrame> frames,
                CancellationToken cancellationToken)
            {
                if (cancellationToken.IsCancellationRequested)
                {
                    return;
                }

                writeLock.Wait(cancellationToken);
                try
                {
                    foreach (var frame in frames)
                    {
                        responseStream.WriteAsync(frame).GetAwaiter().GetResult();
                    }
                }
                catch when (cancellationToken.IsCancellationRequested)
                {
                }
                finally
                {
                    writeLock.Release();
                }
            }

            private static bool HasExplicitEncodingHeader(ServerCallContext context)
            {
                return context.RequestHeaders.Any(e => e.Key.Equals("x-psremoting-encoding", StringComparison.OrdinalIgnoreCase));
            }

            private static PSHostGrpcEncoding ResolveConnectionEncoding(ServerCallContext context, PSHostGrpcEncoding defaultEncoding)
            {
                var entry = context.RequestHeaders.FirstOrDefault(e => e.Key.Equals("x-psremoting-encoding", StringComparison.OrdinalIgnoreCase));
                if (entry != null && Enum.TryParse<PSHostGrpcEncoding>(entry.Value, ignoreCase: true, out var parsedEncoding))
                {
                    return parsedEncoding;
                }

                return defaultEncoding;
            }

            private sealed class DirectPowerShellDriverHook
            {
                private readonly object _rpDriver;
                private readonly object _sessionTM;
                private readonly Delegate _originalHandler;
                private readonly IServerStreamWriter<PSHostGrpcFrame> _responseStream;
                private readonly SemaphoreSlim _writeLock;
                private readonly PSHostGrpcPsrpCodec _outboundCodec;
                private readonly CancellationToken _cancellationToken;

                public DirectPowerShellDriverHook(
                    object rpDriver,
                    object sessionTM,
                    Delegate originalHandler,
                    IServerStreamWriter<PSHostGrpcFrame> responseStream,
                    SemaphoreSlim writeLock,
                    PSHostGrpcPsrpCodec outboundCodec,
                    CancellationToken cancellationToken)
                {
                    _rpDriver = rpDriver;
                    _sessionTM = sessionTM;
                    _originalHandler = originalHandler;
                    _responseStream = responseStream;
                    _writeLock = writeLock;
                    _outboundCodec = outboundCodec;
                    _cancellationToken = cancellationToken;
                }

                public Delegate CreateEventHandlerDelegate(Type eventArgsType)
                {
                    var method = typeof(DirectPowerShellDriverHook).GetMethod(
                        nameof(HandleCreateAndInvokePowerShell),
                        BindingFlags.Instance | BindingFlags.NonPublic)!
                        .MakeGenericMethod(eventArgsType);

                    return Delegate.CreateDelegate(typeof(EventHandler<>).MakeGenericType(eventArgsType), this, method);
                }

                private void HandleCreateAndInvokePowerShell<TEventArgs>(object? sender, TEventArgs eventArgs)
                {
                    object? rdo = s_rdoEventArgsDataProp?.GetValue(eventArgs);
                    PSObject? data = rdo != null ? s_rdoPsObjectDataProp?.GetValue(rdo) as PSObject : null;
                    if (rdo == null || data == null || s_serverPowerShellDriverCtor == null)
                    {
                        ThreadPool.QueueUserWorkItem(_ =>
                        {
                            try
                            {
                                _originalHandler.DynamicInvoke(sender, eventArgs);
                            }
                            catch
                            {
                            }
                        });
                        return;
                    }

                    object? remoteHost = s_rpDriverServerRemoteHostProp?.GetValue(_rpDriver);
                    object? debugger = s_rpDriverServerRemoteDebuggerField?.GetValue(_rpDriver);
                    bool isRunspacePushed = remoteHost != null && s_remoteHostIsRunspacePushedProp?.GetValue(remoteHost) is bool pushed && pushed;
                    bool inBreakpoint = debugger != null && s_debuggerInBreakpointProp?.GetValue(debugger) is bool bp && bp;
                    bool isNested = s_getIsNestedMethod?.Invoke(null, new object[] { data }) is bool nested && nested;

                    if (isRunspacePushed || inBreakpoint || isNested)
                    {
                        ThreadPool.QueueUserWorkItem(_ =>
                        {
                            try
                            {
                                _originalHandler.DynamicInvoke(sender, eventArgs);
                            }
                            catch
                            {
                            }
                        });
                        return;
                    }

                    var powerShell = (System.Management.Automation.PowerShell)s_getPowerShellMethod!.Invoke(null, new object[] { data })!;
                    bool noInput = (bool)s_getNoInputMethod!.Invoke(null, new object[] { data })!;
                    ApartmentState apartmentState = (ApartmentState)s_getApartmentStateMethod!.Invoke(null, new object[] { data })!;
                    object remoteStreamOptions = s_getRemoteStreamOptionsMethod!.Invoke(null, new object[] { data })!;
                    bool addToHistory = (bool)s_getAddToHistoryMethod!.Invoke(null, new object[] { data })!;
                    object? hostInfo = s_getHostInfoMethod?.Invoke(null, new object[] { data });

                    Guid runspacePoolId = (Guid)s_rpDriverInstanceIdProp!.GetValue(_rpDriver)!;
                    Guid powerShellId = (Guid)s_rdoPsObjectPowerShellIdProp!.GetValue(rdo)!;

                    object driver = s_serverPowerShellDriverCtor.Invoke(new object?[]
                    {
                        powerShell,
                        null,
                        noInput,
                        powerShellId,
                        runspacePoolId,
                        _rpDriver,
                        apartmentState,
                        hostInfo,
                        remoteStreamOptions,
                        addToHistory,
                        null
                    });

                    object? inputCollection = s_psDriverInputCollectionProp?.GetValue(driver);
                    s_rpDriverInputCollectionField?.SetValue(_rpDriver, inputCollection);

                    AttachDirectDriverHandlers(driver, powerShellId, runspacePoolId);
                    ThreadPool.QueueUserWorkItem(_ =>
                    {
                        try
                        {
                            s_psDriverStartMethod?.Invoke(driver, Array.Empty<object>());
                        }
                        catch
                        {
                        }
                    });
                }

                private void AttachDirectDriverHandlers(object driver, Guid powerShellId, Guid runspacePoolId)
                {
                    var localOutput = s_psDriverLocalOutputField?.GetValue(driver) as PSDataCollection<PSObject>;
                    var localPowerShell = s_psDriverLocalPowerShellProp?.GetValue(driver) as System.Management.Automation.PowerShell;
                    object? syncObject = s_psDriverSyncObjectField?.GetValue(driver) ?? new object();
                    int streamOptionsValue = s_psDriverRemoteStreamOptionsProp?.GetValue(driver) is object rso
                        ? Convert.ToInt32(rso, CultureInfo.InvariantCulture)
                        : 0;

                    if (localOutput == null || localPowerShell == null)
                    {
                        return;
                    }

                    // Unregister default CLIXML output/stream/state handlers installed by ServerPowerShellDriver..ctor.
                    s_psDriverUnregisterOutputHandlersMethod?.Invoke(driver, new object[] { localOutput });
                    s_psDriverUnregisterPsHandlersMethod?.Invoke(driver, new object[] { localPowerShell });

                    string psGuidString = powerShellId.ToString();

                    void DrainOutput()
                    {
                        lock (syncObject)
                        {
                            var items = localOutput.ReadAll();
                            for (int i = 0; i < items.Count; i++)
                            {
                                PSObject? item = items[i];
                                if (item != null)
                                {
                                    WriteFramesSync(
                                        _responseStream,
                                        _writeLock,
                                        _outboundCodec.EncodeDirectPayloadToFrames(
                                            destination: 1u, // RemotingDestination.Client
                                            dataType: PowerShellOutputDataType,
                                            runspacePoolId: runspacePoolId,
                                            powerShellId: powerShellId,
                                            payloadObject: item,
                                            psGuid: psGuidString),
                                        _cancellationToken);
                                }
                            }
                        }
                    }

                    void SendErrorRecord(ErrorRecord errorRecord)
                    {
                        if ((streamOptionsValue & 1) != 0) // RemoteStreamOptions.AddInvocationInfoToErrorRecord == 1
                        {
                            s_preserveInvocationInfoOnceProp?.SetValue(errorRecord, true);
                        }

                        object? rdo = s_genErrorRecordMethod?.Invoke(null, new object[] { errorRecord, runspacePoolId, powerShellId });
                        if (rdo != null)
                        {
                            WriteFramesSync(
                                _responseStream,
                                _writeLock,
                                _outboundCodec.EncodeRemoteDataObjectToFrames(rdo, psGuidString),
                                _cancellationToken);
                        }
                    }

                    void DrainErrors()
                    {
                        var errors = localPowerShell.Streams.Error.ReadAll();
                        for (int i = 0; i < errors.Count; i++)
                        {
                            if (errors[i] != null)
                            {
                                SendErrorRecord(errors[i]);
                            }
                        }
                    }

                    EventHandler<DataAddedEventArgs>? outputHandler = null;
                    EventHandler<DataAddedEventArgs>? errorHandler = null;
                    EventHandler<DataAddedEventArgs>? warningHandler = null;
                    EventHandler<DataAddedEventArgs>? verboseHandler = null;
                    EventHandler<DataAddedEventArgs>? debugHandler = null;
                    EventHandler<DataAddedEventArgs>? progressHandler = null;
                    EventHandler<DataAddedEventArgs>? infoHandler = null;
                    EventHandler<PSInvocationStateChangedEventArgs>? stateHandler = null;

                    outputHandler = (_, _) => DrainOutput();
                    errorHandler = (_, _) => DrainErrors();

                    warningHandler = (_, _) =>
                    {
                        var records = localPowerShell.Streams.Warning.ReadAll();
                        for (int i = 0; i < records.Count; i++)
                        {
                            if ((streamOptionsValue & 2) != 0)
                            {
                                s_serializeExtendedInfoProp?.SetValue(records[i], true);
                            }

                            object? rdo = s_genInformationalMethod?.Invoke(null, new object?[] { records[i], runspacePoolId, powerShellId, s_warningDataTypeEnum });
                            if (rdo != null)
                            {
                                WriteFramesSync(_responseStream, _writeLock, _outboundCodec.EncodeRemoteDataObjectToFrames(rdo, psGuidString), _cancellationToken);
                            }
                        }
                    };

                    verboseHandler = (_, _) =>
                    {
                        var records = localPowerShell.Streams.Verbose.ReadAll();
                        for (int i = 0; i < records.Count; i++)
                        {
                            if ((streamOptionsValue & 4) != 0)
                            {
                                s_serializeExtendedInfoProp?.SetValue(records[i], true);
                            }

                            object? rdo = s_genInformationalMethod?.Invoke(null, new object?[] { records[i], runspacePoolId, powerShellId, s_verboseDataTypeEnum });
                            if (rdo != null)
                            {
                                WriteFramesSync(_responseStream, _writeLock, _outboundCodec.EncodeRemoteDataObjectToFrames(rdo, psGuidString), _cancellationToken);
                            }
                        }
                    };

                    debugHandler = (_, _) =>
                    {
                        var records = localPowerShell.Streams.Debug.ReadAll();
                        for (int i = 0; i < records.Count; i++)
                        {
                            if ((streamOptionsValue & 8) != 0)
                            {
                                s_serializeExtendedInfoProp?.SetValue(records[i], true);
                            }

                            object? rdo = s_genInformationalMethod?.Invoke(null, new object?[] { records[i], runspacePoolId, powerShellId, s_debugDataTypeEnum });
                            if (rdo != null)
                            {
                                WriteFramesSync(_responseStream, _writeLock, _outboundCodec.EncodeRemoteDataObjectToFrames(rdo, psGuidString), _cancellationToken);
                            }
                        }
                    };

                    progressHandler = (_, _) =>
                    {
                        var records = localPowerShell.Streams.Progress.ReadAll();
                        for (int i = 0; i < records.Count; i++)
                        {
                            object? rdo = s_genProgressMethod?.Invoke(null, new object[] { records[i], runspacePoolId, powerShellId });
                            if (rdo != null)
                            {
                                WriteFramesSync(_responseStream, _writeLock, _outboundCodec.EncodeRemoteDataObjectToFrames(rdo, psGuidString), _cancellationToken);
                            }
                        }
                    };

                    infoHandler = (_, _) =>
                    {
                        var records = localPowerShell.Streams.Information.ReadAll();
                        for (int i = 0; i < records.Count; i++)
                        {
                            object? rdo = s_genInformationMethod?.Invoke(null, new object[] { records[i], runspacePoolId, powerShellId });
                            if (rdo != null)
                            {
                                WriteFramesSync(_responseStream, _writeLock, _outboundCodec.EncodeRemoteDataObjectToFrames(rdo, psGuidString), _cancellationToken);
                            }
                        }
                    };

                    stateHandler = (_, e) =>
                    {
                        PSInvocationState state = e.InvocationStateInfo.State;
                        if (state == PSInvocationState.Completed ||
                            state == PSInvocationState.Failed ||
                            state == PSInvocationState.Stopped)
                        {
                            bool runningExtra = s_psRunningExtraCommandsProp?.GetValue(localPowerShell) is bool extra && extra;
                            if (runningExtra)
                            {
                                return;
                            }

                            localOutput.DataAdded -= outputHandler;
                            localPowerShell.Streams.Error.DataAdded -= errorHandler;
                            localPowerShell.Streams.Warning.DataAdded -= warningHandler;
                            localPowerShell.Streams.Verbose.DataAdded -= verboseHandler;
                            localPowerShell.Streams.Debug.DataAdded -= debugHandler;
                            localPowerShell.Streams.Progress.DataAdded -= progressHandler;
                            localPowerShell.Streams.Information.DataAdded -= infoHandler;
                            localPowerShell.InvocationStateChanged -= stateHandler;

                            DrainOutput();
                            DrainErrors();

                            object? dsHandler = s_psDriverDsHandlerProp?.GetValue(driver);
                            if (dsHandler != null)
                            {
                                s_psDsHandlerRaiseRemoveAssocMethod?.Invoke(dsHandler, Array.Empty<object>());
                                s_psDriverUnregisterDsHandlersMethod?.Invoke(driver, new object[] { dsHandler });
                            }

                            object? stateRdo = s_genStateInfoMethod?.Invoke(
                                null,
                                new object[] { e.InvocationStateInfo, powerShellId, runspacePoolId });
                            if (stateRdo != null)
                            {
                                WriteFramesSync(
                                    _responseStream,
                                    _writeLock,
                                    _outboundCodec.EncodeRemoteDataObjectToFrames(stateRdo, psGuidString),
                                    _cancellationToken);
                            }

                            object? cmdTM = s_getCommandTransportManagerMethod?.Invoke(_sessionTM, new object[] { powerShellId });
                            if (cmdTM != null)
                            {
                                s_closeTransportManagerMethod?.Invoke(cmdTM, new object?[] { null });
                            }

                            object? remoteHost = s_psDriverRemoteHostField?.GetValue(driver);
                            object? methodExecutor = remoteHost != null ? s_remoteHostServerMethodExecutorProp?.GetValue(remoteHost) : null;
                            if (methodExecutor != null)
                            {
                                s_abortAllCallsMethod?.Invoke(methodExecutor, Array.Empty<object>());
                            }

                            localPowerShell.Dispose();
                        }
                    };

                    localOutput.DataAdded += outputHandler;
                    localPowerShell.Streams.Error.DataAdded += errorHandler;
                    localPowerShell.Streams.Warning.DataAdded += warningHandler;
                    localPowerShell.Streams.Verbose.DataAdded += verboseHandler;
                    localPowerShell.Streams.Debug.DataAdded += debugHandler;
                    localPowerShell.Streams.Progress.DataAdded += progressHandler;
                    localPowerShell.Streams.Information.DataAdded += infoHandler;
                    localPowerShell.InvocationStateChanged += stateHandler;
                }
            }

            private sealed class GrpcServerTextWriter : TextWriter
            {
                private readonly IServerStreamWriter<PSHostGrpcFrame> _responseStream;
                private readonly SemaphoreSlim _writeLock;
                private readonly PSHostGrpcPsrpCodec _outboundCodec;
                private readonly bool _isError;
                private readonly CancellationToken _cancellationToken;

                public override Encoding Encoding => Encoding.UTF8;

                public GrpcServerTextWriter(
                    IServerStreamWriter<PSHostGrpcFrame> responseStream,
                    SemaphoreSlim writeLock,
                    PSHostGrpcPsrpCodec outboundCodec,
                    bool isError,
                    CancellationToken cancellationToken)
                {
                    _responseStream = responseStream;
                    _writeLock = writeLock;
                    _outboundCodec = outboundCodec;
                    _isError = isError;
                    _cancellationToken = cancellationToken;
                }

                public override void WriteLine(string? value)
                {
                    if (string.IsNullOrWhiteSpace(value) || _cancellationToken.IsCancellationRequested)
                    {
                        return;
                    }

                    if (_isError)
                    {
                        var frame = new PSHostGrpcFrame
                        {
                            Kind = PSHostGrpcFrameKind.Error,
                            Payload = ByteString.CopyFromUtf8(value + "\n")
                        };
                        WriteFramesSync(_responseStream, _writeLock, new[] { frame }, _cancellationToken);
                        return;
                    }

                    WriteFramesSync(
                        _responseStream,
                        _writeLock,
                        _outboundCodec.EncodePacketLineToFrames(value),
                        _cancellationToken);
                }
            }
        }
    }
}
