using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq.Expressions;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Text;
using System.Threading;
using Google.Protobuf;
using Grpc.Core;
using AwakeCoding.PSRemoting.PowerShell.Grpc;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Connection info for gRPC-based PowerShell remoting client connections.
    /// </summary>
    internal sealed class PSHostGrpcClientInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        public Uri GrpcUri { get; set; }

        public PSHostGrpcEncoding Encoding { get; set; } = PSHostGrpcEncoding.Protobuf;

        internal bool HasExplicitEncoding { get; set; }

        public new int OpenTimeout { get; set; } = 30000;

        public override PSCredential? Credential
        {
            get { return null; }
            set { throw new NotImplementedException(); }
        }

        public override AuthenticationMechanism AuthenticationMechanism
        {
            get { return AuthenticationMechanism.Default; }
            set { throw new NotImplementedException(); }
        }

        public override string CertificateThumbprint
        {
            get { return string.Empty; }
            set { throw new NotImplementedException(); }
        }

        public PSHostGrpcClientInfo(Uri grpcUri)
        {
            GrpcUri = grpcUri;
            ComputerName = grpcUri.Host;
        }

        public override RunspaceConnectionInfo Clone()
        {
            return new PSHostGrpcClientInfo(GrpcUri)
            {
                OpenTimeout = OpenTimeout,
                Encoding = Encoding,
                HasExplicitEncoding = HasExplicitEncoding
            };
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new PSHostGrpcClientSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    /// <summary>
    /// Transport manager for gRPC client connections to PSHostServer.
    /// Directly injects decoded <c>RemoteDataObject&lt;PSObject&gt;</c> instances on the receive path
    /// and reads raw PSRP fragments directly from <c>PrioritySendDataCollection</c> / <c>SerializedDataStream</c>
    /// on the send path, bypassing <c>OutOfProcessUtils</c> XML/Base64 packet encoding and
    /// <c>Fragmentor.DeserializeToPSObject</c> CLIXML deserialization.
    /// </summary>
    internal sealed class PSHostGrpcClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private const string ThreadName = "PSHostGrpcClient Reader Thread";
        private static readonly string s_emptyGuidString = Guid.Empty.ToString();

        private static readonly BindingFlags s_anyFlags = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;

        private static readonly Type? s_baseClientTransportMgrType = typeof(PSObject).Assembly.GetType(
            "System.Management.Automation.Remoting.Client.BaseClientTransportManager");
        private static readonly Type? s_clientSessionTransportMgrBaseType = typeof(PSObject).Assembly.GetType(
            "System.Management.Automation.Remoting.Client.ClientSessionTransportManagerBase");
        private static readonly Type? s_baseClientCmdTransportMgrType = typeof(PSObject).Assembly.GetType(
            "System.Management.Automation.Remoting.Client.BaseClientCommandTransportManager");
        private static readonly Type? s_prioritySendDataCollectionType = typeof(PSObject).Assembly.GetType(
            "System.Management.Automation.Remoting.PrioritySendDataCollection");
        private static readonly Type? s_onDataAvailableToSendCallbackType = s_prioritySendDataCollectionType?.GetNestedType(
            "OnDataAvailableCallback", s_anyFlags);
        private static readonly Type? s_dataPriorityType = typeof(PSObject).Assembly.GetType(
            "System.Management.Automation.Remoting.DataPriorityType");
        private static readonly object? s_defaultDataPriorityValue = s_dataPriorityType != null
            ? Enum.ToObject(s_dataPriorityType, 0)
            : null;
        private static readonly Type? s_serializedDataStreamType = typeof(PSObject).Assembly.GetType(
            "System.Management.Automation.Remoting.SerializedDataStream");

        private static readonly MethodInfo? s_onDataAvailableHandlerMethod = s_baseClientTransportMgrType?.GetMethod(
            "OnDataAvailableHandler", s_anyFlags);
        private static readonly FieldInfo? s_dataToBeSentField = s_baseClientTransportMgrType?.GetField(
            "dataToBeSent", s_anyFlags);
        private static readonly FieldInfo? s_syncObjectField = s_baseClientTransportMgrType?.GetField(
            "syncObject", s_anyFlags);
        private static readonly FieldInfo? s_isClosedField = s_baseClientTransportMgrType?.GetField(
            "isClosed", s_anyFlags);

        private static readonly MethodInfo? s_disposeMessageQueueMethod = s_clientSessionTransportMgrBaseType?.GetMethod(
            "DisposeMessageQueue", s_anyFlags);
        private static readonly MethodInfo? s_getCommandTransportManagerMethod = s_clientSessionTransportMgrBaseType?.GetMethod(
            "GetCommandTransportManager", s_anyFlags);
        private static readonly MethodInfo? s_onCloseAckReceivedMethod = s_clientSessionTransportMgrBaseType?.GetMethod(
            "OnCloseAckReceived", s_anyFlags);
        private static readonly MethodInfo? s_onSignalAckPacketReceivedMethod = s_clientSessionTransportMgrBaseType?.GetMethod(
            "OnSignalAckPacketReceived", s_anyFlags);
        private static readonly MethodInfo? s_onCommandCreationPacketReceivedMethod = s_clientSessionTransportMgrBaseType?.GetMethod(
            "OnCommandCreationPacketReceived", s_anyFlags);
        private static readonly MethodInfo? s_onClosePacketReceivedMethod = s_clientSessionTransportMgrBaseType?.GetMethod(
            "OnClosePacketReceived", s_anyFlags);
        private static readonly MethodInfo? s_onSignalPacketReceivedMethod = s_clientSessionTransportMgrBaseType?.GetMethod(
            "OnSignalPacketReceived", s_anyFlags);

        private static readonly FieldInfo? s_serializedPipelineField = s_baseClientCmdTransportMgrType?.GetField(
            "serializedPipeline", s_anyFlags);
        private static readonly MethodInfo? s_sdsReadOrRegisterCallbackMethod = s_serializedDataStreamType?.GetMethod(
            "ReadOrRegisterCallback", s_anyFlags);
        private static readonly MethodInfo? s_psdcReadOrRegisterCallbackMethod = s_prioritySendDataCollectionType?.GetMethod(
            "ReadOrRegisterCallback", s_anyFlags);

        private readonly PSHostGrpcClientInfo _connectionInfo;
        private readonly object _sendLock = new object();
        private readonly PSHostGrpcPsrpCodec _outboundCodec = new PSHostGrpcPsrpCodec();
        private readonly PSHostGrpcPsrpCodec _inboundCodec = new PSHostGrpcPsrpCodec();
        private readonly Delegate? _sessionSendCallback;
        private readonly ConcurrentDictionary<Guid, Delegate> _commandSendCallbacks = new ConcurrentDictionary<Guid, Delegate>();

        private Channel? _channel;
        private AsyncDuplexStreamingCall<PSHostGrpcFrame, PSHostGrpcFrame>? _call;
        private CancellationTokenSource? _readerCts;
        private volatile bool _isClosed;

        internal PSHostGrpcClientSessionTransportMgr(
            PSHostGrpcClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
            _outboundCodec.Encoding = connectionInfo.Encoding;
            _sessionSendCallback = CreateSendCallbackDelegate(this, Guid.Empty);
        }

        public override void CreateAsync()
        {
            PSHostGrpcPlatform.EnsureSupported();

            // Stop the unused XML packet queue threads created by ClientSessionTransportManagerBase's constructor.
            s_disposeMessageQueueMethod?.Invoke(this, Array.Empty<object>());

            ChannelCredentials credentials = _connectionInfo.GrpcUri.Scheme.Equals("grpcs", StringComparison.OrdinalIgnoreCase)
                ? new SslCredentials()
                : ChannelCredentials.Insecure;

            _channel = new Channel(_connectionInfo.GrpcUri.Host, _connectionInfo.GrpcUri.Port, credentials);

            try
            {
                _channel.ConnectAsync(DateTime.UtcNow.AddMilliseconds(_connectionInfo.OpenTimeout)).GetAwaiter().GetResult();
            }
            catch (RpcException ex) when (ex.StatusCode == StatusCode.DeadlineExceeded)
            {
                throw new TimeoutException($"gRPC connection to {_connectionInfo.GrpcUri} timed out after {_connectionInfo.OpenTimeout}ms", ex);
            }
            catch (TimeoutException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new IOException($"Failed to connect to gRPC endpoint '{_connectionInfo.GrpcUri}'", ex);
            }

            var client = new PSHostGrpcTransport.PSHostGrpcTransportClient(_channel);
            Metadata? headers = null;
            if (_connectionInfo.HasExplicitEncoding)
            {
                headers = new Metadata
                {
                    { "x-psremoting-encoding", _connectionInfo.Encoding.ToString().ToLowerInvariant() }
                };
            }

            _call = client.Connect(headers);

            SetMessageWriter(new GrpcTextWriter(this));

            _readerCts = new CancellationTokenSource();
            StartReaderThread();

            SendOneSessionItem();
        }

        private void OnDataAvailableToSend(byte[]? data, int priorityType, Guid psGuid)
        {
            if (data != null)
            {
                SendRawFragment(data, priorityType, psGuid);
            }
        }

        private void SendOneSessionItem()
        {
            if (_isClosed || s_dataToBeSentField == null || s_psdcReadOrRegisterCallbackMethod == null || _sessionSendCallback == null)
            {
                return;
            }

            object? dataToBeSent = s_dataToBeSentField.GetValue(this);
            if (dataToBeSent == null)
            {
                return;
            }

            object?[] args = new object?[] { _sessionSendCallback, s_defaultDataPriorityValue };
            byte[]? data = s_psdcReadOrRegisterCallbackMethod.Invoke(dataToBeSent, args) as byte[];
            if (data != null)
            {
                int priorityType = args[1] != null ? Convert.ToInt32(args[1], CultureInfo.InvariantCulture) : 0;
                SendRawFragment(data, priorityType, Guid.Empty);
            }
        }

        private void SendOneCommandItem(Guid psGuid)
        {
            if (_isClosed || s_getCommandTransportManagerMethod == null)
            {
                return;
            }

            object? cmdTM = s_getCommandTransportManagerMethod.Invoke(this, new object[] { psGuid });
            if (cmdTM == null)
            {
                return;
            }

            object? syncObj = s_syncObjectField?.GetValue(cmdTM);
            if (syncObj == null)
            {
                return;
            }

            byte[]? data = null;
            int priorityType = 0;

            lock (syncObj)
            {
                if (IsTransportClosed(cmdTM))
                {
                    return;
                }

                if (s_serializedPipelineField?.GetValue(cmdTM) is Stream serializedPipeline &&
                    serializedPipeline.Length > 0 &&
                    s_sdsReadOrRegisterCallbackMethod != null)
                {
                    data = s_sdsReadOrRegisterCallbackMethod.Invoke(serializedPipeline, new object?[] { null }) as byte[];
                }

                if (data == null && s_dataToBeSentField?.GetValue(cmdTM) is object cmdDataToBeSent &&
                    s_psdcReadOrRegisterCallbackMethod != null)
                {
                    Delegate? cmdCallback = GetOrCreateCommandSendCallback(psGuid);
                    object?[] args = new object?[] { cmdCallback, s_defaultDataPriorityValue };
                    data = s_psdcReadOrRegisterCallbackMethod.Invoke(cmdDataToBeSent, args) as byte[];
                    if (args[1] != null)
                    {
                        priorityType = Convert.ToInt32(args[1], CultureInfo.InvariantCulture);
                    }
                }

                if (data != null)
                {
                    SendRawFragment(data, priorityType, psGuid);
                }
            }
        }

        private Delegate? GetOrCreateCommandSendCallback(Guid psGuid)
        {
            if (s_onDataAvailableToSendCallbackType == null)
            {
                return null;
            }

            return _commandSendCallbacks.GetOrAdd(psGuid, id => CreateSendCallbackDelegate(this, id)!);
        }

        private static Delegate? CreateSendCallbackDelegate(PSHostGrpcClientSessionTransportMgr owner, Guid psGuid)
        {
            if (s_onDataAvailableToSendCallbackType == null || s_dataPriorityType == null)
            {
                return null;
            }

            var handlerMethod = typeof(PSHostGrpcClientSessionTransportMgr).GetMethod(
                nameof(OnDataAvailableToSend),
                BindingFlags.Instance | BindingFlags.NonPublic);
            if (handlerMethod == null)
            {
                return null;
            }

            var dataParam = Expression.Parameter(typeof(byte[]), "data");
            var priorityParam = Expression.Parameter(s_dataPriorityType, "priorityType");
            var body = Expression.Call(
                Expression.Constant(owner),
                handlerMethod,
                dataParam,
                Expression.Convert(priorityParam, typeof(int)),
                Expression.Constant(psGuid));

            return Expression.Lambda(s_onDataAvailableToSendCallbackType, body, dataParam, priorityParam).Compile();
        }

        private void SendRawFragment(byte[] rawFragment, int priorityType, Guid psGuid)
        {
            if (_isClosed)
            {
                return;
            }

            string psGuidString = psGuid == Guid.Empty ? s_emptyGuidString : psGuid.ToString();
            // DataPriorityType.PromptResponse == 1, DataPriorityType.Default == 0
            string streamAttr = priorityType == 1 ? "PromptResponse" : "Default";
            WriteFrames(_outboundCodec.EncodeRawFragmentToFrames(rawFragment, psGuidString, streamAttr));
        }

        private void WriteFrame(PSHostGrpcFrame frame)
        {
            if (_isClosed)
            {
                return;
            }

            var call = _call ?? throw new InvalidOperationException("The gRPC stream has not been opened.");

            lock (_sendLock)
            {
                if (_isClosed)
                {
                    return;
                }

                try
                {
                    call.RequestStream.WriteAsync(frame).GetAwaiter().GetResult();
                }
                catch (Exception ex) when (!_isClosed)
                {
                    throw new IOException($"Failed to send frame to gRPC endpoint '{_connectionInfo.GrpcUri}'", ex);
                }
            }
        }

        private void WriteFrames(IEnumerable<PSHostGrpcFrame> frames)
        {
            if (_isClosed)
            {
                return;
            }

            var call = _call ?? throw new InvalidOperationException("The gRPC stream has not been opened.");

            lock (_sendLock)
            {
                if (_isClosed)
                {
                    return;
                }

                try
                {
                    foreach (var frame in frames)
                    {
                        call.RequestStream.WriteAsync(frame).GetAwaiter().GetResult();
                    }
                }
                catch (Exception ex) when (!_isClosed)
                {
                    throw new IOException($"Failed to send frames to gRPC endpoint '{_connectionInfo.GrpcUri}'", ex);
                }
            }
        }

        internal void SendData(string data)
        {
            if (_isClosed)
            {
                throw new IOException("The gRPC transport is closed.");
            }

            var call = _call ?? throw new InvalidOperationException("The gRPC stream has not been opened.");

            lock (_sendLock)
            {
                try
                {
                    using var reader = new StringReader(data);
                    string? line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                        {
                            continue;
                        }

                        foreach (var frame in _outboundCodec.EncodePacketLineToFrames(line))
                        {
                            call.RequestStream.WriteAsync(frame).GetAwaiter().GetResult();
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new IOException($"Failed to send data to gRPC endpoint '{_connectionInfo.GrpcUri}'", ex);
                }
            }
        }

        public override void CloseAsync()
        {
            base.CloseAsync();
        }

        private void StartReaderThread()
        {
            Thread readerThread = new Thread(ProcessReaderThread)
            {
                Name = ThreadName,
                IsBackground = true
            };
            readerThread.Start();
        }

        private void ProcessReaderThread()
        {
            var outputBuffer = new StringBuilder();
            var errorBuffer = new StringBuilder();

            try
            {
                while (_readerCts != null && !_readerCts.IsCancellationRequested && _call != null)
                {
                    bool hasFrame;
                    try
                    {
                        hasFrame = _call.ResponseStream.MoveNext(_readerCts.Token).GetAwaiter().GetResult();
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                    catch (RpcException ex) when (ex.StatusCode == StatusCode.Cancelled)
                    {
                        break;
                    }
                    catch (RpcException ex) when (ex.StatusCode == StatusCode.Unavailable && _isClosed)
                    {
                        break;
                    }

                    if (!hasFrame)
                    {
                        break;
                    }

                    var frame = _call.ResponseStream.Current;
                    switch (frame.Kind)
                    {
                        case PSHostGrpcFrameKind.Data:
                            HandleInboundDataFrame(frame);
                            break;

                        case PSHostGrpcFrameKind.DataAck:
                            HandleInboundDataAckFrame(frame);
                            break;

                        case PSHostGrpcFrameKind.CommandAck:
                            HandleInboundCommandAckFrame(frame);
                            break;

                        case PSHostGrpcFrameKind.CloseAck:
                            {
                                Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                if (psGuid != Guid.Empty)
                                {
                                    _commandSendCallbacks.TryRemove(psGuid, out _);
                                }
                                s_onCloseAckReceivedMethod?.Invoke(this, new object[] { psGuid });
                            }
                            break;

                        case PSHostGrpcFrameKind.SignalAck:
                            {
                                Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                s_onSignalAckPacketReceivedMethod?.Invoke(this, new object[] { psGuid });
                            }
                            break;

                        case PSHostGrpcFrameKind.Command:
                            {
                                Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                s_onCommandCreationPacketReceivedMethod?.Invoke(this, new object[] { psGuid });
                            }
                            break;

                        case PSHostGrpcFrameKind.ClosePacket:
                            {
                                Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                s_onClosePacketReceivedMethod?.Invoke(this, new object[] { psGuid });
                            }
                            break;

                        case PSHostGrpcFrameKind.Signal:
                            {
                                Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
                                s_onSignalPacketReceivedMethod?.Invoke(this, new object[] { psGuid });
                            }
                            break;

                        case PSHostGrpcFrameKind.Error:
                            ProcessPayload(frame.Payload, errorBuffer, HandleErrorDataReceived);
                            break;

                        case PSHostGrpcFrameKind.Close:
                            return;
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                // Normal reader thread end during cleanup.
            }
        }

        private void HandleInboundDataFrame(PSHostGrpcFrame frame)
        {
            if (!_connectionInfo.HasExplicitEncoding)
            {
                _outboundCodec.Encoding = PSHostGrpcPsrpCodec.FromWireEncoding(frame.PayloadEncoding);
            }

            bool hasObject = _inboundCodec.TryDecodeDataFrameToRemoteDataObject(
                frame,
                out Guid psGuid,
                out object? remoteDataObject);

            if (hasObject && remoteDataObject != null)
            {
                if (psGuid == Guid.Empty)
                {
                    if (!IsTransportClosed(this))
                    {
                        s_onDataAvailableHandlerMethod?.Invoke(this, new object[] { remoteDataObject });
                    }
                }
                else
                {
                    object? cmdTM = s_getCommandTransportManagerMethod?.Invoke(this, new object[] { psGuid });
                    if (cmdTM != null && !IsTransportClosed(cmdTM))
                    {
                        s_onDataAvailableHandlerMethod?.Invoke(cmdTM, new object[] { remoteDataObject });
                    }
                }
            }
        }

        private void HandleInboundDataAckFrame(PSHostGrpcFrame frame)
        {
            Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
            if (psGuid == Guid.Empty)
            {
                SendOneSessionItem();
            }
            else
            {
                SendOneCommandItem(psGuid);
            }
        }

        private void HandleInboundCommandAckFrame(PSHostGrpcFrame frame)
        {
            Guid psGuid = PSHostGrpcPsrpCodec.ParsePsGuid(frame.PsGuid);
            SendOneCommandItem(psGuid);
        }

        private static bool IsTransportClosed(object transportManager)
        {
            return s_isClosedField?.GetValue(transportManager) is bool closed && closed;
        }

        private static void ProcessPayload(ByteString payload, StringBuilder lineBuffer, Action<string?> handler)
        {
            if (payload.Length == 0)
            {
                return;
            }

            lineBuffer.Append(payload.ToStringUtf8());
            string bufferContent = lineBuffer.ToString();
            int lastNewlineIndex = bufferContent.LastIndexOf('\n');

            if (lastNewlineIndex < 0)
            {
                return;
            }

            string completeData = bufferContent.Substring(0, lastNewlineIndex + 1);
            lineBuffer.Clear();

            if (lastNewlineIndex < bufferContent.Length - 1)
            {
                lineBuffer.Append(bufferContent.Substring(lastNewlineIndex + 1));
            }

            using var reader = new StringReader(completeData);
            string? line;
            while ((line = reader.ReadLine()) != null)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    handler(line);
                }
            }
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                CleanupConnection();
            }

            base.Dispose(isDisposing);
        }

        protected override void CleanupConnection()
        {
            _isClosed = true;

            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            try
            {
                _call?.RequestStream.CompleteAsync().Wait(1000);
            }
            catch { }

            try
            {
                _call?.Dispose();
            }
            catch { }

            try
            {
                _channel?.ShutdownAsync().Wait(1000);
            }
            catch { }

            _readerCts?.Dispose();
            _readerCts = null;
            _call = null;
            _channel = null;
        }
    }

    /// <summary>
    /// Custom TextWriter that sends line-oriented remoting records over gRPC.
    /// </summary>
    internal sealed class GrpcTextWriter : TextWriter
    {
        private readonly PSHostGrpcClientSessionTransportMgr _transportMgr;
        private readonly StringBuilder _lineBuffer = new StringBuilder();

        public override Encoding Encoding => Encoding.UTF8;

        public GrpcTextWriter(PSHostGrpcClientSessionTransportMgr transportMgr)
        {
            _transportMgr = transportMgr;
        }

        public override void Write(char value)
        {
            _lineBuffer.Append(value);

            if (value == '\n')
            {
                Flush();
            }
        }

        public override void Write(string? value)
        {
            if (value == null)
            {
                return;
            }

            _lineBuffer.Append(value);

            if (value.EndsWith('\n'))
            {
                Flush();
            }
        }

        public override void WriteLine(string? value)
        {
            _lineBuffer.Append(value);
            _lineBuffer.Append('\n');
            Flush();
        }

        public override void Flush()
        {
            if (_lineBuffer.Length == 0)
            {
                return;
            }

            _transportMgr.SendData(_lineBuffer.ToString());
            _lineBuffer.Clear();
        }
    }
}
