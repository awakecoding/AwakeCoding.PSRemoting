using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
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
                OpenTimeout = OpenTimeout
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
    /// </summary>
    internal sealed class PSHostGrpcClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private const string ThreadName = "PSHostGrpcClient Reader Thread";

        private readonly PSHostGrpcClientInfo _connectionInfo;
        private readonly object _sendLock = new object();

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
        }

        public override void CreateAsync()
        {
            PSHostGrpcPlatform.EnsureSupported();

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
            _call = client.Connect();

            SetMessageWriter(new GrpcTextWriter(this));

            _readerCts = new CancellationTokenSource();
            StartReaderThread();

            SendOneItem();
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
                    call.RequestStream.WriteAsync(new PSHostGrpcFrame
                    {
                        Kind = PSHostGrpcFrameKind.Input,
                        Payload = ByteString.CopyFrom(data, Encoding.UTF8)
                    }).GetAwaiter().GetResult();
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
                        case PSHostGrpcFrameKind.Output:
                            ProcessPayload(frame.Payload, outputBuffer, HandleOutputDataReceived);
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
