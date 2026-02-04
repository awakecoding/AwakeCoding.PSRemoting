using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Net.Sockets;
using System.Threading;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Connection info for TCP-based PowerShell remoting client connections
    /// </summary>
    internal sealed class PSHostTcpClientInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        public string HostName { get; set; }

        public int Port { get; set; }

        public new int OpenTimeout { get; set; } = 30000; // 30 second default timeout

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

        public PSHostTcpClientInfo(string hostName, int port)
        {
            HostName = hostName;
            ComputerName = hostName;
            Port = port;
        }

        public override RunspaceConnectionInfo Clone()
        {
            var connectionInfo = new PSHostTcpClientInfo(HostName, Port)
            {
                OpenTimeout = OpenTimeout
            };
            return connectionInfo;
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new PSHostTcpClientSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    /// <summary>
    /// Transport manager for TCP client connections to PSHostServer
    /// </summary>
    internal sealed class PSHostTcpClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly PSHostTcpClientInfo _connectionInfo;
        private TcpClient? _tcpClient = null;
        private NetworkStream? _networkStream = null;
        private StreamWriter? _streamWriter = null;
        private StreamReader? _streamReader = null;
        private CancellationTokenSource? _readerCts = null;
        private const string ThreadName = "PSHostTcpClient Reader Thread";

        internal PSHostTcpClientSessionTransportMgr(
            PSHostTcpClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
        }

        public override void CreateAsync()
        {
            _tcpClient = new TcpClient();

            // Connect with timeout
            using var cts = new CancellationTokenSource(_connectionInfo.OpenTimeout);
            try
            {
                var connectTask = _tcpClient.ConnectAsync(_connectionInfo.HostName, _connectionInfo.Port, cts.Token);
                connectTask.AsTask().Wait(cts.Token);
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException($"TCP connection to {_connectionInfo.HostName}:{_connectionInfo.Port} timed out after {_connectionInfo.OpenTimeout}ms");
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
            {
                throw new TimeoutException($"TCP connection to {_connectionInfo.HostName}:{_connectionInfo.Port} timed out after {_connectionInfo.OpenTimeout}ms");
            }

            _networkStream = _tcpClient.GetStream();

            // Create text reader/writer for line-based PSRP protocol
            _streamReader = new StreamReader(_networkStream, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
            _streamWriter = new StreamWriter(_networkStream, System.Text.Encoding.UTF8, 4096, leaveOpen: true)
            {
                AutoFlush = true
            };

            // Set up the message writer for outgoing data
            SetMessageWriter(_streamWriter);

            // Create cancellation token for reader thread
            _readerCts = new CancellationTokenSource();

            // Start the reader thread
            StartReaderThread();

            // Send initial protocol fragment (must be after SetMessageWriter)
            SendOneItem();
        }

        public override void CloseAsync()
        {
            // Cancel the reader thread first
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Close the network stream to unblock any pending reads
            try
            {
                _networkStream?.Close();
            }
            catch { }

            // Call base implementation
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
            try
            {
                // Start reader loop with cancellation support
                while (_readerCts != null && !_readerCts.IsCancellationRequested)
                {
                    // Use async read with cancellation token for proper timeout handling
                    var readTask = _streamReader?.ReadLineAsync(_readerCts.Token);
                    if (readTask == null)
                    {
                        break;
                    }

                    string? data;
                    try
                    {
                        data = readTask.Value.AsTask().GetAwaiter().GetResult();
                    }
                    catch (OperationCanceledException)
                    {
                        // Cancellation requested - exit gracefully
                        break;
                    }

                    if (data == null)
                    {
                        // End of stream - connection closed
                        break;
                    }

                    // Process the received data (line-based PSRP output)
                    HandleOutputDataReceived(data);
                }
            }
            catch (ObjectDisposedException)
            {
                // Normal reader thread end
            }
            catch (IOException)
            {
                // Connection closed or disconnected
            }
            catch (OperationCanceledException)
            {
                // Cancellation requested
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
            // Cancel the reader thread first
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Close network stream to unblock any pending reads
            try
            {
                _networkStream?.Close();
                _networkStream?.Dispose();
            }
            catch { }

            try
            {
                _streamReader?.Dispose();
            }
            catch { }

            try
            {
                _streamWriter?.Dispose();
            }
            catch { }

            try
            {
                _tcpClient?.Close();
                _tcpClient?.Dispose();
            }
            catch { }

            _readerCts?.Dispose();
            _readerCts = null;
            _streamReader = null;
            _streamWriter = null;
            _networkStream = null;
            _tcpClient = null;
        }
    }
}
