using System;
using System.IO;
using System.IO.Pipes;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Threading;

namespace AwakeCoding.PSRemoting.PowerShell
{
    internal sealed class PSHostNamedPipeInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        public string PipeName { get; set; }

        public string AppDomainName { get; set; }

        public new int OpenTimeout { get; set; } = 5000; // 5 second default timeout

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

        public PSHostNamedPipeInfo(string computerName, string pipeName, string appDomainName = "DefaultAppDomain")
        {
            ComputerName = computerName;
            PipeName = pipeName;
            AppDomainName = appDomainName;
        }

        public override RunspaceConnectionInfo Clone()
        {
            var connectionInfo = new PSHostNamedPipeInfo(ComputerName, PipeName, AppDomainName);
            return connectionInfo;
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new PSHostNamedPipeSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    internal sealed class PSHostNamedPipeSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly PSHostNamedPipeInfo _connectionInfo;
        private NamedPipeClientStream? _pipeStream = null;
        private StreamWriter? _streamWriter = null;
        private StreamReader? _streamReader = null;
        private CancellationTokenSource? _readerCts = null;
        private const string _threadName = "PSHostNamedPipe Reader Thread";

        internal PSHostNamedPipeSessionTransportMgr(
            PSHostNamedPipeInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
        }

        public override void CreateAsync()
        {
            // Create a client stream to the local server using the pipe name without prefix
            // Using duplex direction for bidirectional communication
            _pipeStream = new NamedPipeClientStream(
                ".", // local machine
                _connectionInfo.PipeName,
                PipeDirection.InOut,
                PipeOptions.Asynchronous);

            // Use ConnectAsync with CancellationToken for reliable timeout on all platforms
            // On Linux, the synchronous Connect() timeout parameter doesn't work reliably
            using var cts = new CancellationTokenSource(_connectionInfo.OpenTimeout);
            try
            {
                // ConnectAsync respects the CancellationToken properly
                var connectTask = _pipeStream.ConnectAsync(cts.Token);
                if (!connectTask.Wait(_connectionInfo.OpenTimeout))
                {
                    cts.Cancel();
                    throw new TimeoutException($"Named pipe connection timed out after {_connectionInfo.OpenTimeout}ms");
                }
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException($"Named pipe connection timed out after {_connectionInfo.OpenTimeout}ms");
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
            {
                throw new TimeoutException($"Named pipe connection timed out after {_connectionInfo.OpenTimeout}ms");
            }

            // Create text reader/writer for line-based PSRP protocol
            _streamReader = new StreamReader(_pipeStream, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
            _streamWriter = new StreamWriter(_pipeStream, System.Text.Encoding.UTF8, 4096, leaveOpen: true)
            {
                AutoFlush = true
            };

            // Set up the message writer for outgoing data
            SetMessageWriter(_streamWriter);

            // Create cancellation token for reader thread
            _readerCts = new CancellationTokenSource();

            // Start the reader thread (pattern matches NamedPipeClientSessionTransportManagerBase)
            StartReaderThread();
        }

        public override void CloseAsync()
        {
            // Cancel the reader thread first before attempting to send close packet
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Dispose the pipe stream to unblock any pending reads
            try
            {
                _pipeStream?.Dispose();
            }
            catch { }

            // Call base implementation
            base.CloseAsync();
        }

        private void StartReaderThread()
        {
            Thread readerThread = new Thread(ProcessReaderThread);
            readerThread.Name = _threadName;
            readerThread.IsBackground = true;
            readerThread.Start();
        }

        private void ProcessReaderThread()
        {
            try
            {
                // Send one fragment - must be done inside reader thread
                SendOneItem();

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
                        // End of stream - pipe closed or error
                        break;
                    }

                    // Process the received data (handles both normal and error messages)
                    HandleDataReceived(data);
                }
            }
            catch (ObjectDisposedException)
            {
                // Normal reader thread end
            }
            catch (IOException)
            {
                // Pipe closed or disconnected
            }
            catch (OperationCanceledException)
            {
                // Cancellation requested
            }
        }

        protected override void Dispose(bool isDisposing)
        {
            // Cleanup first to cancel pending reads before base class tries to wait
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

            // Dispose the pipe stream to unblock any pending reads
            // This must be done before disposing readers/writers
            try
            {
                _pipeStream?.Dispose();
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

            _readerCts = null;
            _streamReader = null;
            _streamWriter = null;
            _pipeStream = null;
        }
    }
}
