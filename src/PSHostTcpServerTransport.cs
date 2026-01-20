using System;
using System.Diagnostics;
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
    /// Connection info for TCP-based PowerShell remoting connections
    /// Note: This is used for transient runspaces created by the server, NOT for PSSessions
    /// </summary>
    internal sealed class PSHostTcpConnectionInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        public TcpClient TcpClient { get; set; }

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

        public PSHostTcpConnectionInfo(string computerName, TcpClient tcpClient)
        {
            ComputerName = computerName;
            TcpClient = tcpClient;
        }

        public override RunspaceConnectionInfo Clone()
        {
            throw new NotImplementedException("PSHostTcpConnectionInfo does not support cloning");
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new PSHostTcpConnectionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    /// <summary>
    /// Transport manager for individual TCP connections to PowerShell subprocess
    /// Proxies NetworkStream ↔ subprocess stdin/stdout
    /// </summary>
    internal sealed class PSHostTcpConnectionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly PSHostTcpConnectionInfo _connectionInfo;
        private Process? _process = null;
        private NetworkStream? _networkStream = null;
        private CancellationTokenSource? _readerCts = null;
        private const string ThreadName = "PSHostTcpConnection Reader Thread";

        internal PSHostTcpConnectionTransportMgr(
            PSHostTcpConnectionInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
        }

        public override void CreateAsync()
        {
            try
            {
                // Get PowerShell executable
                var executable = PowerShellFinder.GetPowerShellPath();
                if (executable == null || !File.Exists(executable))
                {
                    throw new InvalidOperationException("PowerShell executable not found");
                }

                // Start PowerShell subprocess in server mode
                _process = new Process();
                _process.StartInfo.FileName = executable;
                _process.StartInfo.Arguments = "-NoLogo -NoProfile -s";
                _process.StartInfo.RedirectStandardInput = true;
                _process.StartInfo.RedirectStandardOutput = true;
                _process.StartInfo.RedirectStandardError = true;
                _process.StartInfo.UseShellExecute = false;
                _process.StartInfo.CreateNoWindow = true;

                _process.Start();

                // Get the network stream from TCP client
                _networkStream = _connectionInfo.TcpClient.GetStream();

                // Set up the message writer for outgoing data (network → subprocess)
                SetMessageWriter(_process.StandardInput);

                // Create cancellation token for reader threads
                _readerCts = new CancellationTokenSource();

                // Start proxy threads
                StartProxyThreads();
            }
            catch (Exception ex)
            {
                CleanupConnection();
                throw new PSRemotingTransportException("Failed to create TCP connection transport", ex);
            }
        }

        private void StartProxyThreads()
        {
            // Thread to proxy subprocess stdout → network
            var outThread = new Thread(ProcessOutputReaderThread)
            {
                Name = ThreadName + " Out",
                IsBackground = true
            };
            outThread.Start();

            // Thread to proxy network → subprocess stdin (already handled by SetMessageWriter)
            // But we need to monitor network input for disconnection
            var inThread = new Thread(NetworkInputReaderThread)
            {
                Name = ThreadName + " In",
                IsBackground = true
            };
            inThread.Start();

            // Thread to discard subprocess stderr
            var errThread = new Thread(ProcessErrorReaderThread)
            {
                Name = ThreadName + " Err",
                IsBackground = true
            };
            errThread.Start();

            // Send initial fragment
            SendOneItem();
        }

        private void ProcessOutputReaderThread()
        {
            try
            {
                byte[] buffer = new byte[4096];
                int bytesRead;

                while (_readerCts != null && !_readerCts.IsCancellationRequested && 
                       _process != null && _networkStream != null)
                {
                    bytesRead = _process.StandardOutput.BaseStream.Read(buffer, 0, buffer.Length);
                    
                    if (bytesRead <= 0)
                        break;

                    _networkStream.Write(buffer, 0, bytesRead);
                    _networkStream.Flush();
                }
            }
            catch (ObjectDisposedException) { }
            catch (IOException) { }
            catch (OperationCanceledException) { }
            finally
            {
                CleanupConnection();
            }
        }

        private void NetworkInputReaderThread()
        {
            try
            {
                byte[] buffer = new byte[4096];
                int bytesRead;

                while (_readerCts != null && !_readerCts.IsCancellationRequested && 
                       _networkStream != null && _process != null)
                {
                    bytesRead = _networkStream.Read(buffer, 0, buffer.Length);
                    
                    if (bytesRead <= 0)
                        break;

                    _process.StandardInput.BaseStream.Write(buffer, 0, bytesRead);
                    _process.StandardInput.BaseStream.Flush();
                }
            }
            catch (ObjectDisposedException) { }
            catch (IOException) { }
            catch (OperationCanceledException) { }
            finally
            {
                CleanupConnection();
            }
        }

        private void ProcessErrorReaderThread()
        {
            try
            {
                while (_readerCts != null && !_readerCts.IsCancellationRequested && 
                       _process != null)
                {
                    var line = _process.StandardError.ReadLine();
                    if (line == null)
                        break;
                    
                    // Discard error output (following pattern from PSHostClientSessionTransportMgr)
                }
            }
            catch (ObjectDisposedException) { }
            catch (IOException) { }
            catch (OperationCanceledException) { }
        }

        public override void CloseAsync()
        {
            // Cancel reader threads first
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Cleanup connection
            CleanupConnection();

            // Call base implementation
            base.CloseAsync();
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
            // Cancel reader threads first
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Close network stream
            try
            {
                _networkStream?.Close();
                _networkStream?.Dispose();
            }
            catch { }

            // Kill subprocess
            if (_process != null)
            {
                try
                {
                    if (!_process.HasExited)
                    {
                        _process.Kill();
                        _process.WaitForExit(500);
                    }
                }
                catch { }
                finally
                {
                    _process?.Dispose();
                }
            }

            // Close TCP client
            try
            {
                _connectionInfo.TcpClient?.Close();
                _connectionInfo.TcpClient?.Dispose();
            }
            catch { }

            _readerCts?.Dispose();
            _readerCts = null;
            _networkStream = null;
            _process = null;
        }
    }
}
