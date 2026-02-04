using System;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Net.WebSockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Connection info for WebSocket-based PowerShell remoting client connections
    /// </summary>
    internal sealed class PSHostWebSocketClientInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        /// <summary>
        /// The full WebSocket URI (e.g., ws://localhost:8080/pwsh or wss://localhost:8443/pwsh)
        /// </summary>
        public Uri WebSocketUri { get; set; }

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

        public PSHostWebSocketClientInfo(Uri webSocketUri)
        {
            WebSocketUri = webSocketUri;
            ComputerName = webSocketUri.Host;
        }

        /// <summary>
        /// Creates connection info from a WebSocket URI string
        /// </summary>
        public static PSHostWebSocketClientInfo FromUri(string uriString)
        {
            var uri = new Uri(uriString);
            
            // Validate scheme
            if (!uri.Scheme.Equals("ws", StringComparison.OrdinalIgnoreCase) &&
                !uri.Scheme.Equals("wss", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"Invalid WebSocket URI scheme: {uri.Scheme}. Expected 'ws' or 'wss'.");
            }

            return new PSHostWebSocketClientInfo(uri);
        }

        /// <summary>
        /// Creates connection info from individual components
        /// </summary>
        public static PSHostWebSocketClientInfo FromComponents(string hostName, int port, string path, bool useSecureConnection)
        {
            // Ensure path starts with /
            if (!path.StartsWith("/"))
            {
                path = "/" + path;
            }

            string scheme = useSecureConnection ? "wss" : "ws";
            var uri = new Uri($"{scheme}://{hostName}:{port}{path}");
            return new PSHostWebSocketClientInfo(uri);
        }

        public override RunspaceConnectionInfo Clone()
        {
            var connectionInfo = new PSHostWebSocketClientInfo(WebSocketUri)
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
            return new PSHostWebSocketClientSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    /// <summary>
    /// Transport manager for WebSocket client connections to PSHostServer
    /// </summary>
    internal sealed class PSHostWebSocketClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly PSHostWebSocketClientInfo _connectionInfo;
        private ClientWebSocket? _webSocket = null;
        private CancellationTokenSource? _readerCts = null;
        private MemoryStream? _sendBuffer = null;
        private readonly object _sendLock = new object();
        private const string ThreadName = "PSHostWebSocketClient Reader Thread";

        internal PSHostWebSocketClientSessionTransportMgr(
            PSHostWebSocketClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
        }

        public override void CreateAsync()
        {
            _webSocket = new ClientWebSocket();
            _sendBuffer = new MemoryStream();

            // Connect with timeout
            using var cts = new CancellationTokenSource(_connectionInfo.OpenTimeout);
            try
            {
                var connectTask = _webSocket.ConnectAsync(_connectionInfo.WebSocketUri, cts.Token);
                connectTask.Wait(cts.Token);
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException($"WebSocket connection to {_connectionInfo.WebSocketUri} timed out after {_connectionInfo.OpenTimeout}ms");
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
            {
                throw new TimeoutException($"WebSocket connection to {_connectionInfo.WebSocketUri} timed out after {_connectionInfo.OpenTimeout}ms");
            }

            // Create a custom TextWriter that sends data over WebSocket
            var webSocketWriter = new WebSocketTextWriter(this);
            SetMessageWriter(webSocketWriter);

            // Create cancellation token for reader thread
            _readerCts = new CancellationTokenSource();

            // Start the reader thread
            StartReaderThread();

            // Send initial protocol fragment (must be after SetMessageWriter)
            SendOneItem();
        }

        /// <summary>
        /// Send data over WebSocket (called by WebSocketTextWriter)
        /// </summary>
        internal void SendData(string data)
        {
            if (_webSocket == null || _webSocket.State != WebSocketState.Open)
                return;

            lock (_sendLock)
            {
                try
                {
                    var bytes = Encoding.UTF8.GetBytes(data);
                    var sendTask = _webSocket.SendAsync(
                        new ArraySegment<byte>(bytes),
                        WebSocketMessageType.Binary,
                        endOfMessage: true,
                        CancellationToken.None);
                    sendTask.Wait();
                }
                catch (Exception)
                {
                    // Connection closed or error
                }
            }
        }

        public override void CloseAsync()
        {
            // DO NOT close the WebSocket or cancel the reader thread here!
            // The base class needs to send a close packet and receive an ack.
            // The reader thread must remain active to receive the CloseAck.
            // Cleanup happens in CleanupConnection() which is called by the base class.
            
            // Call base implementation - it sends close packet and waits for ack
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
            var buffer = new byte[8192];
            var lineBuffer = new StringBuilder(); // Buffer for accumulating partial lines across messages

            try
            {
                // Start reader loop with cancellation support
                while (_readerCts != null && !_readerCts.IsCancellationRequested &&
                       _webSocket != null && _webSocket.State == WebSocketState.Open)
                {
                    WebSocketReceiveResult result;
                    try
                    {
                        var receiveTask = _webSocket.ReceiveAsync(
                            new ArraySegment<byte>(buffer),
                            _readerCts.Token);
                        result = receiveTask.GetAwaiter().GetResult();
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }

                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        break;
                    }

                    if (result.Count > 0)
                    {
                        // Convert received bytes to string
                        var data = Encoding.UTF8.GetString(buffer, 0, result.Count);
                        
                        // Append to line buffer
                        lineBuffer.Append(data);
                        
                        // Process complete lines from the buffer
                        // The server sends raw bytes which may contain partial lines
                        string bufferContent = lineBuffer.ToString();
                        int lastNewlineIndex = bufferContent.LastIndexOf('\n');
                        
                        if (lastNewlineIndex >= 0)
                        {
                            // Extract complete lines (everything up to and including last newline)
                            string completeData = bufferContent.Substring(0, lastNewlineIndex + 1);
                            
                            // Keep any remaining partial line in the buffer
                            if (lastNewlineIndex < bufferContent.Length - 1)
                            {
                                lineBuffer.Clear();
                                lineBuffer.Append(bufferContent.Substring(lastNewlineIndex + 1));
                            }
                            else
                            {
                                lineBuffer.Clear();
                            }
                            
                            // Process each complete line
                            using var reader = new StringReader(completeData);
                            string? line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (!string.IsNullOrEmpty(line))
                                {
                                    HandleOutputDataReceived(line);
                                }
                            }
                        }
                        // If no newline found, data stays in buffer waiting for more
                    }
                }
            }
            catch (ObjectDisposedException)
            {
                // Normal reader thread end
            }
            catch (WebSocketException)
            {
                // Connection closed or error
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

            // Close WebSocket
            if (_webSocket != null)
            {
                try
                {
                    if (_webSocket.State == WebSocketState.Open)
                    {
                        _webSocket.CloseAsync(
                            WebSocketCloseStatus.NormalClosure,
                            "Cleanup",
                            CancellationToken.None).Wait(1000);
                    }
                }
                catch { }

                try
                {
                    _webSocket.Dispose();
                }
                catch { }
            }

            try
            {
                _sendBuffer?.Dispose();
            }
            catch { }

            _readerCts?.Dispose();
            _readerCts = null;
            _sendBuffer = null;
            _webSocket = null;
        }
    }

    /// <summary>
    /// Custom TextWriter that sends data over WebSocket
    /// </summary>
    internal sealed class WebSocketTextWriter : TextWriter
    {
        private readonly PSHostWebSocketClientSessionTransportMgr _transportMgr;
        private readonly StringBuilder _lineBuffer = new StringBuilder();

        public override Encoding Encoding => Encoding.UTF8;

        public WebSocketTextWriter(PSHostWebSocketClientSessionTransportMgr transportMgr)
        {
            _transportMgr = transportMgr;
        }

        public override void Write(char value)
        {
            _lineBuffer.Append(value);
            
            // Send on newline (PSRP protocol is line-based)
            if (value == '\n')
            {
                Flush();
            }
        }

        public override void Write(string? value)
        {
            if (value == null) return;
            
            _lineBuffer.Append(value);
            
            // Check if ends with newline
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
            if (_lineBuffer.Length > 0)
            {
                _transportMgr.SendData(_lineBuffer.ToString());
                _lineBuffer.Clear();
            }
        }
    }
}
