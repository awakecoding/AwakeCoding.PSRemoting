using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.WebSockets;
using System.Threading;
using System.Threading.Tasks;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// WebSocket-based PowerShell remoting server using HttpListener
    /// </summary>
    public class PSHostWebSocketServer : PSHostServerBase
    {
        private HttpListener? _listener;        private string _httpListenerPrefix = string.Empty;        private const string ThreadName = "PSHostWebSocketServer Listener";

        /// <summary>
        /// URL path for WebSocket connections (e.g., "/pwsh")
        /// </summary>
        public string Path { get; private set; }

        /// <summary>
        /// Whether to use secure WebSocket (wss://)
        /// </summary>
        public bool UseSecureConnection { get; private set; }

        /// <summary>
        /// Full URI prefix the server is listening on
        /// </summary>
        public string ListenerPrefix { get; private set; }

        public PSHostWebSocketServer(
            string name, 
            int port, 
            string listenAddress, 
            string path,
            int maxConnections, 
            int drainTimeout,
            bool useSecureConnection)
            : base(name, port, listenAddress, maxConnections, drainTimeout)
        {
            Path = path.StartsWith("/") ? path : "/" + path;
            UseSecureConnection = useSecureConnection;
            
            // Build the listener prefix
            // HttpListener requires trailing slash
            // Use '+' for wildcard binding (listens on all interfaces for the specified port)
            // This allows clients to connect via localhost, 127.0.0.1, or any other network interface
            string scheme = useSecureConnection ? "https" : "http";
            string bindHost = "+"; // Always use '+' for wildcard binding
            string httpListenerPrefix = $"{scheme}://{bindHost}:{port}{Path}/";
            
            // Store the prefix for HttpListener (with + for wildcard)
            _httpListenerPrefix = httpListenerPrefix;
            
            // Store display prefix (with actual listen address)
            ListenerPrefix = $"{scheme}://{listenAddress}:{port}{Path}/";
        }

        public override void StartListenerAsync()
        {
            if (State == ServerState.Running || State == ServerState.Starting)
            {
                throw new InvalidOperationException($"Server '{Name}' is already {State}");
            }

            State = ServerState.Starting;

            try
            {
                _listener = new HttpListener();
                _listener.Prefixes.Add(_httpListenerPrefix);
                _listener.Start();

                // Create cancellation token source
                _serverInstance.CancellationTokenSource = new CancellationTokenSource();

                // Start listener thread
                _serverInstance.ListenerThread = new Thread(ListenerThreadProc)
                {
                    Name = ThreadName,
                    IsBackground = true
                };
                _serverInstance.ListenerThread.Start();

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
            if (State != ServerState.Running && State != ServerState.Starting)
            {
                return;
            }

            State = ServerState.Stopping;

            try
            {
                // Signal cancellation to stop accepting new connections
                _serverInstance.CancellationTokenSource?.Cancel();

                // Stop the HTTP listener
                _listener?.Stop();
                _listener?.Close();

                if (!force && ConnectionCount > 0)
                {
                    // Wait for drain timeout for connections to complete
                    var drainStart = DateTime.UtcNow;
                    var timeout = TimeSpan.FromSeconds(DrainTimeout);
                    
                    while (ConnectionCount > 0 && (DateTime.UtcNow - drainStart) < timeout)
                    {
                        Thread.Sleep(100);
                    }
                }

                // Force kill any remaining connections and their subprocesses
                var connections = _serverInstance.ActiveConnections.Values;
                foreach (var connection in connections)
                {
                    if (connection.ProcessId.HasValue)
                    {
                        try
                        {
                            var process = Process.GetProcessById(connection.ProcessId.Value);
                            if (!process.HasExited)
                            {
                                process.Kill();
                                process.WaitForExit(500);
                            }
                        }
                        catch { }
                    }
                }

                // Clear all connections
                _serverInstance.ActiveConnections.Clear();

                // Wait for listener thread to exit
                _serverInstance.ListenerThread?.Join(1000);

                State = ServerState.Stopped;
            }
            catch (Exception ex)
            {
                State = ServerState.Failed;
                LastError = ex;
                throw;
            }
            finally
            {
                _serverInstance.CancellationTokenSource?.Dispose();
                _serverInstance.CancellationTokenSource = null;
                _listener = null;
            }
        }

        private void ListenerThreadProc()
        {
            var cancellationToken = _serverInstance.CancellationTokenSource?.Token ?? CancellationToken.None;

            try
            {
                while (!cancellationToken.IsCancellationRequested && _listener != null && _listener.IsListening)
                {
                    try
                    {
                        // Accept HTTP context asynchronously
                        var contextTask = _listener.GetContextAsync();
                        
                        // Wait with timeout to check cancellation
                        while (!contextTask.Wait(100))
                        {
                            if (cancellationToken.IsCancellationRequested)
                                return;
                        }

                        var context = contextTask.Result;

                        // Check if we've reached max connections before upgrading to WebSocket
                        if (IsMaxConnectionsReached())
                        {
                            // Reject the request with 503 Service Unavailable
                            context.Response.StatusCode = 503;
                            context.Response.StatusDescription = "Server at maximum capacity";
                            context.Response.Close();
                            continue;
                        }

                        // Check if it's a WebSocket request
                        if (context.Request.IsWebSocketRequest)
                        {
                            // Handle WebSocket connection in a separate task
                            Task.Run(() => HandleWebSocketConnection(context, cancellationToken));
                        }
                        else
                        {
                            // Return 400 for non-WebSocket requests
                            context.Response.StatusCode = 400;
                            context.Response.StatusDescription = "WebSocket connection required";
                            context.Response.Close();
                        }
                    }
                    catch (HttpListenerException) when (cancellationToken.IsCancellationRequested)
                    {
                        // Expected during shutdown
                        break;
                    }
                    catch (ObjectDisposedException)
                    {
                        // Listener was disposed
                        break;
                    }
                    catch (Exception)
                    {
                        // Log and continue accepting
                        if (cancellationToken.IsCancellationRequested)
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!cancellationToken.IsCancellationRequested)
                {
                    State = ServerState.Failed;
                    LastError = ex;
                }
            }
        }

        private async Task HandleWebSocketConnection(HttpListenerContext context, CancellationToken cancellationToken)
        {
            WebSocket? webSocket = null;
            Process? process = null;
            string connectionId = Guid.NewGuid().ToString();

            try
            {
                // Accept the WebSocket upgrade
                var wsContext = await context.AcceptWebSocketAsync(subProtocol: null);
                webSocket = wsContext.WebSocket;

                var clientEndpoint = context.Request.RemoteEndPoint?.ToString() ?? "unknown";

                // Spawn PowerShell subprocess
                var executable = PowerShellFinder.GetPowerShellPath();
                if (executable == null || !File.Exists(executable))
                {
                    await webSocket.CloseAsync(
                        WebSocketCloseStatus.InternalServerError,
                        "PowerShell executable not found",
                        CancellationToken.None);
                    return;
                }

                process = new Process();
                process.StartInfo.FileName = executable;
                process.StartInfo.Arguments = "-NoLogo -NoProfile -s";
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;

                process.Start();

                // Add connection to tracking
                var connectionDetails = new ConnectionDetails(connectionId, clientEndpoint, process.Id);
                AddConnection(connectionDetails);

                // Start proxy tasks
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                
                var wsToProcess = ProxyWebSocketToProcess(webSocket, process.StandardInput.BaseStream, linkedCts.Token);
                var processToWs = ProxyProcessToWebSocket(process.StandardOutput.BaseStream, webSocket, linkedCts.Token);
                var discardErrors = DiscardProcessErrors(process.StandardError, linkedCts.Token);

                // Wait for either direction to complete (indicates connection closed)
                await Task.WhenAny(wsToProcess, processToWs);
                
                // Cancel the other direction
                linkedCts.Cancel();
            }
            catch (WebSocketException)
            {
                // Client disconnected
            }
            catch (Exception)
            {
                // Connection error
            }
            finally
            {
                // Cleanup
                CleanupConnection(connectionId, webSocket, process);
            }
        }

        private async Task ProxyWebSocketToProcess(WebSocket webSocket, Stream processInput, CancellationToken cancellationToken)
        {
            var buffer = new byte[4096];

            try
            {
                while (webSocket.State == WebSocketState.Open && !cancellationToken.IsCancellationRequested)
                {
                    var result = await webSocket.ReceiveAsync(new ArraySegment<byte>(buffer), cancellationToken);

                    if (result.MessageType == WebSocketMessageType.Close)
                    {
                        // Respond to close message
                        if (webSocket.State == WebSocketState.Open || webSocket.State == WebSocketState.CloseReceived)
                        {
                            await webSocket.CloseAsync(
                                WebSocketCloseStatus.NormalClosure,
                                "Closing",
                                CancellationToken.None);
                        }
                        break;
                    }

                    if (result.Count > 0)
                    {
                        await processInput.WriteAsync(buffer, 0, result.Count, cancellationToken);
                        await processInput.FlushAsync(cancellationToken);
                    }
                }
            }
            catch (OperationCanceledException) { }
            catch (WebSocketException) { }
            catch (IOException) { }
        }

        private async Task ProxyProcessToWebSocket(Stream processOutput, WebSocket webSocket, CancellationToken cancellationToken)
        {
            var buffer = new byte[4096];

            try
            {
                while (webSocket.State == WebSocketState.Open && !cancellationToken.IsCancellationRequested)
                {
                    var bytesRead = await processOutput.ReadAsync(buffer, 0, buffer.Length, cancellationToken);

                    if (bytesRead <= 0)
                    {
                        break;
                    }

                    await webSocket.SendAsync(
                        new ArraySegment<byte>(buffer, 0, bytesRead),
                        WebSocketMessageType.Binary,
                        endOfMessage: true,
                        cancellationToken);
                }
            }
            catch (OperationCanceledException) { }
            catch (WebSocketException) { }
            catch (IOException) { }
        }

        private async Task DiscardProcessErrors(StreamReader errorReader, CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    var line = await errorReader.ReadLineAsync(cancellationToken);
                    if (line == null)
                        break;
                    // Discard error output
                }
            }
            catch (OperationCanceledException) { }
            catch (IOException) { }
        }

        private void CleanupConnection(string connectionId, WebSocket? webSocket, Process? process)
        {
            // Close WebSocket gracefully if still open
            if (webSocket != null && webSocket.State == WebSocketState.Open)
            {
                try
                {
                    webSocket.CloseAsync(
                        WebSocketCloseStatus.NormalClosure,
                        "Server closing connection",
                        CancellationToken.None).Wait(1000);
                }
                catch { }
            }

            try
            {
                webSocket?.Dispose();
            }
            catch { }

            // Kill subprocess
            if (process != null)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill();
                        process.WaitForExit(500);
                    }
                }
                catch { }
                finally
                {
                    process.Dispose();
                }
            }

            // Remove from tracking
            RemoveConnection(connectionId);
        }

        protected override void AcceptConnectionAsync()
        {
            // Not used - WebSocket connections are handled via HandleWebSocketConnection
            throw new NotImplementedException("WebSocket connections are handled asynchronously");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _listener?.Stop();
                _listener?.Close();
            }

            base.Dispose(disposing);
        }
    }
}
