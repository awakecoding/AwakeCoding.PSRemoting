using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Manages a single remote PowerShell shell session within the WinRM server.
    /// Wraps a pwsh subprocess and exposes stdin/stdout via thread-safe queues.
    /// </summary>
    internal sealed class WinRMShell : IDisposable
    {
        public string ShellId { get; }

        private Process? _process;
        private readonly BlockingCollection<string> _stdoutQueue;
        private Thread? _stdoutReaderThread;
        private Thread? _stderrReaderThread;
        private volatile bool _disposed;

        public WinRMShell(string shellId, Process process)
        {
            ShellId = shellId ?? throw new ArgumentNullException(nameof(shellId));
            _process = process ?? throw new ArgumentNullException(nameof(process));
            _stdoutQueue = new BlockingCollection<string>(1024);
        }

        /// <summary>
        /// Starts the background reader threads that drain pwsh stdout/stderr.
        /// Must be called after the process has been started.
        /// </summary>
        public void StartReaders()
        {
            _stdoutReaderThread = new Thread(ReadStdout)
            {
                Name = $"WinRM Shell {ShellId} stdout reader",
                IsBackground = true
            };
            _stderrReaderThread = new Thread(ReadStderr)
            {
                Name = $"WinRM Shell {ShellId} stderr reader",
                IsBackground = true
            };
            _stdoutReaderThread.Start();
            _stderrReaderThread.Start();
        }

        /// <summary>
        /// Write PSRP data to the subprocess stdin.
        /// </summary>
        public void WriteToStdin(string data)
        {
            if (_disposed || _process == null || _process.HasExited) return;
            try
            {
                _process.StandardInput.WriteLine(data);
                _process.StandardInput.Flush();
            }
            catch { }
        }

        /// <summary>
        /// Try to dequeue a line from stdout within the given timeout.
        /// Returns null if the timeout expires or the shell has closed.
        /// </summary>
        public string? TryDequeueStdoutLine(int timeoutMs)
        {
            if (_disposed) return null;
            try
            {
                if (_stdoutQueue.TryTake(out string? line, timeoutMs))
                    return line;
                return null;
            }
            catch (InvalidOperationException)
            {
                // Collection was marked completed
                return null;
            }
        }

        private void ReadStdout()
        {
            try
            {
                while (!_disposed && _process != null && !_process.HasExited)
                {
                    string? line = _process.StandardOutput.ReadLine();
                    if (line == null) break;
                    try { _stdoutQueue.Add(line); }
                    catch (InvalidOperationException) { break; }
                }
            }
            catch (IOException) { }
            catch (ObjectDisposedException) { }
            finally
            {
                try { _stdoutQueue.CompleteAdding(); } catch { }
            }
        }

        private void ReadStderr()
        {
            try
            {
                while (!_disposed && _process != null && !_process.HasExited)
                {
                    string? line = _process.StandardError.ReadLine();
                    if (line == null) break;
                    // Discard stderr - swallowed to avoid blocking
                }
            }
            catch (IOException) { }
            catch (ObjectDisposedException) { }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try { _stdoutQueue.CompleteAdding(); } catch { }

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
                    _process.Dispose();
                    _process = null;
                }
            }

            _stdoutQueue.Dispose();
        }
    }

    /// <summary>
    /// WinRM server that listens for WSMan/SOAP requests over HTTP using HttpListener.
    /// Implements the Create/Send/Receive/Delete shell operations to host PowerShell
    /// remoting sessions without requiring system WinRM to be configured.
    /// </summary>
    public class PSHostWinRMServer : PSHostServerBase
    {
        private HttpListener? _listener;
        private readonly ConcurrentDictionary<string, WinRMShell> _shells;
        private string _httpListenerPrefix = string.Empty;
        private const string ThreadName = "PSHostWinRMServer Listener";

        /// <summary>URL path for WSMan requests (typically "/wsman").</summary>
        public string Path { get; private set; }

        /// <summary>Optional credential required for Basic auth. If null, authentication is skipped.</summary>
        public PSCredential? RequiredCredential { get; private set; }

        /// <summary>Full URI prefix the server is listening on.</summary>
        public string ListenerPrefix { get; private set; }

        public PSHostWinRMServer(
            string name,
            int port,
            string listenAddress,
            string path,
            int maxConnections,
            int drainTimeout,
            PSCredential? requiredCredential)
            : base(name, port, listenAddress, maxConnections, drainTimeout)
        {
            Path = path.StartsWith("/") ? path : "/" + path;
            RequiredCredential = requiredCredential;
            _shells = new ConcurrentDictionary<string, WinRMShell>(StringComparer.OrdinalIgnoreCase);

            // Build HttpListener prefix
            string bindHost = DetermineBindHost(listenAddress);
            _httpListenerPrefix = $"http://{bindHost}:{port}{Path}/";
            ListenerPrefix = $"http://{listenAddress}:{port}{Path}/";
        }

        private static string DetermineBindHost(string listenAddress)
        {
            if (listenAddress == "127.0.0.1" ||
                listenAddress.Equals("localhost", StringComparison.OrdinalIgnoreCase) ||
                listenAddress == "::1")
                return "localhost";

            if (listenAddress == "0.0.0.0" || listenAddress == "*" || listenAddress == "+")
                return "+";

            return listenAddress;
        }

        public override void StartListenerAsync()
        {
            if (State == ServerState.Running || State == ServerState.Starting)
                throw new InvalidOperationException($"Server '{Name}' is already {State}");

            State = ServerState.Starting;

            try
            {
                _listener = new HttpListener();
                _listener.Prefixes.Add(_httpListenerPrefix);
                _listener.Start();

                _serverInstance.CancellationTokenSource = new CancellationTokenSource();

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
            if (!TryBeginStopping()) return;

            if (State != ServerState.Running && State != ServerState.Starting)
            {
                ResetStoppingFlag();
                return;
            }

            State = ServerState.Stopping;

            try
            {
                _serverInstance.CancellationTokenSource?.Cancel();
                _listener?.Stop();
                _listener?.Close();

                if (!force && ConnectionCount > 0)
                {
                    var drainStart = DateTime.UtcNow;
                    var timeout = TimeSpan.FromSeconds(DrainTimeout);
                    while (ConnectionCount > 0 && (DateTime.UtcNow - drainStart) < timeout)
                        Thread.Sleep(100);
                }

                // Kill all active shells
                foreach (var kv in _shells)
                {
                    try { kv.Value.Dispose(); } catch { }
                }
                _shells.Clear();

                // Force kill any connection-tracked subprocesses
                foreach (var connection in _serverInstance.ActiveConnections.Values)
                {
                    if (connection.ProcessId.HasValue)
                    {
                        try
                        {
                            var proc = Process.GetProcessById(connection.ProcessId.Value);
                            if (!proc.HasExited)
                            {
                                proc.Kill();
                                proc.WaitForExit(ProcessKillWaitTimeoutMs);
                            }
                        }
                        catch { }
                    }
                }

                _serverInstance.ActiveConnections.Clear();
                _serverInstance.ListenerThread?.Join(ListenerThreadJoinTimeoutMs);

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
                _listener = null;
                ResetStoppingFlag();
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
                        var contextTask = _listener.GetContextAsync();
                        while (!contextTask.Wait(100))
                        {
                            if (cancellationToken.IsCancellationRequested)
                                return;
                        }

                        var context = contextTask.Result;
                        Task.Run(() => HandleRequestAsync(context, cancellationToken));
                    }
                    catch (HttpListenerException) when (cancellationToken.IsCancellationRequested)
                    {
                        break;
                    }
                    catch (ObjectDisposedException)
                    {
                        break;
                    }
                    catch (Exception)
                    {
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

        private async Task HandleRequestAsync(HttpListenerContext ctx, CancellationToken cancellationToken)
        {
            try
            {
                // Only accept POST requests
                if (!ctx.Request.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase))
                {
                    SendHttpResponse(ctx, 405, "Method Not Allowed", "text/plain", "Only POST is supported");
                    return;
                }

                // Authenticate if a credential is configured
                if (RequiredCredential != null && !AuthenticateBasic(ctx.Request))
                {
                    ctx.Response.StatusCode = 401;
                    ctx.Response.Headers.Add("WWW-Authenticate", "Basic realm=\"wsman\"");
                    ctx.Response.ContentLength64 = 0;
                    ctx.Response.Close();
                    return;
                }

                // Read the request body (SOAP XML)
                string requestBody;
                using (var reader = new StreamReader(ctx.Request.InputStream, Encoding.UTF8))
                    requestBody = await reader.ReadToEndAsync().ConfigureAwait(false);

                // Determine the WSMan action
                string action = ExtractSoapAction(requestBody, ctx.Request);

                string responseXml;

                if (action.EndsWith("/Transfer/Create", StringComparison.OrdinalIgnoreCase) ||
                    action == WsManXml.ActionCreate)
                {
                    responseXml = HandleCreate(ctx.Request);
                }
                else if (action == WsManXml.ActionSend)
                {
                    responseXml = HandleSend(requestBody);
                }
                else if (action == WsManXml.ActionReceive)
                {
                    responseXml = HandleReceive(requestBody);
                }
                else if (action == WsManXml.ActionDelete ||
                         action.EndsWith("/Transfer/Delete", StringComparison.OrdinalIgnoreCase))
                {
                    responseXml = HandleDelete(requestBody);
                }
                else
                {
                    SendHttpResponse(ctx, 400, "Bad Request", "text/plain",
                        $"Unsupported WSMan action: {action}");
                    return;
                }

                SendHttpResponse(ctx, 200, "OK", "application/soap+xml;charset=UTF-8", responseXml);
            }
            catch (Exception ex)
            {
                try
                {
                    SendHttpResponse(ctx, 500, "Internal Server Error", "text/plain", ex.Message);
                }
                catch { }
            }
        }

        private bool AuthenticateBasic(HttpListenerRequest request)
        {
            if (RequiredCredential == null) return true;

            string? authHeader = request.Headers["Authorization"];
            if (string.IsNullOrEmpty(authHeader) || !authHeader.StartsWith("Basic ", StringComparison.OrdinalIgnoreCase))
                return false;

            try
            {
                string encoded = authHeader.Substring(6).Trim();
                string decoded = Encoding.ASCII.GetString(Convert.FromBase64String(encoded));
                int sep = decoded.IndexOf(':');
                if (sep < 0) return false;

                string user = decoded.Substring(0, sep);
                string pass = decoded.Substring(sep + 1);

                var netCred = RequiredCredential.GetNetworkCredential();
                return string.Equals(user, netCred.UserName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(pass, netCred.Password, StringComparison.Ordinal);
            }
            catch { return false; }
        }

        private static string ExtractSoapAction(string soapXml, HttpListenerRequest request)
        {
            // Try SOAPAction HTTP header first
            string? soapActionHeader = request.Headers["SOAPAction"];
            if (!string.IsNullOrEmpty(soapActionHeader))
                return soapActionHeader.Trim('"');

            // Parse from SOAP body
            try
            {
                var doc = XDocument.Parse(soapXml);
                XNamespace wsa = WsManXml.NsWsa;
                var actionEl = doc.Descendants(wsa + "Action").GetEnumerator();
                if (actionEl.MoveNext())
                    return actionEl.Current.Value;
            }
            catch { }

            return string.Empty;
        }

        private string HandleCreate(HttpListenerRequest request)
        {
            string shellId = Guid.NewGuid().ToString().ToUpper();

            // Spawn pwsh subprocess
            string? executable = PowerShellFinder.GetPowerShellPath();
            if (executable == null || !File.Exists(executable))
                throw new InvalidOperationException("PowerShell executable not found on this system");

            var process = new Process();
            process.StartInfo.FileName = executable;
            process.StartInfo.Arguments = "-NoLogo -NoProfile -s";
            process.StartInfo.RedirectStandardInput = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            process.Start();

            var shell = new WinRMShell(shellId, process);
            shell.StartReaders();
            _shells[shellId] = shell;

            // Track the connection
            string clientEndpoint = request.RemoteEndPoint?.ToString() ?? "unknown";
            var connectionDetails = new ConnectionDetails(shellId, clientEndpoint, process.Id);
            AddConnection(connectionDetails);

            return BuildCreateResponse(shellId);
        }

        private string HandleSend(string soapXml)
        {
            string? shellId = ParseShellIdFromRequest(soapXml);
            if (shellId == null || !_shells.TryGetValue(shellId, out var shell))
                throw new InvalidOperationException($"Shell not found: {shellId}");

            // Decode stdin stream data and write to pwsh
            try
            {
                var doc = XDocument.Parse(soapXml);
                XNamespace rsp = WsManXml.NsRsp;

                foreach (var stream in doc.Descendants(rsp + "Stream"))
                {
                    string? name = (string?)stream.Attribute("Name");
                    if (name != "stdin") continue;

                    string? base64Data = stream.Value;
                    if (string.IsNullOrEmpty(base64Data)) continue;

                    byte[] bytes = Convert.FromBase64String(base64Data);
                    string text = Encoding.UTF8.GetString(bytes);

                    // Write each line to stdin (text already ends with \n from client encoding)
                    foreach (string line in text.Split('\n'))
                    {
                        string trimmed = line.TrimEnd('\r');
                        if (!string.IsNullOrEmpty(trimmed))
                            shell.WriteToStdin(trimmed);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to process Send request: {ex.Message}", ex);
            }

            return BuildSendResponse();
        }

        private string HandleReceive(string soapXml)
        {
            string? shellId = ParseShellIdFromRequest(soapXml);
            if (shellId == null || !_shells.TryGetValue(shellId, out var shell))
                throw new InvalidOperationException($"Shell not found: {shellId}");

            // Wait up to 500ms for the first stdout line, then return what we have.
            // The client retries immediately on an empty response; a short timeout keeps
            // latency low without busy-looping on the server side.
            string? firstLine = shell.TryDequeueStdoutLine(timeoutMs: 500);

            var lines = new System.Collections.Generic.List<string>();
            if (firstLine != null)
            {
                lines.Add(firstLine);

                // Drain any additional immediately-available lines (non-blocking)
                string? extraLine;
                while ((extraLine = shell.TryDequeueStdoutLine(0)) != null)
                    lines.Add(extraLine);
            }

            return BuildReceiveResponse(shellId, lines);
        }

        private string HandleDelete(string soapXml)
        {
            string? shellId = ParseShellIdFromRequest(soapXml);

            if (shellId != null && _shells.TryRemove(shellId, out var shell))
            {
                int? processId = null;
                try
                {
                    // Get process ID before disposing (for tracking cleanup)
                    var connection = _serverInstance.ActiveConnections.TryGetValue(shellId, out var conn) ? conn : null;
                    processId = connection?.ProcessId;
                }
                catch { }

                shell.Dispose();
                RemoveConnection(shellId);
            }

            return BuildDeleteResponse();
        }

        private static string? ParseShellIdFromRequest(string soapXml)
        {
            try
            {
                var doc = XDocument.Parse(soapXml);
                XNamespace wsman = WsManXml.NsWsMan;

                foreach (var sel in doc.Descendants(wsman + "Selector"))
                {
                    if ((string?)sel.Attribute("Name") == "ShellId")
                        return sel.Value;
                }
            }
            catch { }

            return null;
        }

        private static void SendHttpResponse(
            HttpListenerContext ctx,
            int statusCode,
            string statusDescription,
            string contentType,
            string body)
        {
            try
            {
                byte[] bytes = Encoding.UTF8.GetBytes(body);
                ctx.Response.StatusCode = statusCode;
                ctx.Response.StatusDescription = statusDescription;
                ctx.Response.ContentType = contentType;
                ctx.Response.ContentLength64 = bytes.Length;
                ctx.Response.OutputStream.Write(bytes, 0, bytes.Length);
                ctx.Response.Close();
            }
            catch { }
        }

        // --- SOAP response builders ---

        private static readonly string SoapNs =
            "xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"" +
            " xmlns:wsa=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"" +
            " xmlns:wsman=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"" +
            " xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\"";

        private static string BuildCreateResponse(string shellId)
        {
            string msgId = Guid.NewGuid().ToString().ToUpper();
            return $"<s:Envelope {SoapNs}>" +
                   "<s:Header>" +
                   $"<wsa:Action>http://schemas.xmlsoap.org/ws/2004/09/transfer/CreateResponse</wsa:Action>" +
                   $"<wsa:MessageID>uuid:{msgId}</wsa:MessageID>" +
                   $"<wsman:SelectorSet><wsman:Selector Name=\"ShellId\">{shellId}</wsman:Selector></wsman:SelectorSet>" +
                   "</s:Header>" +
                   "<s:Body>" +
                   "<rsp:Shell>" +
                   $"<rsp:ShellId>{shellId}</rsp:ShellId>" +
                   "<rsp:InputStreams>stdin</rsp:InputStreams>" +
                   "<rsp:OutputStreams>stdout stderr</rsp:OutputStreams>" +
                   "</rsp:Shell>" +
                   "</s:Body>" +
                   "</s:Envelope>";
        }

        private static string BuildSendResponse()
        {
            string msgId = Guid.NewGuid().ToString().ToUpper();
            return $"<s:Envelope {SoapNs}>" +
                   "<s:Header>" +
                   $"<wsa:Action>http://schemas.microsoft.com/wbem/wsman/1/windows/shell/SendResponse</wsa:Action>" +
                   $"<wsa:MessageID>uuid:{msgId}</wsa:MessageID>" +
                   "</s:Header>" +
                   "<s:Body><rsp:SendResponse/></s:Body>" +
                   "</s:Envelope>";
        }

        private static string BuildReceiveResponse(string shellId, System.Collections.Generic.List<string> stdoutLines)
        {
            string msgId = Guid.NewGuid().ToString().ToUpper();
            var sb = new StringBuilder();
            sb.Append($"<s:Envelope {SoapNs}>");
            sb.Append("<s:Header>");
            sb.Append($"<wsa:Action>http://schemas.microsoft.com/wbem/wsman/1/windows/shell/ReceiveResponse</wsa:Action>");
            sb.Append($"<wsa:MessageID>uuid:{msgId}</wsa:MessageID>");
            sb.Append("</s:Header>");
            sb.Append("<s:Body><rsp:ReceiveResponse>");

            foreach (string line in stdoutLines)
            {
                // Encode line + \n so client decodes to complete PSRP lines
                string base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(line + "\n"));
                sb.Append($"<rsp:Stream Name=\"stdout\" CommandId=\"{shellId}\">{base64}</rsp:Stream>");
            }

            sb.Append("</rsp:ReceiveResponse></s:Body></s:Envelope>");
            return sb.ToString();
        }

        private static string BuildDeleteResponse()
        {
            string msgId = Guid.NewGuid().ToString().ToUpper();
            return $"<s:Envelope {SoapNs}>" +
                   "<s:Header>" +
                   $"<wsa:Action>http://schemas.xmlsoap.org/ws/2004/09/transfer/DeleteResponse</wsa:Action>" +
                   $"<wsa:MessageID>uuid:{msgId}</wsa:MessageID>" +
                   "</s:Header>" +
                   "<s:Body/>" +
                   "</s:Envelope>";
        }

        protected override void AcceptConnectionAsync()
        {
            throw new NotImplementedException("WinRM connections are handled asynchronously via HandleRequestAsync");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _listener?.Stop();
                _listener?.Close();
                foreach (var kv in _shells)
                    try { kv.Value.Dispose(); } catch { }
                _shells.Clear();
            }
            base.Dispose(disposing);
        }
    }
}
