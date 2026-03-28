using System;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Authentication mechanism for WinRM connections
    /// </summary>
    public enum WinRMAuthenticationMechanism
    {
        /// <summary>Negotiate (NTLM or Kerberos)</summary>
        Default,
        /// <summary>HTTP Basic authentication (username + password in Base64)</summary>
        Basic,
        /// <summary>Negotiate authentication (NTLM or Kerberos via SSPI/GSSAPI)</summary>
        Negotiate,
        /// <summary>Kerberos-only authentication</summary>
        Kerberos
    }

    /// <summary>
    /// Connection info for WinRM-based PowerShell remoting connections.
    /// Uses HttpClient to implement the WSMan/SOAP protocol over HTTP/HTTPS.
    /// </summary>
    public sealed class WinRMClientInfo : RunspaceConnectionInfo
    {
        private const int DefaultHttpPort = 5985;
        private const int DefaultHttpsPort = 5986;
        private const string DefaultApplicationName = "/wsman";

        public override string ComputerName { get; set; }

        /// <summary>Port number. If 0, uses 5985 (HTTP) or 5986 (HTTPS).</summary>
        public int Port { get; set; }

        /// <summary>Use HTTPS (port 5986 by default).</summary>
        public bool UseSSL { get; set; }

        /// <summary>Authentication mechanism (Basic, Negotiate, etc.).</summary>
        public WinRMAuthenticationMechanism Authentication { get; set; }

        /// <summary>WSMan application name path (default "/wsman").</summary>
        public string ApplicationName { get; set; }

        /// <summary>Whether to follow HTTP redirects.</summary>
        public bool AllowRedirection { get; set; }

        public override PSCredential? Credential { get; set; }

        public override AuthenticationMechanism AuthenticationMechanism
        {
            get => AuthenticationMechanism.Default;
            set => throw new NotImplementedException();
        }

        public override string CertificateThumbprint
        {
            get => string.Empty;
            set => throw new NotImplementedException();
        }

        public WinRMClientInfo(string computerName)
        {
            ComputerName = computerName ?? throw new ArgumentNullException(nameof(computerName));
            ApplicationName = DefaultApplicationName;
            Authentication = WinRMAuthenticationMechanism.Basic;
        }

        /// <summary>Effective port: uses 5985/5986 if Port is 0.</summary>
        public int EffectivePort => Port > 0 ? Port : (UseSSL ? DefaultHttpsPort : DefaultHttpPort);

        /// <summary>Full HTTP/HTTPS endpoint URI for WSMan requests.</summary>
        public string BuildEndpointUri()
        {
            string scheme = UseSSL ? "https" : "http";
            return $"{scheme}://{ComputerName}:{EffectivePort}{ApplicationName}";
        }

        public override RunspaceConnectionInfo Clone()
        {
            return new WinRMClientInfo(ComputerName)
            {
                Port = Port,
                UseSSL = UseSSL,
                Authentication = Authentication,
                ApplicationName = ApplicationName,
                AllowRedirection = AllowRedirection,
                Credential = Credential
            };
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new WinRMClientSessionTransportMgr(this, instanceId, cryptoHelper);
        }
    }

    /// <summary>
    /// Helpers for building and parsing WSMan SOAP XML messages.
    /// Uses string templates to avoid XML namespace declaration issues.
    /// </summary>
    internal static class WsManXml
    {
        // WSMan namespace URIs
        public const string NsSoap = "http://www.w3.org/2003/05/soap-envelope";
        public const string NsWsa  = "http://schemas.xmlsoap.org/ws/2004/08/addressing";
        public const string NsWsMan = "http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd";
        public const string NsRsp  = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell";

        public const string ResourceUriPowerShell =
            "http://schemas.microsoft.com/powershell/Microsoft.PowerShell";

        public const string ActionCreate  = "http://schemas.xmlsoap.org/ws/2004/09/transfer/Create";
        public const string ActionDelete  = "http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete";
        public const string ActionSend    = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Send";
        public const string ActionReceive = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive";

        private const string SoapPrologue =
            "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"" +
            " xmlns:wsa=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"" +
            " xmlns:wsman=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"" +
            " xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\">";

        private static string BuildHeader(string to, string action, string msgId, string? shellId)
        {
            var sb = new StringBuilder();
            sb.Append("<s:Header>");
            sb.Append($"<wsa:To>{to}</wsa:To>");
            sb.Append($"<wsman:ResourceURI>{ResourceUriPowerShell}</wsman:ResourceURI>");
            sb.Append("<wsa:ReplyTo><wsa:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</wsa:Address></wsa:ReplyTo>");
            sb.Append($"<wsa:Action>{action}</wsa:Action>");
            sb.Append($"<wsa:MessageID>uuid:{msgId}</wsa:MessageID>");
            sb.Append("<wsman:MaxEnvelopeSize>153600</wsman:MaxEnvelopeSize>");
            sb.Append("<wsman:Locale xml:lang=\"en-US\"/>");
            sb.Append("<wsman:DataLocale xml:lang=\"en-US\"/>");
            sb.Append("<wsman:OperationTimeout>PT60.000S</wsman:OperationTimeout>");

            if (shellId != null)
            {
                sb.Append("<wsman:SelectorSet>");
                sb.Append($"<wsman:Selector Name=\"ShellId\">{shellId}</wsman:Selector>");
                sb.Append("</wsman:SelectorSet>");
            }

            sb.Append("</s:Header>");
            return sb.ToString();
        }

        public static string BuildCreateShell(string endpoint, string msgId)
        {
            string header = BuildHeader(endpoint, ActionCreate, msgId, null);
            string body =
                "<s:Body>" +
                "<rsp:Shell>" +
                "<rsp:InputStreams>stdin</rsp:InputStreams>" +
                "<rsp:OutputStreams>stdout stderr</rsp:OutputStreams>" +
                "</rsp:Shell>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        public static string BuildSend(string endpoint, string msgId, string shellId, string commandId, string base64Data)
        {
            string header = BuildHeader(endpoint, ActionSend, msgId, shellId);
            string body =
                "<s:Body>" +
                "<rsp:Send>" +
                $"<rsp:Stream Name=\"stdin\" CommandId=\"{commandId}\">{base64Data}</rsp:Stream>" +
                "</rsp:Send>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        public static string BuildReceive(string endpoint, string msgId, string shellId, string commandId)
        {
            string header = BuildHeader(endpoint, ActionReceive, msgId, shellId);
            string body =
                "<s:Body>" +
                "<rsp:Receive>" +
                $"<rsp:DesiredStream CommandId=\"{commandId}\">stdout stderr</rsp:DesiredStream>" +
                "</rsp:Receive>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        public static string BuildDelete(string endpoint, string msgId, string shellId)
        {
            string header = BuildHeader(endpoint, ActionDelete, msgId, shellId);
            return SoapPrologue + header + "<s:Body/></s:Envelope>";
        }

        /// <summary>
        /// Parses a CreateShell response and returns the ShellId, or null if not found.
        /// </summary>
        public static string? ParseShellId(string responseXml)
        {
            try
            {
                var doc = XDocument.Parse(responseXml);
                XNamespace rsp = NsRsp;
                XNamespace wsman = NsWsMan;

                // Try body: <rsp:Shell><rsp:ShellId>
                var shellIdEl = doc.Descendants(rsp + "ShellId").GetEnumerator();
                if (shellIdEl.MoveNext())
                    return shellIdEl.Current.Value;

                // Try header SelectorSet: <wsman:Selector Name="ShellId">
                foreach (var sel in doc.Descendants(wsman + "Selector"))
                {
                    if ((string?)sel.Attribute("Name") == "ShellId")
                        return sel.Value;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Parses a ReceiveResponse and returns streams keyed by name (e.g. "stdout", "stderr").
        /// </summary>
        public static IEnumerable<(string Name, string Base64Data)> ParseStreams(string responseXml)
        {
            XNamespace rsp = NsRsp;
            XDocument doc;
            try { doc = XDocument.Parse(responseXml); }
            catch { yield break; }

            foreach (var stream in doc.Descendants(rsp + "Stream"))
            {
                string? name = (string?)stream.Attribute("Name");
                string? data = stream.Value;
                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(data))
                    yield return (name, data);
            }
        }

        /// <summary>
        /// Returns true if the SOAP response contains a Fault element.
        /// </summary>
        public static bool IsFault(string responseXml)
        {
            try
            {
                var doc = XDocument.Parse(responseXml);
                XNamespace soap = NsSoap;
                return doc.Descendants(soap + "Fault").GetEnumerator().MoveNext();
            }
            catch { return false; }
        }
    }

    /// <summary>
    /// Custom TextWriter that sends PSRP data via WSMan Send requests.
    /// The PSRP base class calls WriteLine() to transmit each protocol message.
    /// </summary>
    internal sealed class WinRMTextWriter : TextWriter
    {
        private readonly WinRMClientSessionTransportMgr _transport;

        public WinRMTextWriter(WinRMClientSessionTransportMgr transport)
        {
            _transport = transport ?? throw new ArgumentNullException(nameof(transport));
        }

        public override Encoding Encoding => Encoding.UTF8;

        public override void WriteLine(string? value)
        {
            if (value != null)
                _transport.SendPsrpLine(value);
        }

        public override void Write(string? value)
        {
            if (value != null)
                _transport.SendPsrpLine(value);
        }

        public override void Write(char value)
        {
            _transport.SendPsrpLine(value.ToString());
        }

        public override void Flush() { }
    }

    /// <summary>
    /// WinRM transport manager using HttpClient and WSMan SOAP protocol.
    /// Implements custom PSRemoting transport that communicates via HTTP POST requests.
    /// </summary>
    internal sealed class WinRMClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly WinRMClientInfo _connectionInfo;
        private HttpClient? _httpClient;
        private string? _shellId;
        private string? _commandId;
        private CancellationTokenSource? _readerCts;
        private volatile bool _isClosed;

        private const string ReaderThreadName = "WinRM Transport Reader Thread";

        internal WinRMClientSessionTransportMgr(
            WinRMClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            _connectionInfo = connectionInfo ?? throw new PSArgumentException(nameof(connectionInfo));
        }

        public override void CreateAsync()
        {
            Task.Run(async () =>
            {
                try
                {
                    await CreateConnectionAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    HandleTransportError(new PSRemotingTransportException(
                        $"Failed to establish WinRM connection to '{_connectionInfo.ComputerName}': {ex.Message}", ex));
                }
            }).Wait();
        }

        private async Task CreateConnectionAsync()
        {
            var handler = new HttpClientHandler
            {
                AllowAutoRedirect = _connectionInfo.AllowRedirection
            };

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(120)
            };

            // Add Basic auth header if credentials are provided
            if (_connectionInfo.Authentication == WinRMAuthenticationMechanism.Basic ||
                _connectionInfo.Authentication == WinRMAuthenticationMechanism.Default)
            {
                if (_connectionInfo.Credential != null)
                {
                    var netCred = _connectionInfo.Credential.GetNetworkCredential();
                    string encoded = Convert.ToBase64String(
                        Encoding.ASCII.GetBytes($"{netCred.UserName}:{netCred.Password}"));
                    _httpClient.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Basic", encoded);
                }
            }

            // Create the remote shell and get ShellId
            _shellId = await CreateShellAsync().ConfigureAwait(false);
            _commandId = _shellId; // Use ShellId as CommandId for PSRP

            // Set up PSRP writer - PSRP base class calls WriteLine() on this
            SetMessageWriter(new WinRMTextWriter(this));

            // Start the reader thread that polls Receive
            _readerCts = new CancellationTokenSource();
            StartReaderThread();

            // Trigger the initial PSRP protocol fragment
            SendOneItem();
        }

        private async Task<string> CreateShellAsync()
        {
            string endpoint = _connectionInfo.BuildEndpointUri();
            string msgId = Guid.NewGuid().ToString().ToUpper();
            string soapXml = WsManXml.BuildCreateShell(endpoint, msgId);

            string response = await PostSoapAsync(endpoint, soapXml, WsManXml.ActionCreate)
                .ConfigureAwait(false);

            string? shellId = WsManXml.ParseShellId(response);
            if (string.IsNullOrEmpty(shellId))
                throw new InvalidOperationException(
                    "WinRM server did not return a ShellId in the CreateShell response");

            return shellId;
        }

        /// <summary>
        /// Called by WinRMTextWriter to send a PSRP protocol line via WSMan Send.
        /// The line is UTF-8 encoded then Base64-encoded into the SOAP stream element.
        /// </summary>
        internal void SendPsrpLine(string line)
        {
            if (_isClosed || _shellId == null || _httpClient == null) return;

            try
            {
                // Encode PSRP line: append \n so pwsh stdin receives complete lines
                byte[] lineBytes = Encoding.UTF8.GetBytes(line + "\n");
                string base64Data = Convert.ToBase64String(lineBytes);

                string endpoint = _connectionInfo.BuildEndpointUri();
                string msgId = Guid.NewGuid().ToString().ToUpper();
                string soapXml = WsManXml.BuildSend(endpoint, msgId, _shellId, _commandId!, base64Data);

                PostSoapAsync(endpoint, soapXml, WsManXml.ActionSend).GetAwaiter().GetResult();
            }
            catch (Exception ex) when (!_isClosed)
            {
                HandleTransportError(new PSRemotingTransportException(
                    $"WinRM Send failed: {ex.Message}", ex));
            }
        }

        private async Task<string> PostSoapAsync(string endpoint, string soapXml, string action)
        {
            if (_httpClient == null)
                throw new ObjectDisposedException(nameof(_httpClient));

            var content = new StringContent(soapXml, Encoding.UTF8, "application/soap+xml");

            var request = new HttpRequestMessage(HttpMethod.Post, endpoint)
            {
                Content = content
            };
            request.Headers.Add("SOAPAction", action);

            var response = await _httpClient.SendAsync(request).ConfigureAwait(false);

            string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new InvalidOperationException(
                    $"WSMan request to '{endpoint}' failed with HTTP {(int)response.StatusCode}: {body}");
            }

            return body;
        }

        private void StartReaderThread()
        {
            var thread = new Thread(ProcessReaderThread)
            {
                Name = ReaderThreadName,
                IsBackground = true
            };
            thread.Start();
        }

        /// <summary>
        /// Reader thread: continuously polls the server with WSMan Receive requests
        /// and feeds returned PSRP lines into the base class via HandleOutputDataReceived.
        /// </summary>
        private void ProcessReaderThread()
        {
            try
            {
                var cancellationToken = _readerCts?.Token ?? CancellationToken.None;

                while (!_isClosed && !cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        string endpoint = _connectionInfo.BuildEndpointUri();
                        string msgId = Guid.NewGuid().ToString().ToUpper();
                        string soapXml = WsManXml.BuildReceive(endpoint, msgId, _shellId!, _commandId!);

                        string response = PostSoapAsync(endpoint, soapXml, WsManXml.ActionReceive)
                            .GetAwaiter().GetResult();

                        if (string.IsNullOrWhiteSpace(response))
                            continue;

                        // Decode each returned stdout stream element as a PSRP line
                        foreach (var (name, base64Data) in WsManXml.ParseStreams(response))
                        {
                            if (name != "stdout") continue;

                            byte[] bytes = Convert.FromBase64String(base64Data);
                            string text = Encoding.UTF8.GetString(bytes);

                            // Each line is a separate PSRP protocol message
                            foreach (string rawLine in text.Split('\n'))
                            {
                                string psrpLine = rawLine.TrimEnd('\r');
                                if (!string.IsNullOrEmpty(psrpLine))
                                    HandleOutputDataReceived(psrpLine);
                            }
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                    catch (AggregateException ae) when (ae.InnerException is OperationCanceledException)
                    {
                        break;
                    }
                    catch (Exception) when (_isClosed)
                    {
                        break;
                    }
                    catch (Exception ex)
                    {
                        if (!_isClosed)
                        {
                            HandleTransportError(new PSRemotingTransportException(
                                $"WinRM Receive failed: {ex.Message}", ex));
                        }
                        break;
                    }
                }
            }
            catch (Exception ex) when (!_isClosed)
            {
                HandleTransportError(new PSRemotingTransportException(
                    $"WinRM reader thread ended unexpectedly: {ex.Message}", ex));
            }
        }

        public override void CloseAsync()
        {
            // DO NOT set _isClosed or cancel the reader here.
            // The reader must stay active to receive the PSRP CloseAck from the server.
            // base.CloseAsync() sends the PSRP close packet via WinRMTextWriter;
            // CleanupConnection() is called by the base class once CloseAck arrives.
            base.CloseAsync();
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
                CleanupConnection();

            base.Dispose(isDisposing);
        }

        protected override void CleanupConnection()
        {
            _isClosed = true;
            try { _readerCts?.Cancel(); } catch { }

            // Send WSMan Delete to cleanly close the remote shell before disposing HttpClient
            if (_shellId != null && _httpClient != null)
            {
                try
                {
                    string endpoint = _connectionInfo.BuildEndpointUri();
                    string msgId = Guid.NewGuid().ToString().ToUpper();
                    string soapXml = WsManXml.BuildDelete(endpoint, msgId, _shellId);
                    PostSoapAsync(endpoint, soapXml, WsManXml.ActionDelete).Wait(5000);
                }
                catch { }
            }

            try { _httpClient?.Dispose(); } catch { }
            _readerCts?.Dispose();
            _readerCts = null;
            _httpClient = null;
        }

        private void HandleTransportError(PSRemotingTransportException ex)
        {
            RaiseErrorHandler(new TransportErrorOccuredEventArgs(
                ex, TransportMethodEnum.CloseShellOperationEx));
            CleanupConnection();
        }
    }
}
