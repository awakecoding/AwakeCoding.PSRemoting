using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
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

    internal static class WinRMTrace
    {
        public const string EnvironmentVariableName = "AWAKECODING_PSREMOTING_WINRM_TRACE";

        private static readonly bool s_isEnabled = IsEnabled();

        public static void WriteLine(string message)
        {
            if (s_isEnabled)
            {
                Console.Error.WriteLine(message);
            }
        }

        private static bool IsEnabled()
        {
            string? value = Environment.GetEnvironmentVariable(EnvironmentVariableName);
            return !string.IsNullOrWhiteSpace(value)
                && !string.Equals(value, "0", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(value, "no", StringComparison.OrdinalIgnoreCase);
        }
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

        /// <summary>PowerShell session configuration name (e.g. "Microsoft.PowerShell", "PowerShell.7").</summary>
        public string ConfigurationName { get; set; }

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
            ConfigurationName = WsManXml.DefaultConfigurationName;
            Authentication = WinRMAuthenticationMechanism.Negotiate;
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
                ConfigurationName = ConfigurationName,
                AllowRedirection = AllowRedirection,
                Credential = Credential,
                OpenTimeout = OpenTimeout,
                CancelTimeout = CancelTimeout,
                OperationTimeout = OperationTimeout,
                IdleTimeout = IdleTimeout
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

        public const string ResourceUriBase = "http://schemas.microsoft.com/powershell/";
        public const string DefaultConfigurationName = "Microsoft.PowerShell";

        public static string GetResourceUri(string configurationName) =>
            ResourceUriBase + (string.IsNullOrEmpty(configurationName) ? DefaultConfigurationName : configurationName);

        public const string ActionCreate  = "http://schemas.xmlsoap.org/ws/2004/09/transfer/Create";
        public const string ActionDelete  = "http://schemas.xmlsoap.org/ws/2004/09/transfer/Delete";
        public const string ActionSend    = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Send";
        public const string ActionReceive = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Receive";
        public const string ActionCommand = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Command";
        public const string ActionSignal  = "http://schemas.microsoft.com/wbem/wsman/1/windows/shell/Signal";

        private const string SoapPrologue =
            "<s:Envelope xmlns:s=\"http://www.w3.org/2003/05/soap-envelope\"" +
            " xmlns:wsa=\"http://schemas.xmlsoap.org/ws/2004/08/addressing\"" +
            " xmlns:wsman=\"http://schemas.dmtf.org/wbem/wsman/1/wsman.xsd\"" +
            " xmlns:rsp=\"http://schemas.microsoft.com/wbem/wsman/1/windows/shell\"" +
            " xmlns:pwsh=\"http://schemas.microsoft.com/powershell\">";

        private static string BuildHeader(string to, string action, string msgId, string? shellId, string? configurationName = null, string? extraHeaders = null, string operationTimeout = "PT60.000S")
        {
            var sb = new StringBuilder();
            string resourceUri = GetResourceUri(configurationName ?? DefaultConfigurationName);
            sb.Append("<s:Header>");
            sb.Append($"<wsa:To>{to}</wsa:To>");
            sb.Append($"<wsman:ResourceURI>{resourceUri}</wsman:ResourceURI>");
            sb.Append("<wsa:ReplyTo><wsa:Address>http://schemas.xmlsoap.org/ws/2004/08/addressing/role/anonymous</wsa:Address></wsa:ReplyTo>");
            sb.Append($"<wsa:Action>{action}</wsa:Action>");
            sb.Append($"<wsa:MessageID>uuid:{msgId}</wsa:MessageID>");
            sb.Append("<wsman:MaxEnvelopeSize>512000</wsman:MaxEnvelopeSize>");
            sb.Append("<wsman:Locale xml:lang=\"en-US\"/>");
            sb.Append("<wsman:DataLocale xml:lang=\"en-US\"/>");
            sb.Append($"<wsman:OperationTimeout>{operationTimeout}</wsman:OperationTimeout>");

            if (shellId != null)
            {
                sb.Append("<wsman:SelectorSet>");
                sb.Append($"<wsman:Selector Name=\"ShellId\">{shellId}</wsman:Selector>");
                sb.Append("</wsman:SelectorSet>");
            }

            if (extraHeaders != null)
                sb.Append(extraHeaders);

            sb.Append("</s:Header>");
            return sb.ToString();
        }

        public static string BuildCreateShell(string endpoint, string msgId, string? configurationName = null, string? initialPsrpData = null)
        {
            // protocolversion=2.3 is required by the Windows PS remoting plugin (pwrshplugin.dll).
            // Without it in the OptionSet, wsmprovhost.exe throws a NullReferenceException during CreateShell.
            const string optionSet =
                "<wsman:OptionSet xmlns:xsi=\"http://www.w3.org/1999/XMLSchema-instance\">" +
                "<wsman:Option Name=\"protocolversion\">2.3</wsman:Option>" +
                "</wsman:OptionSet>";
            string header = BuildHeader(endpoint, ActionCreate, msgId, null, configurationName, optionSet);
            string body =
                "<s:Body>" +
                "<rsp:Shell>" +
                "<rsp:InputStreams>stdin pr</rsp:InputStreams>" +
                "<rsp:OutputStreams>stdout</rsp:OutputStreams>" +
                (string.IsNullOrEmpty(initialPsrpData) ? string.Empty : $"<pwsh:creationXml>{initialPsrpData}</pwsh:creationXml>") +
                "</rsp:Shell>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        public static string BuildSend(string endpoint, string msgId, string shellId, string? commandId, string base64Data, string? configurationName = null, string inputStreamName = "stdin")
        {
            string header = BuildHeader(endpoint, ActionSend, msgId, shellId, configurationName);
            string streamAttr = commandId != null ? $" CommandId=\"{commandId}\"" : "";
            string body =
                "<s:Body>" +
                "<rsp:Send>" +
                $"<rsp:Stream Name=\"{inputStreamName}\"{streamAttr}>{base64Data}</rsp:Stream>" +
                "</rsp:Send>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        public static string BuildReceive(string endpoint, string msgId, string shellId, string? commandId, string? configurationName = null)
        {
            // WSMAN_CMDSHELL_OPTION_KEEPALIVE=TRUE tells the server to hold the long-poll
            // connection open and send data as it becomes available (required for PSRP).
            const string optionSet =
                "<wsman:OptionSet xmlns:xsi=\"http://www.w3.org/1999/XMLSchema-instance\">" +
                "<wsman:Option Name=\"WSMAN_CMDSHELL_OPTION_KEEPALIVE\">TRUE</wsman:Option>" +
                "</wsman:OptionSet>";
            string header = BuildHeader(endpoint, ActionReceive, msgId, shellId, configurationName, optionSet, "PT1.000S");
            string streamAttr = commandId != null ? $" CommandId=\"{commandId}\"" : "";
            string body =
                "<s:Body>" +
                "<rsp:Receive>" +
                $"<rsp:DesiredStream{streamAttr}>stdout</rsp:DesiredStream>" +
                "</rsp:Receive>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        public static string BuildDelete(string endpoint, string msgId, string shellId, string? configurationName = null)
        {
            string header = BuildHeader(endpoint, ActionDelete, msgId, shellId, configurationName);
            return SoapPrologue + header + "<s:Body/></s:Envelope>";
        }

        public static string BuildRunCommand(string endpoint, string msgId, string shellId, string pipelineGuid, string? base64PsrpData = null, string? configurationName = null)
        {
            // Native PS7 client (WSManRunShellCommandEx) sends:
            //   commandId   = PowershellInstanceId (pipeline GUID) → maps to CommandId attribute on CommandLine
            //   commandLine = " " (single space)
            //   argSet      = serialized PSRP data (the first message, e.g. PublicKey) → maps to <rsp:Arguments>
            //   options     = IntPtr.Zero (no OptionSet)
            // pwrshplugin.dll requires the PSRP data in <rsp:Arguments> to create the pipeline context;
            // without it, it returns HTTP 500 "fatal error while processing WSManPluginCommand arguments".
            string header = BuildHeader(endpoint, ActionCommand, msgId, shellId, configurationName);
            string args = base64PsrpData != null ? $"<rsp:Arguments>{base64PsrpData}</rsp:Arguments>" : string.Empty;
            string body =
                "<s:Body>" +
                $"<rsp:CommandLine CommandId=\"{pipelineGuid}\">" +
                "<rsp:Command> </rsp:Command>" +
                args +
                "</rsp:CommandLine>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        public static string BuildSignal(string endpoint, string msgId, string shellId, string commandId, string? configurationName = null)
        {
            string header = BuildHeader(endpoint, ActionSignal, msgId, shellId, configurationName);
            string body =
                "<s:Body>" +
                $"<rsp:Signal CommandId=\"{commandId}\">" +
                "<rsp:Code>http://schemas.microsoft.com/wbem/wsman/1/windows/shell/signal/terminate</rsp:Code>" +
                "</rsp:Signal>" +
                "</s:Body>";
            return SoapPrologue + header + body + "</s:Envelope>";
        }

        /// <summary>
        /// Parses a RunCommand response and returns the CommandId, or null if not found.
        /// </summary>
        public static string? ParseCommandId(string responseXml)
        {
            try
            {
                var doc = XDocument.Parse(responseXml);
                XNamespace rsp = NsRsp;
                var el = doc.Descendants(rsp + "CommandId").GetEnumerator();
                if (el.MoveNext())
                    return el.Current.Value;
                return null;
            }
            catch { return null; }
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
        /// Extracts the base64 PSRP fragment content from a PSRP stdio Data element.
        /// Input format: <Data Stream='Default' PSGuid='GUID'>BASE64</Data>
        /// Returns the BASE64 string for use in WinRM SOAP stream elements.
        /// </summary>
        public static string ExtractBase64FromDataElement(string dataLine)
        {
            int gtPos = dataLine.IndexOf('>');
            int ltPos = dataLine.LastIndexOf('<');
            if (gtPos >= 0 && ltPos > gtPos)
                return dataLine.Substring(gtPos + 1, ltPos - gtPos - 1);
            return dataLine; // fallback: return as-is
        }

        /// <summary>
        /// Extracts the PSGuid attribute value from a PSRP stdio Data element.
        /// Input format: &lt;Data Stream='Default' PSGuid='GUID'&gt;BASE64&lt;/Data&gt;
        /// Returns the GUID string (without braces) or null if not found.
        /// </summary>
        public static string? ExtractPsGuidFromDataElement(string dataLine)
        {
            int idx = dataLine.IndexOf("PSGuid='", StringComparison.OrdinalIgnoreCase);
            char quote = '\'';
            if (idx < 0) { idx = dataLine.IndexOf("PSGuid=\"", StringComparison.OrdinalIgnoreCase); quote = '"'; }
            if (idx < 0) return null;
            int start = idx + 8; // skip PSGuid=' (8 chars)
            int end = dataLine.IndexOf(quote, start);
            if (end < 0) return null;
            return dataLine.Substring(start, end - start);
        }

        public static string? ExtractStreamNameFromDataElement(string dataLine)
        {
            int idx = dataLine.IndexOf("Stream='", StringComparison.OrdinalIgnoreCase);
            char quote = '\'';
            if (idx < 0) { idx = dataLine.IndexOf("Stream=\"", StringComparison.OrdinalIgnoreCase); quote = '"'; }
            if (idx < 0) return null;
            int start = idx + 8; // skip Stream=' (8 chars)
            int end = dataLine.IndexOf(quote, start);
            if (end < 0) return null;
            return dataLine.Substring(start, end - start);
        }

        /// <summary>
        /// Wraps a base64 PSRP fragment in a stdio Data element for use with HandleOutputDataReceived.
        /// Output format: <Data Stream="Default" PSGuid="GUID">BASE64</Data>
        /// Both Stream and PSGuid attributes are required by the PSRP base class parser.
        /// The PSGuid is extracted from the binary PSRP fragment to enable correct pipeline routing.
        /// </summary>
        public static string WrapInDataElement(string base64Data, Dictionary<ulong, string> fragmentGuids)
        {
            string psGuid = ExtractPsGuidFromFragment(base64Data, fragmentGuids);
            return $"<Data Stream='Default' PSGuid='{psGuid}'>{base64Data}</Data>";
        }

        /// <summary>
        /// Extracts the PSGuid from a binary PSRP fragment for reconstructing stdio Data elements.
        /// Tracks multi-fragment object IDs in <paramref name="fragmentGuids"/>.
        /// For START fragments: reads RunspacePool GUID and Pipeline GUID from the message header.
        /// For continuation fragments: looks up the stored GUID by ObjectId.
        /// </summary>
        public static string ExtractPsGuidFromFragment(string base64Data, Dictionary<ulong, string>? fragmentGuids = null)
        {
            const string zero = "00000000-0000-0000-0000-000000000000";
            try
            {
                byte[] bytes = Convert.FromBase64String(base64Data);
                // Fragment header: ObjectId(8) + FragmentId(8) + Flags(1) + BlobLength(4) = 21 bytes
                if (bytes.Length < 21) return zero;

                ulong objectId = BitConverter.ToUInt64(bytes, 0);
                byte flags = bytes[16];
                bool isStart = (flags & 0x02) != 0;
                bool isEnd   = (flags & 0x01) != 0;

                if (isStart)
                {
                    // Message header after 21-byte fragment header:
                    // Destination(4) + MessageType(4) + RunspacePoolId(16) + PipelineId(16)
                    if (bytes.Length < 61) return zero;
                    var rpid = new Guid(bytes[29..45]);
                    var pid  = new Guid(bytes[45..61]);
                    // Session-level messages use PSGuid=zeros so they route to _sessionMessageQueue.
                    // Command-level messages use PipelineId so they route to _commandMessageQueue.
                    // (NOT the actual RPID — the session queue consumer handles all non-pipeline PSRP messages.)
                    string psGuid = pid == Guid.Empty ? zero : pid.ToString();

                    if (fragmentGuids != null && !isEnd)
                        fragmentGuids[objectId] = psGuid; // store for continuation fragments
                    else if (fragmentGuids != null && isEnd)
                        fragmentGuids.Remove(objectId);

                    return psGuid;
                }
                else if (fragmentGuids != null && fragmentGuids.TryGetValue(objectId, out string? stored))
                {
                    if (isEnd) fragmentGuids.Remove(objectId);
                    return stored;
                }
                return zero;
            }
            catch { return zero; }
        }


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

        /// <summary>
        /// Returns true when the SOAP fault is a <c>w:TimedOut</c> — the WinRM long-poll
        /// "no data available" signal. This is normal; the caller should simply re-poll.
        /// </summary>
        public static bool IsTimeoutFault(string responseXml)
        {
            try
            {
                var doc = XDocument.Parse(responseXml);
                return doc.Descendants()
                    .Any(e => e.Name.LocalName == "Value"
                           && e.Parent?.Name.LocalName == "Subcode"
                           && e.Value.EndsWith("TimedOut", StringComparison.OrdinalIgnoreCase));
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
        private readonly Guid _runspaceId;
        private readonly PSRemotingCryptoHelper? _cryptoHelper; // stored from constructor for key exchange bypass
        private readonly System.Diagnostics.Stopwatch _sw = System.Diagnostics.Stopwatch.StartNew();
        private string T => $"+{_sw.ElapsedMilliseconds,6}ms";
        private HttpClient? _httpClient;
        private string? _shellId;
        private volatile string? _commandId;          // WinRM CommandId for current active pipeline
        private volatile string? _commandPipelineGuid; // PSRP PipelineId matching _commandId
        private volatile string? _pendingRunCommandGuid; // pipeline GUID awaiting first <Data> to trigger RunCommand
        private volatile int _runCommandCount;         // number of RunCommand calls made (first = key exchange)
        private string? _keyExchangePublicKeyB64;      // saved PUBLIC_KEY fragment for RSA encryption of injected key
        private volatile bool _publicKeyRequestInjected;
        private volatile bool _keyExchangeInjected;    // true once we have injected the synthetic EncryptedSessionKey
        private volatile bool _pendingEncryptedSessionKeyInjection;
        private volatile bool _encryptedSessionKeyInjectionScheduled;
        private bool _sessionStateLoggingHooked;
        private CancellationTokenSource? _readerCts;
        private volatile bool _isClosed;
        // Tracks multi-fragment PSRP object GUIDs for PSGuid reconstruction on receive
        private readonly Dictionary<ulong, string> _receiveFragmentGuids = new();

        private const string ReaderThreadName = "WinRM Transport Reader Thread";

        internal WinRMClientSessionTransportMgr(
            WinRMClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            _connectionInfo = connectionInfo ?? throw new PSArgumentException(nameof(connectionInfo));
            _runspaceId = runspaceId;
            _cryptoHelper = cryptoHelper;
        }

        public override void CreateAsync()
        {
            // Do NOT block here — blocking the PS pipeline thread prevents state change events
            // from being processed, causing WaitOne in CreateAndOpenRunspace to never unblock.
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
            });
        }

        private HttpMessageHandler CreateHttpHandler()
        {
            bool useNegotiate = _connectionInfo.Authentication == WinRMAuthenticationMechanism.Negotiate
                             || _connectionInfo.Authentication == WinRMAuthenticationMechanism.Default
                             || _connectionInfo.Authentication == WinRMAuthenticationMechanism.Kerberos;

            // On Windows, use WinHttpHandler for Negotiate/NTLM/Kerberos — it correctly handles
            // the multi-round-trip SSPI handshake even when the server sends Connection: close.
            if (useNegotiate && RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                var winHandler = new WinHttpHandler
                {
                    AutomaticDecompression = System.Net.DecompressionMethods.None,
                    AutomaticRedirection = _connectionInfo.AllowRedirection,
                    SendTimeout = TimeSpan.FromSeconds(120),
                    ReceiveDataTimeout = TimeSpan.FromSeconds(120),
                    ReceiveHeadersTimeout = TimeSpan.FromSeconds(120)
                };

                if (_connectionInfo.Credential != null)
                {
                    var netCred = _connectionInfo.Credential.GetNetworkCredential();
                    string authType = _connectionInfo.Authentication == WinRMAuthenticationMechanism.Kerberos
                        ? "Kerberos" : "Negotiate";
                    string endpoint = _connectionInfo.BuildEndpointUri();
                    var uri = new Uri(endpoint);
                    var cache = new CredentialCache();
                    cache.Add(new Uri($"{uri.Scheme}://{uri.Host}:{uri.Port}"), authType, netCred);
                    winHandler.ServerCredentials = cache;
                }

                return winHandler;
            }

            // Fallback: HttpClientHandler (Basic auth or non-Windows Negotiate via GSS-API)
            var handler = new HttpClientHandler
            {
                AllowAutoRedirect = _connectionInfo.AllowRedirection,
                PreAuthenticate = false
            };

            if (_connectionInfo.Credential != null)
            {
                var netCred = _connectionInfo.Credential.GetNetworkCredential();
                string endpoint = _connectionInfo.BuildEndpointUri();
                var uri = new Uri(endpoint);
                var cache = new CredentialCache();

                if (useNegotiate)
                {
                    cache.Add(new Uri($"{uri.Scheme}://{uri.Host}:{uri.Port}"), "Negotiate", netCred);
                }
                else
                {
                    cache.Add(new Uri($"{uri.Scheme}://{uri.Host}:{uri.Port}"), "Basic", netCred);
                }

                handler.Credentials = cache;
            }

            return handler;
        }

        private async Task CreateConnectionAsync()
        {
            HttpMessageHandler handler = CreateHttpHandler();

            _httpClient = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromSeconds(120)
            };

            byte[]? initialSessionData = TryReadInitialSessionData();

            // Create the remote shell and get ShellId
            _shellId = await CreateShellAsync(initialSessionData).ConfigureAwait(false);
            WinRMTrace.WriteLine($"[WINRM-DBG {T}] CreateShell OK, ShellId={_shellId}");
            WinRMTrace.WriteLine($"[WINRM-DBG] RunspaceInstanceId={_runspaceId}");
            TryHookSessionStateLogging();

            // PSRP over WinRM does NOT use RunCommand. Unlike WinRS (which runs cmd.exe commands),
            // PSRP sends protocol messages directly to the shell via Send/Receive without a command context.
            // The ShellId selector (set on all requests) is sufficient to identify the PS session.

            // Set up PSRP writer - PSRP base class calls WriteLine() on this
            SetMessageWriter(new WinRMTextWriter(this));
            RaiseCreateCompletedCompat();
            QueueDataAck(null);

            // Start the reader thread that polls Receive
            _readerCts = new CancellationTokenSource();
            StartReaderThread();
        }

        private async Task<string> CreateShellAsync(byte[]? initialPsrpData = null)
        {
            string endpoint = _connectionInfo.BuildEndpointUri();
            string msgId = Guid.NewGuid().ToString().ToUpper();
            string? initialPsrpDataB64 = initialPsrpData != null && initialPsrpData.Length > 0
                ? Convert.ToBase64String(initialPsrpData)
                : null;
            string soapXml = WsManXml.BuildCreateShell(endpoint, msgId, _connectionInfo.ConfigurationName, initialPsrpDataB64);

            string response = await PostSoapAsync(endpoint, soapXml, WsManXml.ActionCreate)
                .ConfigureAwait(false);

            string? shellId = WsManXml.ParseShellId(response);
            if (string.IsNullOrEmpty(shellId))
                throw new InvalidOperationException(
                    "WinRM server did not return a ShellId in the CreateShell response");

            return shellId;
        }

        private async Task<string> RunCommandAsync(string shellId)
        {
            string endpoint = _connectionInfo.BuildEndpointUri();
            string msgId = Guid.NewGuid().ToString().ToUpper();
            string soapXml = WsManXml.BuildRunCommand(endpoint, msgId, shellId, _connectionInfo.ConfigurationName);

            string response = await PostSoapAsync(endpoint, soapXml, WsManXml.ActionCommand)
                .ConfigureAwait(false);

            // Real WinRM returns a distinct CommandId; our custom server may return ShellId
            string? commandId = WsManXml.ParseCommandId(response) ?? WsManXml.ParseShellId(response);
            if (string.IsNullOrEmpty(commandId))
                throw new InvalidOperationException(
                    "WinRM server did not return a CommandId in the RunCommand response");

            return commandId;
        }

        /// <summary>
        /// Called by WinRMTextWriter to send a PSRP protocol line via WSMan Send.
        /// The line is UTF-8 encoded then Base64-encoded into the SOAP stream element.
        /// </summary>
        internal void SendPsrpLine(string line)
        {
            if (_isClosed || _shellId == null || _httpClient == null) return;

            WinRMTrace.WriteLine($"[WINRM-DBG] SendPsrpLine RAW len={line.Length} first40=[{line.Substring(0, Math.Min(40, line.Length))}]");

            try
            {
                string trimmedLine = line.TrimEnd('\n', '\r').TrimStart();

                // <Command PSGuid='GUID' /> — PSRP pipeline creation signal.
                // Per the PS7 WinRM transport (WSManRunShellCommandEx), the pipeline GUID goes as
                // CommandId on CommandLine, and the FIRST PSRP data fragment goes as <rsp:Arguments>.
                // So we defer the RunCommand until we see the first <Data> for this pipeline GUID.
                if (trimmedLine.StartsWith("<Command ", StringComparison.OrdinalIgnoreCase))
                {
                    string? pipelineGuid = ExtractPsGuidAttribute(trimmedLine);
                    if (pipelineGuid != null && _shellId != null)
                    {
                        // Remember this pipeline GUID: the next <Data PSGuid='GUID'> will trigger RunCommand.
                        _pendingRunCommandGuid = pipelineGuid;

                        // Inject a synthetic <CommandAck> to unblock the PS SDK consumer so it
                        // proceeds to send the first <Data> (PublicKey for key exchange).
                        // Use Task.Run to avoid reentrancy (consumer exits SendPsrpLine first).
                        var ackGuid = pipelineGuid;
                        _ = Task.Run(async () =>
                        {
                            await Task.Delay(50);
                            HandleOutputDataReceived($"<CommandAck PSGuid='{ackGuid}'/>");
                        });
                    }
                    return;
                }

                // <Close PSGuid='GUID' /> — PSRP pipeline close.
                // Signal terminate on the command context, then clear it.
                if (trimmedLine.StartsWith("<Close ", StringComparison.OrdinalIgnoreCase))
                {
                    string closeGuid = ExtractPsGuidAttribute(trimmedLine) ?? Guid.Empty.ToString();
                    string? closingCommandId = _commandId;
                    if (closingCommandId != null && _shellId != null)
                    {
                        _commandId = null;
                        _commandPipelineGuid = null;
                        string endpoint = _connectionInfo.BuildEndpointUri();
                        string sigMsgId = Guid.NewGuid().ToString().ToUpper();
                        string sigSoap = WsManXml.BuildSignal(endpoint, sigMsgId, _shellId, closingCommandId, _connectionInfo.ConfigurationName);
                        try { PostSoapAsync(endpoint, sigSoap, WsManXml.ActionSignal).GetAwaiter().GetResult(); } catch { /* best effort */ }
                    }

                    // Native WSMan transport reports command/session close completion through transport callbacks
                    // rather than out-of-proc text packets. Synthesize the expected CloseAck control packet so the
                    // client transport managers can finish their close bookkeeping.
                    var ackGuid = closeGuid;
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(50).ConfigureAwait(false);
                        HandleOutputDataReceived($"<CloseAck PSGuid='{ackGuid}'/>");
                    });
                    return;
                }

                // Regular <Data Stream='Default' PSGuid='GUID'>BASE64</Data> message.
                // Extract the base64 payload and send via WinRM Send.
                string base64Data = WsManXml.ExtractBase64FromDataElement(trimmedLine);
                if (string.IsNullOrEmpty(base64Data) || base64Data.TrimStart().StartsWith("<"))
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG] SendPsrpLine SKIPPED (not a Data element)");
                    return;
                }

                // Decode the PSRP fragment to detect message types relevant to key exchange.
                // After we inject PublicKeyRequest the client should send an actual PUBLIC_KEY
                // session message, and only then should we inject EncryptedSessionKey.
                try
                {
                    byte[] outFrag = Convert.FromBase64String(base64Data);
                    if (outFrag.Length >= 29)
                    {
                        byte outFlags = outFrag[16];
                        bool outIsStart = (outFlags & 0x02) != 0;
                        if (outIsStart)
                        {
                            uint outMt = BitConverter.ToUInt32(outFrag, 25);
                            WinRMTrace.WriteLine($"[WINRM-DBG] OUTGOING MsgType=0x{outMt:X8} Flags=0x{outFlags:X2} b64Len={base64Data.Length} reqInjected={_publicKeyRequestInjected} keyInjected={_keyExchangeInjected}");
                            if (outMt == 0x00010005)
                            {
                                _keyExchangePublicKeyB64 = base64Data;
                                WinRMTrace.WriteLine($"[WINRM-DBG] PUBLIC_KEY (0x{outMt:X8}) saved for RSA injection");
                                WinRMTrace.WriteLine($"[WINRM-DBG] OUTGOING FULL_B64={base64Data}");
                                WinRMTrace.WriteLine($"[WINRM-DBG] Session state at PUBLIC_KEY send: {GetCurrentSessionState()}");

                                if (_publicKeyRequestInjected && !_keyExchangeInjected)
                                {
                                    _pendingEncryptedSessionKeyInjection = true;
                                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] PublicKey observed after PublicKeyRequest - waiting for command setup before EncryptedSessionKey");
                                }
                            }
                            else if (outMt == 0x00021006 && !_keyExchangeInjected)
                            {
                                if (_publicKeyRequestInjected)
                                {
                                    _pendingEncryptedSessionKeyInjection = true;
                                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] CreatePowerShell observed after PublicKeyRequest - waiting for command setup before EncryptedSessionKey");
                                }
                                else
                                {
                                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] CreatePowerShell observed with no key request pending");
                                }
                            }
                        }
                    }
                }
                catch { }

                // If this is the first <Data> for a pending pipeline, send RunCommand with the data
                // as <rsp:Arguments> (matching what WSManRunShellCommandEx sends natively).
                // pwrshplugin.dll requires the PSRP data in <rsp:Arguments> to initialize the pipeline.
                string? dataPipelineGuid = WsManXml.ExtractPsGuidFromDataElement(trimmedLine);
                if (dataPipelineGuid != null && dataPipelineGuid.Equals(_pendingRunCommandGuid, StringComparison.OrdinalIgnoreCase) && _shellId != null)
                {
                    _pendingRunCommandGuid = null;
                    try
                    {
                        string ep = _connectionInfo.BuildEndpointUri();
                        string cmdMsgId = Guid.NewGuid().ToString().ToUpper();
                        string cmdSoap = WsManXml.BuildRunCommand(ep, cmdMsgId, _shellId, dataPipelineGuid, base64Data, _connectionInfo.ConfigurationName);
                        WinRMTrace.WriteLine($"[WINRM-DBG {T}] RunCommand+Args for pipeline {dataPipelineGuid}");
                        string cmdResponse = PostSoapAsync(ep, cmdSoap, WsManXml.ActionCommand).GetAwaiter().GetResult();
                        string? newCommandId = WsManXml.ParseCommandId(cmdResponse) ?? WsManXml.ParseShellId(cmdResponse);
                        WinRMTrace.WriteLine($"[WINRM-DBG {T}] RunCommand OK → commandId={newCommandId}");
                        _commandId = newCommandId;
                        _commandPipelineGuid = dataPipelineGuid;

                        System.Threading.Interlocked.Increment(ref _runCommandCount);
                        if (_pendingEncryptedSessionKeyInjection)
                        {
                            WinRMTrace.WriteLine($"[WINRM-DBG {T}] RunCommand established for key exchange pipeline");
                        }

                        if (newCommandId != null)
                        {
                            var capturedCommandId = newCommandId;
                            var capturedCts = _readerCts;
                            _ = Task.Run(() => DrainCommandReceive(capturedCommandId, capturedCts?.Token ?? CancellationToken.None));
                        }
                        QueueDataAck(dataPipelineGuid);
                        return;
                    }
                    catch (Exception runCmdEx)
                    {
                        WinRMTrace.WriteLine($"[WINRM-DBG] RunCommand+Args failed: {runCmdEx.Message.Split('\n')[0]}");
                        // Fall through to try a plain shell-level Send as last resort.
                    }
                }

                string endpoint2 = _connectionInfo.BuildEndpointUri();
                string msgId = Guid.NewGuid().ToString().ToUpper();
                string outOfProcStream = WsManXml.ExtractStreamNameFromDataElement(trimmedLine) ?? "Default";
                string wsManInputStream = string.Equals(outOfProcStream, "PromptResponse", StringComparison.OrdinalIgnoreCase) ? "pr" : "stdin";
                bool isSessionLevelMessage = string.IsNullOrEmpty(dataPipelineGuid) ||
                    string.Equals(dataPipelineGuid, Guid.Empty.ToString(), StringComparison.OrdinalIgnoreCase);
                string? targetCommandId = isSessionLevelMessage ? null : _commandId;
                string soapXml = WsManXml.BuildSend(endpoint2, msgId, _shellId!, targetCommandId, base64Data, _connectionInfo.ConfigurationName, wsManInputStream);

                WinRMTrace.WriteLine($"[WINRM-DBG] SEND b64.len={base64Data.Length} commandId={targetCommandId ?? "null"} stream={outOfProcStream}->{wsManInputStream}");
                PostSoapAsync(endpoint2, soapXml, WsManXml.ActionSend).GetAwaiter().GetResult();
                WinRMTrace.WriteLine($"[WINRM-DBG] SEND OK");
                QueueDataAck(dataPipelineGuid);
            }
            catch (Exception ex) when (!_isClosed)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] SEND ERROR: {ex.Message}");
                HandleTransportError(new PSRemotingTransportException(
                    $"WinRM Send failed: {ex.Message}", ex));
            }
        }

        private void TryInjectPendingEncryptedSessionKey(string reason)
        {
            if (!_pendingEncryptedSessionKeyInjection || _keyExchangeInjected || _commandId == null || _encryptedSessionKeyInjectionScheduled)
            {
                return;
            }

            _encryptedSessionKeyInjectionScheduled = true;
            _ = Task.Run(async () =>
            {
                await Task.Delay(50).ConfigureAwait(false);
                _encryptedSessionKeyInjectionScheduled = false;
                if (!_pendingEncryptedSessionKeyInjection || _keyExchangeInjected || _commandId == null)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] {reason} - synthetic EncryptedSessionKey cancelled");
                    return;
                }

                _pendingEncryptedSessionKeyInjection = false;
                _keyExchangeInjected = true;
                WinRMTrace.WriteLine($"[WINRM-DBG {T}] {reason} - injecting EncryptedSessionKey");
                InjectSyntheticEncryptedSessionKey();
            });
        }

        private void ObserveRealEncryptedSessionKey(string source)
        {
            _pendingEncryptedSessionKeyInjection = false;
            _keyExchangeInjected = true;
            WinRMTrace.WriteLine($"[WINRM-DBG {T}] Real EncryptedSessionKey received via {source}; synthetic injection disabled; state={GetCurrentSessionState()}");
        }

        /// <summary>
        /// Extracts the PSGuid attribute value from a PSRP control element.
        /// Handles both single-quote and double-quote attribute delimiters.
        /// </summary>
        private static string? ExtractPsGuidAttribute(string line)
        {
            // Look for PSGuid='...' or PSGuid="..."
            int idx = line.IndexOf("PSGuid='", StringComparison.OrdinalIgnoreCase);
            char quote = '\'';
            if (idx < 0) { idx = line.IndexOf("PSGuid=\"", StringComparison.OrdinalIgnoreCase); quote = '"'; }
            if (idx < 0) return null;
            idx += 8; // skip PSGuid=' or PSGuid="
            int end = line.IndexOf(quote, idx);
            return end > idx ? line.Substring(idx, end - idx) : null;
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
                // w:TimedOut is the WinRM long-poll "no data available" signal — not a real error.
                // The reader thread should simply re-poll on an empty return value.
                if (response.StatusCode == System.Net.HttpStatusCode.InternalServerError
                    && WsManXml.IsTimeoutFault(body))
                    return string.Empty;

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

        private byte[]? TryReadInitialSessionData()
        {
            try
            {
                var queue = FindField(typeof(BaseClientTransportManager), "dataToBeSent")?.GetValue(this);
                if (queue == null)
                {
                    return null;
                }

                var method = queue.GetType().GetMethod(
                    "ReadOrRegisterCallback",
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.NonPublic);

                if (method == null)
                {
                    return null;
                }

                var parameters = method.GetParameters();
                object?[] args = new object?[parameters.Length];
                args[0] = null;
                if (parameters.Length > 1)
                {
                    Type outType = parameters[1].ParameterType.GetElementType() ?? parameters[1].ParameterType;
                    args[1] = Activator.CreateInstance(outType);
                }

                byte[]? data = method.Invoke(queue, args) as byte[];
                if (data != null && data.Length > 0)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG] Initial session payload queued ({data.Length} bytes)");
                }

                return data;
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] TryReadInitialSessionData failed: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        private void ProcessRawDataCompat(byte[] data, string stream)
        {
            try
            {
                typeof(BaseClientTransportManager).GetMethod(
                    "ProcessRawData",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                    binder: null,
                    types: new[] { typeof(byte[]), typeof(string) },
                    modifiers: null)
                    ?.Invoke(this, new object[] { data, stream });
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] ProcessRawDataCompat failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private void RaiseCreateCompletedCompat()
        {
            try
            {
                const System.Reflection.BindingFlags flags =
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.NonPublic |
                    System.Reflection.BindingFlags.Public;

                Type? eventArgsType = typeof(BaseClientTransportManager).Assembly
                    .GetType("System.Management.Automation.Remoting.CreateCompleteEventArgs");
                var ctor = eventArgsType?.GetConstructor(
                    flags,
                    binder: null,
                    types: new[] { typeof(RunspaceConnectionInfo) },
                    modifiers: null);
                object? eventArgs = ctor?.Invoke(new object[] { _connectionInfo.Clone() });
                if (eventArgs == null)
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] RaiseCreateCompletedCompat could not construct CreateCompleteEventArgs");
                    return;
                }

                typeof(BaseClientTransportManager).GetMethod("RaiseCreateCompleted", flags)
                    ?.Invoke(this, new[] { eventArgs });
                WinRMTrace.WriteLine("[WINRM-DBG] Raised CreateCompleted");
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] RaiseCreateCompletedCompat failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private void QueueDataAck(string? psGuid)
        {
            string ackGuid = string.IsNullOrWhiteSpace(psGuid) ? Guid.Empty.ToString() : psGuid;
            _ = Task.Run(async () =>
            {
                await Task.Delay(10).ConfigureAwait(false);
                HandleOutputDataReceived($"<DataAck PSGuid='{ackGuid}'/>");
            });
        }

        private byte[]? BuildEncodedPsrpMessage(string encoderMethodName, params object[] args)
        {
            try
            {
                Type? encoderType = typeof(PSObject).Assembly.GetType("System.Management.Automation.RemotingEncoder");
                var encoderMethod = encoderType?.GetMethod(
                    encoderMethodName,
                    System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic);
                if (encoderMethod == null)
                {
                    return null;
                }

                object? remoteDataObject = encoderMethod.Invoke(null, args);
                if (remoteDataObject == null)
                {
                    return null;
                }

                object? fragmentor = typeof(BaseClientTransportManager).GetProperty(
                    "Fragmentor",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic)
                    ?.GetValue(this);

                if (fragmentor == null)
                {
                    return null;
                }

                using var ms = new MemoryStream();
                remoteDataObject.GetType().GetMethod(
                    "Serialize",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic,
                    binder: null,
                    types: new[] { typeof(Stream), fragmentor.GetType() },
                    modifiers: null)
                    ?.Invoke(remoteDataObject, new object[] { ms, fragmentor });

                return ms.ToArray();
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] BuildEncodedPsrpMessage({encoderMethodName}) failed: {ex.GetType().Name}: {ex.Message}");
                return null;
            }
        }

        private bool TryInjectEncodedPsrpMessage(string encoderMethodName, params object[] encoderArgs)
        {
            try
            {
                byte[]? data = BuildEncodedPsrpMessage(encoderMethodName, encoderArgs);
                if (data == null || data.Length == 0)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG] TryInjectEncodedPsrpMessage({encoderMethodName}) produced no data");
                    return false;
                }

                string wrappedData = WsManXml.WrapInDataElement(Convert.ToBase64String(data), _receiveFragmentGuids);
                WinRMTrace.WriteLine($"[WINRM-DBG] Injecting encoded {encoderMethodName} via HandleOutputDataReceived ({data.Length} bytes)");
                HandleOutputDataReceived(wrappedData);
                return true;
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] TryInjectEncodedPsrpMessage({encoderMethodName}) failed: {ex.GetType().Name}: {ex.Message}");
                return false;
            }
        }

        private bool TryRaiseKeyExchangeMessage(string encoderMethodName, params object[] encoderArgs)
        {
            try
            {
                const System.Reflection.BindingFlags flags =
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.NonPublic |
                    System.Reflection.BindingFlags.Public;
                const System.Reflection.BindingFlags staticFlags =
                    System.Reflection.BindingFlags.Static |
                    System.Reflection.BindingFlags.NonPublic |
                    System.Reflection.BindingFlags.Public;

                var asm = typeof(PSObject).Assembly;
                var encoderType = asm.GetType("System.Management.Automation.RemotingEncoder");
                var encoderMethod = encoderType?.GetMethod(encoderMethodName, staticFlags);
                object? remoteDataObject = encoderMethod?.Invoke(null, encoderArgs);
                if (remoteDataObject == null || _cryptoHelper == null)
                {
                    return false;
                }

                Type remoteDataObjectType = remoteDataObject.GetType();
                object? destination = remoteDataObjectType.GetProperty("Destination", flags)?.GetValue(remoteDataObject);
                object? dataType = remoteDataObjectType.GetProperty("DataType", flags)?.GetValue(remoteDataObject);
                Guid runspacePoolId = (Guid)(remoteDataObjectType.GetProperty("RunspacePoolId", flags)?.GetValue(remoteDataObject) ?? Guid.Empty);
                Guid powerShellId = (Guid)(remoteDataObjectType.GetProperty("PowerShellId", flags)?.GetValue(remoteDataObject) ?? Guid.Empty);
                object? data = remoteDataObjectType.GetProperty("Data", flags)?.GetValue(remoteDataObject);

                Type? genericRemoteDataType = asm.GetType("System.Management.Automation.Remoting.RemoteDataObject`1")?.MakeGenericType(typeof(PSObject));
                var createFrom = genericRemoteDataType?.GetMethods(staticFlags)
                    .FirstOrDefault(m => m.Name == "CreateFrom" && m.GetParameters().Length == 5);
                if (createFrom == null || destination == null || dataType == null)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG] TryRaiseKeyExchangeMessage({encoderMethodName}) missing CreateFrom or metadata");
                    return false;
                }

                object psData = data as PSObject ?? PSObject.AsPSObject(data ?? string.Empty);
                object? typedRemoteData = createFrom.Invoke(null, new[] { destination, dataType, runspacePoolId, powerShellId, psData });
                if (typedRemoteData == null)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG] TryRaiseKeyExchangeMessage({encoderMethodName}) CreateFrom returned null");
                    return false;
                }

                object? session = _cryptoHelper.GetType().GetProperty("Session", flags)?.GetValue(_cryptoHelper)
                    ?? FindField(_cryptoHelper.GetType(), "<Session>k__BackingField")?.GetValue(_cryptoHelper);
                object? dataStructureHandler = session != null
                    ? FindField(session.GetType(), "<BaseSessionDataStructureHandler>k__BackingField")?.GetValue(session)
                    : null;
                var raiseMethod = dataStructureHandler?.GetType().GetMethod("RaiseKeyExchangeMessageReceived", flags);
                if (raiseMethod == null)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG] TryRaiseKeyExchangeMessage({encoderMethodName}) missing RaiseKeyExchangeMessageReceived");
                    return false;
                }
                raiseMethod?.Invoke(dataStructureHandler, new[] { typedRemoteData });
                return true;
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] TryRaiseKeyExchangeMessage({encoderMethodName}) failed: {ex.GetType().Name}: {ex.Message}");
                WinRMTrace.WriteLine($"[WINRM-DBG]   detail: {ex}");
                if (ex is System.Reflection.TargetInvocationException tie && tie.InnerException != null)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG]   inner: {tie.InnerException.GetType().Name}: {tie.InnerException.Message}");
                    WinRMTrace.WriteLine($"[WINRM-DBG]   inner detail: {tie.InnerException}");
                }
                return false;
            }
        }

        private void InjectPublicKeyRequest()
        {
            try
            {
                _publicKeyRequestInjected = true;
                WinRMTrace.WriteLine($"[WINRM-DBG {T}] Injecting synthetic PublicKeyRequest");
                if (TryRaiseKeyExchangeMessage("GeneratePublicKeyRequest", _runspaceId))
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] Synthetic PublicKeyRequest via hook injected OK");
                    return;
                }

                WinRMTrace.WriteLine("[WINRM-DBG] Hook PublicKeyRequest injection failed; falling back to encoded data");
                if (!TryInjectEncodedPsrpMessage("GeneratePublicKeyRequest", _runspaceId))
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] Failed to inject PublicKeyRequest");
                    return;
                }
                WinRMTrace.WriteLine($"[WINRM-DBG {T}] Synthetic PublicKeyRequest via encoded data injected OK");
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] InjectPublicKeyRequest failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// Disables PSRP session-level key exchange in the CryptoHelper via reflection.
        /// PS5.1 WinRM servers do not send EncryptedSessionKey; transport-level NTLM/Kerberos
        /// encryption already protects the channel. With key exchange disabled, the PS SDK's
        /// ImportEncryptedSessionKey becomes a no-op and accepts a synthetic reply.
        /// </summary>
        private void TryDisableKeyExchange()
        {
            try
            {
                if (_cryptoHelper == null)
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] TryDisableKeyExchange: _cryptoHelper is null");
                    return;
                }

                // Try direct field on PSRemotingCryptoHelper
                if (TrySetField(_cryptoHelper, "_isKeyExchangeEnabled", false)) return;

                // In PS 7.x the flag lives on the Session object held by the helper.
                var sessionField = _cryptoHelper.GetType().GetField("<Session>k__BackingField",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                var session = sessionField?.GetValue(_cryptoHelper);
                if (session != null && TrySetField(session, "_isKeyExchangeEnabled", false)) return;

                WinRMTrace.WriteLine("[WINRM-DBG] TryDisableKeyExchange: field not found anywhere in helper/session hierarchy");
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] TryDisableKeyExchange failed: {ex.Message}");
            }
        }

        private static System.Reflection.FieldInfo? FindField(Type? type, string fieldName)
        {
            while (type != null && type != typeof(object))
            {
                var field = type.GetField(fieldName,
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Public);
                if (field != null)
                {
                    return field;
                }
                type = type.BaseType;
            }
            return null;
        }

        private static bool TrySetField(object obj, string fieldName, object value)
        {
            var field = FindField(obj.GetType(), fieldName);
            if (field == null)
            {
                return false;
            }

            field.SetValue(obj, value);
            WinRMTrace.WriteLine($"[WINRM-DBG] Set {field.DeclaringType?.Name ?? obj.GetType().Name}.{fieldName} = {value}");
            return true;
        }

        private string GetCurrentSessionState()
        {
            const System.Reflection.BindingFlags flags =
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Public;

            try
            {
                if (_cryptoHelper == null)
                {
                    return "<no-crypto-helper>";
                }

                object? session = _cryptoHelper.GetType().GetProperty("Session", flags)?.GetValue(_cryptoHelper)
                    ?? FindField(_cryptoHelper.GetType(), "<Session>k__BackingField")?.GetValue(_cryptoHelper);
                object? dataStructureHandler = session != null
                    ? FindField(session.GetType(), "<BaseSessionDataStructureHandler>k__BackingField")?.GetValue(session)
                    : null;
                object? stateMachine = dataStructureHandler != null
                    ? FindField(dataStructureHandler.GetType(), "_stateMachine")?.GetValue(dataStructureHandler)
                    : null;
                object? state = stateMachine?.GetType().GetProperty("State", flags)?.GetValue(stateMachine);
                return state?.ToString() ?? "<unknown>";
            }
            catch (Exception ex)
            {
                return $"<state-error:{ex.GetType().Name}>";
            }
        }

        private void TryHookSessionStateLogging()
        {
            if (_sessionStateLoggingHooked || _cryptoHelper == null)
            {
                return;
            }

            const System.Reflection.BindingFlags flags =
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Public;

            try
            {
                object? session = _cryptoHelper.GetType().GetProperty("Session", flags)?.GetValue(_cryptoHelper)
                    ?? FindField(_cryptoHelper.GetType(), "<Session>k__BackingField")?.GetValue(_cryptoHelper);
                if (session == null)
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] Remote session object not available for StateChanged hook");
                    return;
                }

                var eventInfo = session.GetType().GetEvent("StateChanged", flags);
                if (eventInfo?.EventHandlerType == null)
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] Remote session StateChanged event not found");
                    return;
                }

                var handlerMethod = GetType().GetMethod(nameof(OnRemoteSessionStateChanged), flags);
                if (handlerMethod == null)
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] OnRemoteSessionStateChanged method not found");
                    return;
                }

                Delegate? handler = Delegate.CreateDelegate(eventInfo.EventHandlerType, this, handlerMethod, false);
                if (handler == null)
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] Unable to bind remote session StateChanged handler");
                    return;
                }

                eventInfo.AddEventHandler(session, handler);
                _sessionStateLoggingHooked = true;
                WinRMTrace.WriteLine("[WINRM-DBG] Hooked remote session StateChanged logging");
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] Failed to hook remote session StateChanged: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private void OnRemoteSessionStateChanged(object? sender, EventArgs eventArgs)
        {
            try
            {
                object? sessionStateInfo = eventArgs.GetType().GetProperty("SessionStateInfo")?.GetValue(eventArgs);
                object? state = sessionStateInfo?.GetType().GetProperty("State")?.GetValue(sessionStateInfo);
                object? reason = sessionStateInfo?.GetType().GetProperty("Reason")?.GetValue(sessionStateInfo);
                WinRMTrace.WriteLine($"[WINRM-DBG {T}] REMOTE SESSION STATE: {state} Reason={reason}");
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] REMOTE SESSION STATE logging failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private static void TryRaiseSessionStateEvent(object stateMachine, string eventName)
        {
            const System.Reflection.BindingFlags flags =
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Public;

            var asm = typeof(PSObject).Assembly;
            Type? remoteSessionEventType = asm.GetType("System.Management.Automation.RemoteSessionEvent");
            Type? eventArgsType = asm.GetType("System.Management.Automation.RemoteSessionStateMachineEventArgs");
            if (remoteSessionEventType == null || eventArgsType == null)
            {
                return;
            }

            object? eventValue = Enum.Parse(remoteSessionEventType, eventName);
            object? eventArgs = eventArgsType.GetConstructor(new[] { remoteSessionEventType })?.Invoke(new[] { eventValue });
            stateMachine.GetType().GetMethod("RaiseEvent", flags)?.Invoke(stateMachine, new object?[] { eventArgs, false });
        }

        private void TryStartKeyExchangeLocally()
        {
            if (_cryptoHelper == null)
            {
                return;
            }

            const System.Reflection.BindingFlags flags =
                System.Reflection.BindingFlags.Instance |
                System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Public;

            try
            {
                object? session = _cryptoHelper.GetType().GetProperty("Session", flags)?.GetValue(_cryptoHelper)
                    ?? FindField(_cryptoHelper.GetType(), "<Session>k__BackingField")?.GetValue(_cryptoHelper);
                object? dataStructureHandler = session != null
                    ? FindField(session.GetType(), "<BaseSessionDataStructureHandler>k__BackingField")?.GetValue(session)
                    : null;
                object? stateMachine = dataStructureHandler != null
                    ? FindField(dataStructureHandler.GetType(), "_stateMachine")?.GetValue(dataStructureHandler)
                    : null;
                object? stateBefore = stateMachine?.GetType().GetProperty("State", flags)?.GetValue(stateMachine);
                WinRMTrace.WriteLine($"[WINRM-DBG] Session state before StartKeyExchange: {stateBefore}");

                if (string.Equals(stateBefore?.ToString(), "Established", StringComparison.Ordinal))
                {
                    session?.GetType().GetMethod("StartKeyExchange", flags)?.Invoke(session, null);
                }

                object? stateAfter = stateMachine?.GetType().GetProperty("State", flags)?.GetValue(stateMachine);
                WinRMTrace.WriteLine($"[WINRM-DBG] Session state after StartKeyExchange: {stateAfter}");
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] TryStartKeyExchangeLocally failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        private bool TryCompleteKeyExchangeLocally(string encryptedKeyB64)
        {
            if (_cryptoHelper == null || string.IsNullOrEmpty(encryptedKeyB64))
            {
                return false;
            }

            try
            {
                const System.Reflection.BindingFlags flags =
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.NonPublic |
                    System.Reflection.BindingFlags.Public;

                Type helperType = _cryptoHelper.GetType();
                helperType.GetMethod("RunKeyExchangeIfRequired", flags)?.Invoke(_cryptoHelper, null);

                bool imported = helperType.GetMethod("ImportEncryptedSessionKey", flags)?.Invoke(_cryptoHelper, new object[] { encryptedKeyB64 }) as bool? ?? false;
                WinRMTrace.WriteLine($"[WINRM-DBG] ImportEncryptedSessionKey returned {imported}");
                if (!imported)
                {
                    return false;
                }

                helperType.GetMethod("CompleteKeyExchange", flags)?.Invoke(_cryptoHelper, null);

                if (FindField(helperType, "_keyExchangeCompleted")?.GetValue(_cryptoHelper) is EventWaitHandle keyExchangeCompleted)
                {
                    keyExchangeCompleted.Set();
                    WinRMTrace.WriteLine("[WINRM-DBG] Signaled helper key exchange event");
                }

                object? session = helperType.GetProperty("Session", flags)?.GetValue(_cryptoHelper)
                    ?? FindField(helperType, "<Session>k__BackingField")?.GetValue(_cryptoHelper);

                if (session != null)
                {
                    Type sessionType = session.GetType();
                    sessionType.GetMethod("CompleteKeyExchange", flags)?.Invoke(session, null);

                    object? dataStructureHandler = FindField(sessionType, "<BaseSessionDataStructureHandler>k__BackingField")?.GetValue(session);
                    object? stateMachine = dataStructureHandler != null
                        ? FindField(dataStructureHandler.GetType(), "_stateMachine")?.GetValue(dataStructureHandler)
                        : null;

                    if (stateMachine != null)
                    {
                        Type stateMachineType = stateMachine.GetType();
                        object? stateBefore = stateMachineType.GetProperty("State", flags)?.GetValue(stateMachine);
                        WinRMTrace.WriteLine($"[WINRM-DBG] Session state before local completion event: {stateBefore}");

                        if (string.Equals(stateBefore?.ToString(), "Established", StringComparison.Ordinal))
                        {
                            TryRaiseSessionStateEvent(stateMachine, "KeyRequested");
                            TryRaiseSessionStateEvent(stateMachine, "KeySent");
                        }

                        TrySetField(stateMachine, "_keyExchanged", true);
                        TryRaiseSessionStateEvent(stateMachine, "KeyReceived");

                        object? stateAfter = stateMachineType.GetProperty("State", flags)?.GetValue(stateMachine);
                        WinRMTrace.WriteLine($"[WINRM-DBG] Session state after local completion event: {stateAfter}");
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] TryCompleteKeyExchangeLocally failed: {ex.GetType().Name}: {ex.Message}");
                WinRMTrace.WriteLine($"[WINRM-DBG]   detail: {ex}");
                if (ex is System.Reflection.TargetInvocationException tie && tie.InnerException != null)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG]   inner: {tie.InnerException.GetType().Name}: {tie.InnerException.Message}");
                    WinRMTrace.WriteLine($"[WINRM-DBG]   inner detail: {tie.InnerException}");
                }
                return false;
            }
        }

        /// <summary>
        /// Injects a synthetic ENCRYPTED_SESSION_KEY (MsgType=0x0002100B) PSRP fragment.
        /// If we have the outgoing PUBLIC_KEY fragment saved, we extract the client's RSA public key,
        /// generate a random 32-byte session key, RSA-encrypt it (PKCS1v15), and inject that.
        /// The PS SDK will decrypt it successfully with its private key and complete key exchange,
        /// transitioning the runspace to Opened without needing to disable key exchange via reflection.
        /// </summary>
        private void InjectSyntheticEncryptedSessionKey()
        {
            try
            {
                string encryptedKeyB64 = BuildEncryptedSessionKeyPayload();
                TryStartKeyExchangeLocally();
                WinRMTrace.WriteLine($"[WINRM-DBG {T}] Injecting synthetic EncryptedSessionKey");
                if (TryRaiseKeyExchangeMessage("GenerateEncryptedSessionKeyResponse", _runspaceId, encryptedKeyB64))
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] Synthetic EncryptedSessionKey via hook injected OK");
                    WinRMTrace.WriteLine($"[WINRM-DBG] Session state immediately after hook: {GetCurrentSessionState()}");
                    Thread.Sleep(50);
                    WinRMTrace.WriteLine($"[WINRM-DBG] Session state 50ms after hook: {GetCurrentSessionState()}");
                    return;
                }

                WinRMTrace.WriteLine("[WINRM-DBG] Hook EncryptedSessionKey injection failed; falling back to encoded data");
                if (!TryInjectEncodedPsrpMessage("GenerateEncryptedSessionKeyResponse", _runspaceId, encryptedKeyB64))
                {
                    WinRMTrace.WriteLine("[WINRM-DBG] Encoded EncryptedSessionKey injection failed; trying local key completion fallback");
                    if (!TryCompleteKeyExchangeLocally(encryptedKeyB64))
                    {
                        WinRMTrace.WriteLine("[WINRM-DBG] Failed to inject EncryptedSessionKey");
                        return;
                    }

                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] Local key completion fallback succeeded");
                    return;
                }
                WinRMTrace.WriteLine($"[WINRM-DBG {T}] Synthetic EncryptedSessionKey via hook injected OK");
                WinRMTrace.WriteLine($"[WINRM-DBG] Session state immediately after hook: {GetCurrentSessionState()}");
                Thread.Sleep(50);
                WinRMTrace.WriteLine($"[WINRM-DBG] Session state 50ms after hook: {GetCurrentSessionState()}");
            }
            catch (Exception ex)
            {
                WinRMTrace.WriteLine($"[WINRM-DBG] InjectSyntheticEncryptedSessionKey failed: {ex.GetType().Name}: {ex.Message}");
            }
        }

        /// <summary>
        /// Builds the base64-encoded encrypted session key to use in the injected ENCRYPTED_SESSION_KEY message.
        /// If <see cref="_keyExchangePublicKeyB64"/> was saved from the outgoing PUBLIC_KEY fragment, this method
        /// extracts the client's RSA public key from it, generates a random 32-byte session key, and returns its
        /// RSA-PKCS1v15 encryption as base64.  If anything fails it falls back to "AAAA" (which will fail RSA
        /// decryption but is harmless when key exchange gets disabled via reflection).
        /// </summary>
        private string BuildEncryptedSessionKeyPayload()
        {
            // Strategy: export the client's RSA public key from the live crypto helper,
            // encrypt a random AES session key with it, then wrap the RSA ciphertext in the
            // CAPI SIMPLEBLOB format that ImportEncryptedSessionKey expects.
            if (_cryptoHelper != null)
            {
                try
                {
                    var helperType = _cryptoHelper.GetType();
                    var method = helperType.GetMethod(
                        "ExportLocalPublicKey",
                        System.Reflection.BindingFlags.Instance |
                        System.Reflection.BindingFlags.Public   |
                        System.Reflection.BindingFlags.NonPublic);

                    if (method != null)
                    {
                        // ExportLocalPublicKey has one out/ref string parameter.
                        object[] args = new object[] { null! };
                        method.Invoke(_cryptoHelper, args);
                        string? pubKeyB64 = args[0] as string;

                        WinRMTrace.WriteLine($"[WINRM-DBG] ExportLocalPublicKey result (first 60): {pubKeyB64?.Substring(0, Math.Min(60, pubKeyB64?.Length ?? 0))}");

                        if (!string.IsNullOrEmpty(pubKeyB64))
                        {
                            // pubKeyB64 = base64-encoded Windows CAPI PUBLICKEYBLOB
                            byte[] cspBlob = Convert.FromBase64String(pubKeyB64);
                            WinRMTrace.WriteLine($"[WINRM-DBG] CSP blob length: {cspBlob.Length}, type byte: 0x{cspBlob[0]:X2}");

                            byte[] sessionKey = new byte[32];
                            System.Security.Cryptography.RandomNumberGenerator.Fill(sessionKey);

                            var rsaProvider = new System.Security.Cryptography.RSACryptoServiceProvider();
                            rsaProvider.ImportCspBlob(cspBlob);
                            byte[] enc = rsaProvider.Encrypt(sessionKey, false); // false = PKCS1v15
                            byte[] simpleBlob = new byte[12 + enc.Length];
                            simpleBlob[0] = 0x01; // SIMPLEBLOB
                            simpleBlob[1] = 0x02; // CUR_BLOB_VERSION
                            simpleBlob[4] = 0x10; // CALG_AES_256 low byte
                            simpleBlob[5] = 0x66; // CALG_AES_256 high byte
                            simpleBlob[9] = 0xA4; // CALG_RSA_KEYX low byte
                            for (int i = 0; i < enc.Length; i++)
                            {
                                simpleBlob[12 + i] = enc[enc.Length - 1 - i];
                            }

                            WinRMTrace.WriteLine($"[WINRM-DBG] RSA encrypt OK ({enc.Length} bytes), SIMPLEBLOB len={simpleBlob.Length}");
                            return Convert.ToBase64String(simpleBlob);
                        }
                    }
                }
                catch (Exception ex)
                {
                    WinRMTrace.WriteLine($"[WINRM-DBG] RSA via ExportLocalPublicKey failed: {ex.GetType().Name}: {ex.Message}");
                }
            }

            WinRMTrace.WriteLine("[WINRM-DBG] RSA key extraction failed; using AAAA fallback");
            return "AAAA";
        }


        /// the command stream; this task drains those until the command completes or the
        /// shell is closed. The main reader thread handles shell-level messages in parallel.
        /// <para>
        /// When <paramref name="isKeyExchangePipeline"/> is true, the server is expected to
        /// complete key exchange by sending PipelineState=Done without EncryptedSessionKey
        /// (PS5.1 WinRM behaviour). After PipelineState arrives we inject a synthetic
        /// EncryptedSessionKey so the PS SDK's state machine can transition to Opened.
        /// </para>
        /// </summary>
        private void DrainCommandReceive(string commandId, CancellationToken ct)
        {
            var cmdFragGuids = new Dictionary<ulong, string>();
            try
            {
                while (!_isClosed && !ct.IsCancellationRequested)
                {
                    string endpoint = _connectionInfo.BuildEndpointUri();
                    string msgId = Guid.NewGuid().ToString().ToUpper();
                    string soapXml = WsManXml.BuildReceive(endpoint, msgId, _shellId!, commandId, _connectionInfo.ConfigurationName);

                    string response;
                    try { response = PostSoapAsync(endpoint, soapXml, WsManXml.ActionReceive).GetAwaiter().GetResult(); }
                    catch when (_isClosed || ct.IsCancellationRequested) { break; }
                    catch { break; }

                    if (string.IsNullOrWhiteSpace(response)) continue;

                    var parsedStreams = WsManXml.ParseStreams(response)
                        .OrderBy(stream =>
                        {
                            try
                            {
                                byte[] bytes = Convert.FromBase64String(stream.Item2);
                                return IsTerminalCommandState(bytes) ? 1 : 0;
                            }
                            catch
                            {
                                return 0;
                            }
                        })
                        .ToList();

                    bool commandDone = false;
                    foreach (var (name, base64Data) in parsedStreams)
                    {
                        WinRMTrace.WriteLine($"[WINRM-DBG {T}] CMD-DRAIN stream={name} b64len={base64Data.Length}");
                        TryInjectPendingEncryptedSessionKey("Command receive active");
                        try
                        {
                            byte[] fb = Convert.FromBase64String(base64Data);
                            bool observedRealEncryptedSessionKey = false;
                            string pipelineStateXmlSuffix = string.Empty;
                            if (fb.Length >= 29)
                            {
                                byte flags = fb[16];
                                bool isStart = (flags & 0x02) != 0;
                                uint mt = BitConverter.ToUInt32(fb, 25);
                                string rpidStr = string.Empty;
                                string pidStr = string.Empty;
                                if (isStart && fb.Length >= 61)
                                {
                                    var rpid = new Guid(fb[29..45]);
                                    var pid = new Guid(fb[45..61]);
                                    rpidStr = $" RPID={rpid}";
                                    pidStr = $" PID={pid}";
                                    if (_commandPipelineGuid != null &&
                                        pid != Guid.Empty &&
                                        !pid.ToString().Equals(_commandPipelineGuid, StringComparison.OrdinalIgnoreCase))
                                    {
                                        WinRMTrace.WriteLine($"[WINRM-DBG {T}] CMD-DRAIN pipeline mismatch expected={_commandPipelineGuid} actual={pid}");
                                    }
                                }

                                if (mt == 0x00010006)
                                {
                                    observedRealEncryptedSessionKey = true;
                                }
                                if (mt == 0x00041006) // PIPELINE_STATE terminal → command done
                                {
                                    string xml = Encoding.UTF8.GetString(fb, 61, fb.Length - 61).TrimStart('\uFEFF');
                                    pipelineStateXmlSuffix = $" XML={xml}";
                                    if (Regex.IsMatch(xml, @"<I32[^>]+>[3-9]<"))
                                    {
                                        commandDone = true;
                                        if (!_keyExchangeInjected)
                                        {
                                            WinRMTrace.WriteLine($"[WINRM-DBG {T}] PipelineState=Done observed for command stream");
                                        }
                                    }
                                }
                                WinRMTrace.WriteLine($"[WINRM-DBG] CMD-DRAIN MsgType=0x{mt:X8} Flags=0x{flags:X2}{rpidStr}{pidStr}{pipelineStateXmlSuffix}");
                            }
                            try
                            {
                                HandleOutputDataReceived(WsManXml.WrapInDataElement(base64Data, cmdFragGuids));
                                if (observedRealEncryptedSessionKey)
                                {
                                    ObserveRealEncryptedSessionKey("command receive");
                                }
                                if (fb.Length >= 29 && BitConverter.ToUInt32(fb, 25) == 0x00041006)
                                {
                                    WinRMTrace.WriteLine($"[WINRM-DBG] Session state after command pipeline state dispatch: {GetCurrentSessionState()}");
                                }
                            }
                            catch
                            {
                                try
                                {
                                    ProcessRawDataCompat(fb, name);
                                    if (observedRealEncryptedSessionKey)
                                    {
                                        ObserveRealEncryptedSessionKey("command receive");
                                    }
                                    if (fb.Length >= 29 && BitConverter.ToUInt32(fb, 25) == 0x00041006)
                                    {
                                        WinRMTrace.WriteLine($"[WINRM-DBG] Session state after command pipeline state dispatch: {GetCurrentSessionState()}");
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }

                    // Also check for CommandState=Done (no rsp:Stream, just rsp:CommandState)
                    if (parsedStreams.Count == 0 && response.Contains("CommandState"))
                        commandDone = true;

                    if (commandDone) break;
                }
            }
            catch { }
        }

        private static bool IsTerminalCommandState(byte[] fragmentBytes)
        {
            if (fragmentBytes.Length < 61)
            {
                return false;
            }

            if (BitConverter.ToUInt32(fragmentBytes, 25) != 0x00041006)
            {
                return false;
            }

            string xml = Encoding.UTF8.GetString(fragmentBytes, 61, fragmentBytes.Length - 61).TrimStart('\uFEFF');
            return Regex.IsMatch(xml, @"<I32[^>]+>[3-9]<");
        }

        /// <summary>
        /// Reader thread: continuously polls the server with WSMan Receive (shell-level) requests
        /// and feeds returned PSRP lines into the base class via HandleOutputDataReceived.
        /// RunCommand pipeline messages are handled by DrainCommandReceive running in parallel.
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
                        // Always use shell-level Receive (commandId=null) for RunspacePool messages.
                        // Command-level data (EncryptedSessionKey, PipelineState) is handled by
                        // DrainCommandReceive running in parallel after each RunCommand.
                        string soapXml = WsManXml.BuildReceive(endpoint, msgId, _shellId!, null, _connectionInfo.ConfigurationName);

                        WinRMTrace.WriteLine($"[WINRM-DBG] RECEIVE polling...");
                        string response = PostSoapAsync(endpoint, soapXml, WsManXml.ActionReceive)
                            .GetAwaiter().GetResult();
                        WinRMTrace.WriteLine($"[WINRM-DBG] RECEIVE got {response.Length} chars");

                        if (string.IsNullOrWhiteSpace(response))
                            continue;

                        // Log all stream names and full body when no streams found
                        bool hasAnyStream = WsManXml.ParseStreams(response).Any();
                        if (!hasAnyStream)
                            WinRMTrace.WriteLine($"[WINRM-DBG] RECEIVE no-streams body={response}");
                        else
                        {
                            foreach (var (streamName, _) in WsManXml.ParseStreams(response))
                                if (streamName != "stdout")
                                    WinRMTrace.WriteLine($"[WINRM-DBG] RECEIVE non-stdout stream: {streamName}");
                        }

                        // The WinRM SOAP stream carries raw base64 PSRP fragments.
                        // Feed raw PSRP bytes into the transport manager so the SDK's normal
                        // fragment parsing and key-exchange handlers see the same shape of data
                        // they would get from a native WSMan transport.
                        foreach (var (name, base64Data) in WsManXml.ParseStreams(response))
                        {
                            // Decode and log fragment details for diagnostics
                            try
                            {
                                byte[] fragBytes = Convert.FromBase64String(base64Data);
                                bool observedRealEncryptedSessionKey = false;
                                if (fragBytes.Length >= 21)
                                {
                                    ulong objId = BitConverter.ToUInt64(fragBytes, 0);
                                    ulong fragId = BitConverter.ToUInt64(fragBytes, 8);
                                    byte flags = fragBytes[16];
                                    bool isStart = (flags & 0x02) != 0;
                                    bool isEnd = (flags & 0x01) != 0;
                                    uint blobLen = BitConverter.ToUInt32(fragBytes, 17);

                                    string msgTypeStr = "";
                                    string rpidStr = "";
                                    string pidStr = "";
                                    if (isStart && fragBytes.Length >= 61)
                                    {
                                        uint msgType = BitConverter.ToUInt32(fragBytes, 25);
                                        var rpid = new Guid(fragBytes[29..45]);
                                        var pid = new Guid(fragBytes[45..61]);
                                        msgTypeStr = $" MsgType=0x{msgType:X8}";
                                        rpidStr = $" RPID={rpid}";
                                        pidStr = $" PID={pid}";

                                        if (msgType == 0x00021005)
                                        {
                                            string xml = Encoding.UTF8.GetString(fragBytes, 61, fragBytes.Length - 61).TrimStart('\uFEFF');
                                            if (Regex.IsMatch(xml, @"<I32[^>]*>2<"))
                                            {
                                                WinRMTrace.WriteLine($"[WINRM-DBG {T}] RunspacePool opened without synthetic PublicKeyRequest injection");
                                            }
                                        }
                                        else if (msgType == 0x00010006)
                                        {
                                            observedRealEncryptedSessionKey = true;
                                        }
                                    }
                                    WinRMTrace.WriteLine($"[WINRM-DBG {T}] RECV stream={name} b64len={base64Data.Length} fragLen={fragBytes.Length} ObjId={objId} FragId={fragId} Flags=0x{flags:X2}(Start={isStart},End={isEnd}) BlobLen={blobLen}{msgTypeStr}{rpidStr}{pidStr}");
                                    WinRMTrace.WriteLine($"[WINRM-DBG] RECV FULL_B64={base64Data}");
                                    ProcessRawDataCompat(fragBytes, name);
                                    if (observedRealEncryptedSessionKey)
                                    {
                                        ObserveRealEncryptedSessionKey("session receive");
                                    }
                                }
                            }
                            catch (Exception dbgEx)
                            {
                                WinRMTrace.WriteLine($"[WINRM-DBG] decode error: {dbgEx.Message}");
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

            // Send WSMan Signal(terminate) to end the command, then Delete the shell.
            if (_shellId != null && _httpClient != null)
            {
                try
                {
                    string endpoint = _connectionInfo.BuildEndpointUri();
                    if (_commandId != null)
                    {
                        string sigMsgId = Guid.NewGuid().ToString().ToUpper();
                        string sigSoap = WsManXml.BuildSignal(endpoint, sigMsgId, _shellId, _commandId, _connectionInfo.ConfigurationName);
                        PostSoapAsync(endpoint, sigSoap, WsManXml.ActionSignal).Wait(3000);
                    }

                    string msgId = Guid.NewGuid().ToString().ToUpper();
                    string soapXml = WsManXml.BuildDelete(endpoint, msgId, _shellId, _connectionInfo.ConfigurationName);
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

