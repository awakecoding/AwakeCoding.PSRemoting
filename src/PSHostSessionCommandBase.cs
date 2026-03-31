using System;
using System.Collections;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Abstract base class for PSHostSession cmdlets that share common parameters and connection logic
    /// </summary>
    public abstract class PSHostSessionCommandBase : PSCmdlet
    {
        protected const string SubprocessParameterSet = "Subprocess";
        protected const string TcpParameterSet = "TCP";
        protected const string WebSocketParameterSet = "WebSocket";
        protected const string NamedPipeParameterSet = "NamedPipe";
        protected const string ProcessIdParameterSet = "ProcessId";
        protected const string SSHParameterSet = "SSH";
        protected const string WinRMComputerNameParameterSet = "WinRMComputerName";
        protected const string WinRMConnectionUriParameterSet = "WinRMConnectionUri";

        protected const int DefaultOpenTimeoutMs = 30000;
        protected const int DefaultReadinessTimeoutMs = 5000;
        protected const int DefaultSSHPort = 22;
        protected const string DefaultSSHSubsystem = "powershell";

        protected RunspaceConnectionInfo? _connectionInfo;
        protected Runspace? _runspace;
        protected ManualResetEvent? _openAsync;
        protected DateTime _runspaceOpenStart;

        #region Subprocess Parameters

        [Parameter(ParameterSetName = SubprocessParameterSet)]
        [ValidateNotNullOrEmpty()]
        public string? ExecutablePath { get; set; }

        [Parameter(ParameterSetName = SubprocessParameterSet)]
        public SwitchParameter UseWindowsPowerShell { get; set; }

        [Parameter(ParameterSetName = SubprocessParameterSet)]
        [ValidateNotNull()]
        public string[]? ArgumentList { get; set; }

        #endregion

        #region TCP and SSH Shared Parameters

        /// <summary>
        /// Target host name or IP address for TCP or SSH connections
        /// </summary>
        [Parameter(ParameterSetName = TcpParameterSet, Mandatory = true)]
        [Parameter(ParameterSetName = SSHParameterSet, Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        public string? HostName { get; set; }

        /// <summary>
        /// Port number. Required for TCP, optional for SSH (default 22) and WinRM (default 5985/5986).
        /// </summary>
        [Parameter(ParameterSetName = TcpParameterSet, Mandatory = true)]
        [Parameter(ParameterSetName = SSHParameterSet)]
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [ValidateRange(1, 65535)]
        public int Port { get; set; }

        #endregion

        #region SSH Parameters

        /// <summary>
        /// Switch to indicate SSH transport should be used.
        /// Required to disambiguate from TCP when using -HostName.
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet, Mandatory = true)]
        public SwitchParameter SSHTransport { get; set; }

        /// <summary>
        /// SSH or WinRM user name (can include domain, e.g., "Administrator@domain").
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet)]
        [ValidateNotNullOrEmpty()]
        public string? UserName { get; set; }

        /// <summary>
        /// Path to SSH private key file for key-based authentication
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        [ValidateNotNullOrEmpty()]
        [Alias("IdentityFilePath")]
        public string? KeyFilePath { get; set; }

        /// <summary>
        /// SSH subsystem to request (default "powershell")
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        [ValidateNotNullOrEmpty()]
        public string Subsystem { get; set; } = DefaultSSHSubsystem;

        /// <summary>
        /// Timeout in milliseconds for SSH connection attempt.
        /// Default is -1 (infinite, wait indefinitely).
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        [ValidateRange(-1, int.MaxValue)]
        public int ConnectingTimeout { get; set; } = -1;

        /// <summary>
        /// Additional SSH options as hashtable (e.g., @{StrictHostKeyChecking="no"})
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        public Hashtable? Options { get; set; }

        /// <summary>
        /// Path to SSH executable (auto-detected if not specified)
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        [ValidateNotNullOrEmpty()]
        public string? SSHExecutablePath { get; set; }

        /// <summary>
        /// Credential for SSH password authentication
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet)]
        [Credential()]
        public PSCredential? Credential { get; set; }

        /// <summary>
        /// Skip host key verification for SSH connections (not recommended for production)
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        public SwitchParameter SkipHostKeyCheck { get; set; }

        #endregion

        #region WinRM Parameters

        /// <summary>
        /// Target computer name for WinRM connections.
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet, Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        public string? ComputerName { get; set; }

        /// <summary>
        /// WinRM endpoint URI (for example, http://server01:5985/wsman).
        /// </summary>
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet, Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        public Uri? ConnectionUri { get; set; }

        /// <summary>
        /// Use HTTPS for WinRM (port 5986 by default).
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        public SwitchParameter UseSSL { get; set; }

        /// <summary>
        /// Authentication mechanism for WinRM (Default = Negotiate).
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet)]
        public WinRMAuthenticationMechanism Authentication { get; set; } = WinRMAuthenticationMechanism.Negotiate;

        /// <summary>
        /// WinRM password to combine with -UserName when you do not want to construct a PSCredential yourself.
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet)]
        [ValidateNotNull()]
        public SecureString? Password { get; set; }

        /// <summary>
        /// Use the current Windows logon token for WinRM Negotiate/Kerberos instead of prompting for explicit credentials.
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet)]
        public SwitchParameter UseImplicitCredential { get; set; }

        /// <summary>
        /// Allow HTTP redirects for WinRM connections.
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet)]
        public SwitchParameter AllowRedirection { get; set; }

        /// <summary>
        /// WSMan application name (default "/wsman").
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [ValidateNotNullOrEmpty()]
        public string ApplicationName { get; set; } = "/wsman";

        /// <summary>
        /// PowerShell session configuration name (default "Microsoft.PowerShell").
        /// Use "PowerShell.7" to connect to a PS7 endpoint.
        /// </summary>
        [Parameter(ParameterSetName = WinRMComputerNameParameterSet)]
        [Parameter(ParameterSetName = WinRMConnectionUriParameterSet)]
        [ValidateNotNullOrEmpty()]
        public string ConfigurationName { get; set; } = WsManXml.DefaultConfigurationName;

        #endregion

        #region WebSocket Parameters

        /// <summary>
        /// WebSocket URI (e.g., ws://localhost:8080/pwsh or wss://localhost:8443/pwsh)
        /// </summary>
        [Parameter(ParameterSetName = WebSocketParameterSet, Mandatory = true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [Alias("Url")]
        public Uri? Uri { get; set; }

        #endregion

        #region NamedPipe Parameters

        [Parameter(ParameterSetName = NamedPipeParameterSet, Mandatory = true)]
        [ValidateNotNullOrEmpty()]
        public string? PipeName { get; set; }

        [Parameter(ParameterSetName = ProcessIdParameterSet, Mandatory = true)]
        [Alias("Id")]
        public int ProcessId { get; set; }

        [Parameter(ParameterSetName = ProcessIdParameterSet)]
        [ValidateNotNullOrEmpty()]
        public string AppDomainName { get; set; } = "DefaultAppDomain";

        #endregion

        #region Common Parameters

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string RunspaceName { get; set; } = "PSHostClient";

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string? TransportName { get; set; }

        /// <summary>
        /// Timeout in milliseconds for opening the runspace connection.
        /// Default is 30000 (30 seconds).
        /// </summary>
        [Parameter()]
        [ValidateRange(1000, 300000)]
        public int OpenTimeout { get; set; } = DefaultOpenTimeoutMs;

        #endregion

        #region Protected Shared Methods

        /// <summary>
        /// Creates and opens a runspace based on the current parameter set
        /// </summary>
        protected Runspace CreateAndOpenRunspace()
        {
            // Create the appropriate connection info based on parameter set
            switch (ParameterSetName)
            {
                case TcpParameterSet:
                    _connectionInfo = CreateTcpConnectionInfo();
                    TransportName ??= "PSHostTcp";
                    break;

                case WebSocketParameterSet:
                    _connectionInfo = CreateWebSocketConnectionInfo();
                    TransportName ??= "PSHostWebSocket";
                    break;

                case NamedPipeParameterSet:
                case ProcessIdParameterSet:
                    _connectionInfo = CreateNamedPipeConnectionInfo();
                    TransportName ??= "PSHostNamedPipe";
                    break;

                case SSHParameterSet:
                    _connectionInfo = CreateSSHConnectionInfo();
                    TransportName ??= "SSH";
                    break;

                case WinRMComputerNameParameterSet:
                case WinRMConnectionUriParameterSet:
                    _connectionInfo = CreateWinRMConnectionInfo();
                    TransportName ??= "WinRM";
                    break;

                case SubprocessParameterSet:
                default:
                    _connectionInfo = CreateSubprocessConnectionInfo();
                    TransportName ??= "PSHostSession";
                    break;
            }

            _runspace = RunspaceFactory.CreateRunspace(
                connectionInfo: _connectionInfo,
                host: Host,
                typeTable: TypeTable.LoadDefaultTypeFiles(),
                applicationArguments: null,
                name: RunspaceName);

            _openAsync = new ManualResetEvent(false);
            _runspaceOpenStart = DateTime.UtcNow;
            _runspace.StateChanged += HandleRunspaceStateChanged;

            try
            {
                _runspace.OpenAsync();

                // Wait for runspace to reach Opened/Closed/Broken state with timeout.
                int effectiveTimeout = Math.Max(OpenTimeout, _connectionInfo.OpenTimeout);
                if (!_openAsync.WaitOne(effectiveTimeout))
                {
                    // Timeout - clean up and throw
                    try { _runspace.Dispose(); } catch { }
                    throw new TimeoutException($"Runspace open timed out after {effectiveTimeout}ms");
                }

                // Check if runspace opened successfully
                if (_runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                {
                    var reason = _runspace.RunspaceStateInfo.Reason;
                    throw new InvalidOperationException(
                        $"Failed to open runspace: {reason?.Message ?? "Unknown error"}",
                        reason);
                }

                // Wait for runspace to be fully ready (Available)
                if (!RunspaceReadinessHelper.WaitForRunspaceReady(_runspace, DefaultReadinessTimeoutMs))
                {
                    WriteWarning("Runspace opened but may not be fully ready for commands");
                }

                return _runspace;
            }
            finally
            {
                _openAsync.Dispose();
            }
        }

        private RunspaceConnectionInfo CreateSubprocessConnectionInfo()
        {
            string computerName = "localhost";
            
            // Default arguments for PowerShell server mode
            string[] arguments;
            if (ArgumentList != null && ArgumentList.Length > 0)
            {
                arguments = ArgumentList;
            }
            else
            {
                arguments = new[] { "-NoLogo", "-NoProfile", "-s" };
            }

            string? executable;

            if (!string.IsNullOrWhiteSpace(ExecutablePath))
            {
                executable = ExecutablePath;
            }
            else
            {
                executable = PowerShellFinder.GetExecutablePath(UseWindowsPowerShell);
            }

            if (executable == null || !File.Exists(executable))
            {
                throw new InvalidOperationException("The PowerShell executable path could not be found");
            }

            return new PSHostClientInfo(computerName, executable, arguments);
        }

        private RunspaceConnectionInfo CreateTcpConnectionInfo()
        {
            if (string.IsNullOrWhiteSpace(HostName))
            {
                throw new ArgumentException("HostName is required for TCP connections");
            }

            var connectionInfo = new PSHostTcpClientInfo(HostName, Port)
            {
                OpenTimeout = OpenTimeout
            };

            return connectionInfo;
        }

        private RunspaceConnectionInfo CreateWebSocketConnectionInfo()
        {
            if (Uri == null)
            {
                throw new ArgumentException("Uri is required for WebSocket connections");
            }

            // Validate WebSocket scheme
            if (!Uri.Scheme.Equals("ws", StringComparison.OrdinalIgnoreCase) &&
                !Uri.Scheme.Equals("wss", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"Invalid WebSocket URI scheme: {Uri.Scheme}. Expected 'ws' or 'wss'.");
            }

            var connectionInfo = new PSHostWebSocketClientInfo(Uri)
            {
                OpenTimeout = OpenTimeout
            };

            return connectionInfo;
        }

        private RunspaceConnectionInfo CreateNamedPipeConnectionInfo()
        {
            string pipeName;

            if (!string.IsNullOrWhiteSpace(PipeName))
            {
                // Direct pipe name provided
                pipeName = PipeName;
            }
            else if (ProcessId > 0)
            {
                // Generate pipe name from process ID
                pipeName = PowerShellFinder.GetProcessPipeName(ProcessId, AppDomainName);
            }
            else
            {
                throw new ArgumentException("Either PipeName or ProcessId must be provided");
            }

            var connectionInfo = new PSHostNamedPipeInfo("localhost", pipeName, AppDomainName)
            {
                OpenTimeout = OpenTimeout
            };

            return connectionInfo;
        }

        private RunspaceConnectionInfo CreateSSHConnectionInfo()
        {
// Use the new SSHClientInfo with thread-based transport for interactive console passthrough
            var connectionInfo = new SSHClientInfo(
                computerName: HostName!,
                userName: UserName,
                keyFilePath: KeyFilePath,
                port: Port > 0 ? Port : DefaultSSHPort,
                subsystem: Subsystem,
                connectingTimeout: ConnectingTimeout,
                options: Options,
                sshExecutablePath: SSHExecutablePath);

            // Set credential for password authentication
            if (Credential != null)
            {
                connectionInfo.Credential = Credential;
            }

            // Set skip host key check if specified
            if (SkipHostKeyCheck)
            {
                connectionInfo.SkipHostKeyCheck = true;
            }

            // Pass PSHost for interactive prompting
            connectionInfo.PSHost = Host;

            return connectionInfo;
        }

        private RunspaceConnectionInfo CreateWinRMConnectionInfo()
        {
            string computerName;
            int port;
            bool useSsl;
            string applicationName;

            if (ParameterSetName == WinRMConnectionUriParameterSet)
            {
                if (ConnectionUri == null)
                {
                    throw new ArgumentException("ConnectionUri is required for WinRM URI connections");
                }

                if (!ConnectionUri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase)
                    && !ConnectionUri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException($"Invalid WinRM URI scheme: {ConnectionUri.Scheme}. Expected 'http' or 'https'.");
                }

                applicationName = string.IsNullOrEmpty(ConnectionUri.AbsolutePath)
                    ? "/wsman"
                    : ConnectionUri.AbsolutePath;
                computerName = ConnectionUri.Host;
                port = ConnectionUri.IsDefaultPort ? 0 : ConnectionUri.Port;
                useSsl = ConnectionUri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                computerName = ComputerName!;
                port = Port > 0 ? Port : (UseSSL ? 5986 : 5985);
                useSsl = UseSSL;
                applicationName = ApplicationName;
            }

            (PSCredential? credential, bool useImplicitCredential) = ResolveWinRMCredential(computerName);

            return new WinRMClientInfo(computerName)
            {
                Port = port,
                UseSSL = useSsl,
                Authentication = Authentication,
                Credential = credential,
                UseImplicitCredential = useImplicitCredential,
                ApplicationName = applicationName,
                ConfigurationName = ConfigurationName,
                AllowRedirection = AllowRedirection,
                OpenTimeout = OpenTimeout
            };
        }

        private (PSCredential? Credential, bool UseImplicitCredential) ResolveWinRMCredential(string targetName)
        {
            bool hasCredential = Credential != null;
            bool hasUserName = !string.IsNullOrWhiteSpace(UserName);
            bool hasPassword = Password != null;

            if (UseImplicitCredential)
            {
                if (Authentication == WinRMAuthenticationMechanism.Basic)
                {
                    throw new InvalidOperationException("WinRM implicit credentials cannot be used with Basic authentication.");
                }

                if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    throw new PlatformNotSupportedException("WinRM implicit credentials are only supported on Windows. Supply -Credential or -UserName with -Password instead.");
                }

                if (hasCredential || hasUserName || hasPassword)
                {
                    throw new InvalidOperationException("Do not combine -UseImplicitCredential with -Credential, -UserName, or -Password.");
                }

                return (null, true);
            }

            if (hasCredential)
            {
                if (hasUserName || hasPassword)
                {
                    throw new InvalidOperationException("Do not combine -Credential with -UserName or -Password for WinRM connections.");
                }

                return (Credential, false);
            }

            if (hasPassword && !hasUserName)
            {
                throw new InvalidOperationException("Specify -UserName when using -Password for WinRM connections.");
            }

            if (hasUserName && hasPassword)
            {
                return (new PSCredential(UserName!, Password!), false);
            }

            return (PromptForWinRMCredential(targetName), false);
        }

        private PSCredential PromptForWinRMCredential(string targetName)
        {
            if (Host?.UI == null)
            {
                throw new InvalidOperationException(
                    "WinRM connections require explicit credentials by default. Supply -Credential or -UserName with -Password, or use -UseImplicitCredential on Windows.");
            }

            try
            {
                PSCredential? credential = Host.UI.PromptForCredential(
                    caption: "WinRM Authentication",
                    message: $"Enter credentials for the WinRM connection to '{targetName}'.",
                    userName: UserName ?? string.Empty,
                    targetName: targetName);

                return credential ?? throw new InvalidOperationException("The WinRM credential prompt was cancelled.");
            }
            catch (System.Management.Automation.Host.HostException ex)
            {
                throw new InvalidOperationException(
                    "WinRM connections require explicit credentials by default, but the current host cannot prompt for them. Supply -Credential or -UserName with -Password, or use -UseImplicitCredential on Windows.",
                    ex);
            }
            catch (NotImplementedException ex)
            {
                throw new InvalidOperationException(
                    "WinRM connections require explicit credentials by default, but the current host cannot prompt for them. Supply -Credential or -UserName with -Password, or use -UseImplicitCredential on Windows.",
                    ex);
            }
        }

        protected override void StopProcessing()
        {
            ReleaseWait();
        }

        protected void HandleRunspaceStateChanged(object? source, RunspaceStateEventArgs stateEventArgs)
        {
            var state = stateEventArgs.RunspaceStateInfo.State;
            WinRMTrace.WriteLine($"[WINRM-DBG +{(DateTime.UtcNow - _runspaceOpenStart).TotalMilliseconds:F0}ms] RUNSPACE STATE: {state}  Reason={stateEventArgs.RunspaceStateInfo.Reason?.Message}");
            switch (state)
            {
                case RunspaceState.Opened:
                case RunspaceState.Closed:
                case RunspaceState.Broken:
                    if (_runspace != null)
                    {
                        _runspace.StateChanged -= HandleRunspaceStateChanged;
                    }
                    ReleaseWait();
                    break;
            }
        }

        protected void ReleaseWait()
        {
            try
            {
                _openAsync?.Set();
            }
            catch (ObjectDisposedException)
            {

            }
        }

        #endregion

        #region Abstract Methods

        /// <summary>
        /// Process the opened runspace. Derived classes implement their specific behavior.
        /// </summary>
        /// <param name="runspace">The opened and ready runspace</param>
        protected abstract void ProcessSession(Runspace runspace);

        #endregion
    }
}
