using System;
using System.Collections;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
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

        protected const int DefaultOpenTimeoutMs = 30000;
        protected const int DefaultReadinessTimeoutMs = 5000;
        protected const int DefaultSSHPort = 22;
        protected const string DefaultSSHSubsystem = "powershell";

        protected RunspaceConnectionInfo? _connectionInfo;
        protected Runspace? _runspace;
        protected ManualResetEvent? _openAsync;

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
        [Alias("ComputerName")]
        public string? HostName { get; set; }

        /// <summary>
        /// Port number. Required for TCP, optional for SSH (default 22).
        /// </summary>
        [Parameter(ParameterSetName = TcpParameterSet, Mandatory = true)]
        [Parameter(ParameterSetName = SSHParameterSet)]
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
        /// SSH user name (can include domain, e.g., "Administrator@domain")
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
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
        [Credential()]
        public PSCredential? Credential { get; set; }

        /// <summary>
        /// Skip host key verification for SSH connections (not recommended for production)
        /// </summary>
        [Parameter(ParameterSetName = SSHParameterSet)]
        public SwitchParameter SkipHostKeyCheck { get; set; }

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
            _runspace.StateChanged += HandleRunspaceStateChanged;

            try
            {
                _runspace.OpenAsync();

                // Wait for runspace to reach Opened/Closed/Broken state with timeout
                if (!_openAsync.WaitOne(OpenTimeout))
                {
                    // Timeout - clean up and throw
                    try { _runspace.Dispose(); } catch { }
                    throw new TimeoutException($"Runspace open timed out after {OpenTimeout}ms");
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

        protected override void StopProcessing()
        {
            ReleaseWait();
        }

        protected void HandleRunspaceStateChanged(object? source, RunspaceStateEventArgs stateEventArgs)
        {
            switch (stateEventArgs.RunspaceStateInfo.State)
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
