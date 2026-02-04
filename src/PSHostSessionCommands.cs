using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Threading;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Helper class for runspace readiness checking
    /// </summary>
    internal static class RunspaceReadinessHelper
    {
        private const int DefaultReadinessTimeoutMs = 5000;
        private const int ReadinessPollIntervalMs = 50;

        /// <summary>
        /// Wait for runspace to be fully ready (Opened state + Available)
        /// </summary>
        public static bool WaitForRunspaceReady(Runspace runspace, int timeoutMs = DefaultReadinessTimeoutMs)
        {
            if (runspace == null)
                return false;

            var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

            // First wait for Opened state
            while (runspace.RunspaceStateInfo.State == RunspaceState.Opening && DateTime.UtcNow < deadline)
            {
                Thread.Sleep(ReadinessPollIntervalMs);
            }

            if (runspace.RunspaceStateInfo.State != RunspaceState.Opened)
            {
                return false;
            }

            // Then wait for Available (runspace ready to accept commands)
            while (runspace.RunspaceAvailability != RunspaceAvailability.Available && DateTime.UtcNow < deadline)
            {
                Thread.Sleep(ReadinessPollIntervalMs);
            }

            return runspace.RunspaceStateInfo.State == RunspaceState.Opened &&
                   runspace.RunspaceAvailability == RunspaceAvailability.Available;
        }
    }

    [Cmdlet(VerbsCommon.New, "PSHostSession", DefaultParameterSetName = SubprocessParameterSet)]
    [OutputType(typeof(PSSession))]
    public sealed class NewPSHostSessionCommand : PSCmdlet
    {
        private const string SubprocessParameterSet = "Subprocess";
        private const string TcpParameterSet = "TCP";
        private const string WebSocketParameterSet = "WebSocket";
        private const string NamedPipeParameterSet = "NamedPipe";
        private const string ProcessIdParameterSet = "ProcessId";
        private const string SSHParameterSet = "SSH";

        private const int DefaultOpenTimeoutMs = 30000;
        private const int DefaultReadinessTimeoutMs = 5000;
        private const int DefaultSSHPort = 22;
        private const string DefaultSSHSubsystem = "powershell";

        private RunspaceConnectionInfo? _connectionInfo;
        private Runspace? _runspace;
        private ManualResetEvent? _openAsync;

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

        protected override void BeginProcessing()
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

                WriteObject(
                    PSSession.Create(
                        runspace: _runspace,
                        transportName: TransportName,
                        psCmdlet: this));
            }
            finally
            {
                _openAsync.Dispose();
            }
        }

        private PSHostClientInfo CreateSubprocessConnectionInfo()
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

        private PSHostTcpClientInfo CreateTcpConnectionInfo()
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

        private PSHostWebSocketClientInfo CreateWebSocketConnectionInfo()
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

        private PSHostNamedPipeInfo CreateNamedPipeConnectionInfo()
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

        private SSHClientInfo CreateSSHConnectionInfo()
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

        private void HandleRunspaceStateChanged(object? source, RunspaceStateEventArgs stateEventArgs)
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

        private void ReleaseWait()
        {
            try
            {
                _openAsync?.Set();
            }
            catch (ObjectDisposedException)
            {

            }
        }
    }

    [Cmdlet(VerbsCommunications.Connect, "PSHostProcess", DefaultParameterSetName = "ProcessIdParameterSet")]
    [OutputType(typeof(PSSession))]
    public sealed class ConnectPSHostProcessCommand : PSCmdlet
    {
        private PSHostNamedPipeInfo? _connectionInfo;
        private Runspace? _runspace;
        private ManualResetEvent? _openAsync;

        [Parameter(ParameterSetName = "ProcessIdParameterSet", Position = 0, Mandatory = true)]
        [ValidateNotNull()]
        public int? Id { get; set; }

        [Parameter(ParameterSetName = "ProcessParameterSet", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        [ValidateNotNull()]
        public Process? Process { get; set; }

        [Parameter(ParameterSetName = "ProcessNameParameterSet", Position = 0, Mandatory = true)]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter(ParameterSetName = "PSHostProcessInfoParameterSet", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        [ValidateNotNull()]
        public PSObject? HostProcessInfo { get; set; }

        [Parameter(ParameterSetName = "CustomPipeNameParameterSet", Mandatory = true)]
        [ValidateNotNullOrEmpty()]
        public string? CustomPipeName { get; set; }

        [Parameter(ParameterSetName = "ProcessIdParameterSet")]
        [Parameter(ParameterSetName = "ProcessParameterSet")]
        [Parameter(ParameterSetName = "ProcessNameParameterSet")]
        [Parameter(ParameterSetName = "PSHostProcessInfoParameterSet")]
        [ValidateNotNullOrEmpty()]
        public string AppDomainName { get; set; } = "DefaultAppDomain";

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string RunspaceName { get; set; } = "PSHostClient";

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string TransportName { get; set; } = "PSHostProcess";

        /// <summary>
        /// Timeout in milliseconds for opening the connection.
        /// Default is 5000 (5 seconds).
        /// </summary>
        [Parameter()]
        [ValidateRange(1000, 300000)]
        public int OpenTimeout { get; set; } = 5000;

        /// <summary>
        /// Timeout in milliseconds waiting for runspace to be fully ready.
        /// Default is 5000 (5 seconds).
        /// </summary>
        [Parameter()]
        [ValidateRange(500, 60000)]
        public int ReadinessTimeout { get; set; } = 5000;

        protected override void BeginProcessing()
        {
            string computerName = "localhost";
            string pipeName;

            // Determine the pipe name based on parameter set
            if (!string.IsNullOrWhiteSpace(CustomPipeName))
            {
                pipeName = CustomPipeName;
            }
            else if (Id.HasValue)
            {
                pipeName = PowerShellFinder.GetProcessPipeName(Id.Value, AppDomainName);
            }
            else if (Process != null)
            {
                pipeName = PowerShellFinder.GetProcessPipeName(Process.Id, AppDomainName);
            }
            else if (!string.IsNullOrWhiteSpace(Name))
            {
                Process[] processes = Process.GetProcessesByName(Name);
                if (processes.Length == 0)
                {
                    throw new ItemNotFoundException($"No process found with name '{Name}'");
                }
                if (processes.Length > 1)
                {
                    throw new PSInvalidOperationException($"Multiple processes found with name '{Name}'");
                }
                pipeName = PowerShellFinder.GetProcessPipeName(processes[0].Id, AppDomainName);
                processes[0].Dispose();
            }
            else if (HostProcessInfo != null)
            {
                // Extract pipe name from PSHostProcessInfo object
                // PSHostProcessInfo has a method GetPipeNameFilePath() that returns the pipe name
                var method = HostProcessInfo.BaseObject?.GetType().GetMethod("GetPipeNameFilePath");
                if (method != null)
                {
                    string? fullPipePath = method.Invoke(HostProcessInfo.BaseObject, null) as string;
                    if (fullPipePath != null)
                    {
                        // Remove the platform-specific prefix to get the pipe name
                        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                        {
                            pipeName = fullPipePath.Replace(@"\\.\pipe\", string.Empty);
                        }
                        else
                        {
                            string tempPath = Path.GetTempPath();
                            pipeName = fullPipePath.Replace(Path.Combine(tempPath, "CoreFxPipe_"), string.Empty);
                        }
                    }
                    else
                    {
                        throw new InvalidOperationException("Failed to extract pipe name from HostProcessInfo");
                    }
                }
                else
                {
                    throw new InvalidOperationException("PSHostProcessInfo object does not have GetPipeNameFilePath method");
                }
            }
            else
            {
                throw new PSInvalidOperationException("No valid parameter combination provided");
            }

            _connectionInfo = new PSHostNamedPipeInfo(computerName, pipeName, AppDomainName);
            _connectionInfo.OpenTimeout = OpenTimeout;

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

                // Wait with timeout to prevent hanging on busy/unresponsive connections
                if (!_openAsync.WaitOne(OpenTimeout))
                {
                    // Timeout - clean up and throw
                    try { _runspace.Dispose(); } catch { }
                    throw new TimeoutException($"Named pipe connection timed out after {OpenTimeout}ms");
                }

                // Check if runspace opened successfully
                if (_runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                {
                    var reason = _runspace.RunspaceStateInfo.Reason;
                    throw new InvalidOperationException(
                        $"Failed to connect to process: {reason?.Message ?? "Unknown error"}",
                        reason);
                }

                // Wait for runspace to be fully ready (Available)
                if (!RunspaceReadinessHelper.WaitForRunspaceReady(_runspace, ReadinessTimeout))
                {
                    WriteWarning("Runspace opened but may not be fully ready for commands");
                }

                WriteObject(
                    PSSession.Create(
                        runspace: _runspace,
                        transportName: TransportName,
                        psCmdlet: this));
            }
            finally
            {
                _openAsync.Dispose();
            }
        }

        protected override void StopProcessing()
        {
            ReleaseWait();
        }

        private void HandleRunspaceStateChanged(object? source, RunspaceStateEventArgs stateEventArgs)
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

        private void ReleaseWait()
        {
            try
            {
                _openAsync?.Set();
            }
            catch (ObjectDisposedException)
            {

            }
        }
    }
}
