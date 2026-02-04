using System;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
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
    public sealed class NewPSHostSessionCommand : PSHostSessionCommandBase
    {
        protected override void BeginProcessing()
        {
            Runspace runspace = CreateAndOpenRunspace();
            ProcessSession(runspace);
        }

        protected override void ProcessSession(Runspace runspace)
        {
            WriteObject(
                PSSession.Create(
                    runspace: runspace,
                    transportName: TransportName,
                    psCmdlet: this));
        }
    }

    [Cmdlet(VerbsCommon.Enter, "PSHostSession", DefaultParameterSetName = SubprocessParameterSet)]
    public sealed class EnterPSHostSessionCommand : PSHostSessionCommandBase
    {
        protected override void BeginProcessing()
        {
            Runspace runspace = CreateAndOpenRunspace();
            ProcessSession(runspace);
        }

        protected override void ProcessSession(Runspace runspace)
        {
            // Create PSSession object
            var session = PSSession.Create(
                runspace: runspace,
                transportName: TransportName,
                psCmdlet: this);

            // Enter the session using the built-in Enter-PSSession cmdlet
            using var ps = System.Management.Automation.PowerShell.Create(RunspaceMode.CurrentRunspace);
            ps.AddCommand("Enter-PSSession")
              .AddParameter("Session", session);
            
            try
            {
                ps.Invoke();
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(
                    ex,
                    "EnterPSHostSessionFailed",
                    ErrorCategory.InvalidOperation,
                    session));
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
