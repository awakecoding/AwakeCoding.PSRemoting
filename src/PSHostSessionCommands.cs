using System;
using System.Collections;
using System.Collections.Generic;
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

    public sealed class PSHostProcessEnvironmentUpdateResult
    {
        public int ProcessId { get; set; }

        public string ProcessName { get; set; } = string.Empty;

        public string Status { get; set; } = string.Empty;

        public string Mode { get; set; } = string.Empty;

        public bool? WasUpToDate { get; set; }

        public int? DriftCount { get; set; }

        public int? AppliedCount { get; set; }

        public string? PSVersion { get; set; }

        public string? Reason { get; set; }
    }

    [Cmdlet(VerbsData.Update, "PSHostProcessEnvironment", DefaultParameterSetName = AllProcessesParameterSet, SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Low)]
    [OutputType(typeof(PSHostProcessEnvironmentUpdateResult))]
    public sealed class UpdatePSHostProcessEnvironmentCommand : PSCmdlet
    {
        private const string AllProcessesParameterSet = "AllProcesses";
        private const string ProcessIdParameterSet = "ProcessId";

        private const string RefreshEnvironmentScript = @"
$machine = [Environment]::GetEnvironmentVariables('Machine')
$user = [Environment]::GetEnvironmentVariables('User')

$desired = @{}
foreach ($k in $machine.Keys) { $desired[[string]$k] = [string]$machine[$k] }
foreach ($k in $user.Keys) { $desired[[string]$k] = [string]$user[$k] }

$mPath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
$uPath = [Environment]::GetEnvironmentVariable('Path', 'User')
$desiredPath = (($mPath, $uPath) -join ';').Trim(';')
$desired['Path'] = $desiredPath

$driftIgnore = @('PSModulePath', 'PATHEXT')

$driftCount = 0
foreach ($k in $desired.Keys) {
    if ($driftIgnore -contains [string]$k) {
        continue
    }

    $current = [Environment]::GetEnvironmentVariable([string]$k, 'Process')
    $expected = [string]$desired[$k]

    if ($null -eq $current) { $current = '' }
    if ($null -eq $expected) { $expected = '' }

    if (-not [string]::Equals($current, $expected, [System.StringComparison]::Ordinal)) {
        $driftCount++
    }
}

$wasUpToDate = ($driftCount -eq 0)

foreach ($k in $machine.Keys) {
    [System.Environment]::SetEnvironmentVariable([string]$k, [string]$machine[$k], 'Process')
}

foreach ($k in $user.Keys) {
    [System.Environment]::SetEnvironmentVariable([string]$k, [string]$user[$k], 'Process')
}

[System.Environment]::SetEnvironmentVariable('Path', $desiredPath, 'Process')

[pscustomobject]@{
    WasUpToDate = $wasUpToDate
    DriftCount = $driftCount
    PSVersion = [string]$PSVersionTable.PSVersion
}
";

        private const string ExplicitEnvironmentScript = @"
param([hashtable]$variables)

$appliedCount = 0

if ($null -ne $variables) {
    foreach ($entry in $variables.GetEnumerator()) {
        $key = [string]$entry.Key
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        if ($null -eq $entry.Value) {
            [System.Environment]::SetEnvironmentVariable($key, $null, 'Process')
        }
        else {
            [System.Environment]::SetEnvironmentVariable($key, [string]$entry.Value, 'Process')
        }

        $appliedCount++
    }
}

[pscustomobject]@{
    AppliedCount = $appliedCount
    PSVersion = [string]$PSVersionTable.PSVersion
}
";

        [Parameter()]
        [ValidateNotNull()]
        public Hashtable? Environment { get; set; }

        [Parameter(ParameterSetName = ProcessIdParameterSet, Position = 0, Mandatory = true, ValueFromPipelineByPropertyName = true)]
        [Alias("Id", "ProcessId")]
        [ValidateNotNull()]
        public int[]? TargetProcessId { get; set; }

        [Parameter()]
        [ValidateRange(1000, 300000)]
        public int OpenTimeout { get; set; } = 5000;

        [Parameter()]
        [ValidateRange(500, 60000)]
        public int ReadinessTimeout { get; set; } = 5000;

        protected override void BeginProcessing()
        {
            foreach (var target in ResolveTargets())
            {
                ProcessTarget(target);
            }
        }

        private IEnumerable<(int Id, string Name)> ResolveTargets()
        {
            var seen = new HashSet<int>();
            int currentPid = System.Environment.ProcessId;

            if (ParameterSetName == ProcessIdParameterSet && TargetProcessId != null)
            {
                foreach (int processId in TargetProcessId)
                {
                    if (processId == currentPid || !seen.Add(processId))
                    {
                        continue;
                    }

                    string processName = "Unknown";
                    try
                    {
                        using var process = Process.GetProcessById(processId);
                        processName = process.ProcessName;
                    }
                    catch
                    {
                    }

                    yield return (processId, processName);
                }

                yield break;
            }

            foreach (string processName in new[] { "pwsh", "powershell" })
            {
                Process[] processes;

                try
                {
                    processes = Process.GetProcessesByName(processName);
                }
                catch
                {
                    continue;
                }

                foreach (var process in processes)
                {
                    using (process)
                    {
                        if (process.Id == currentPid || !seen.Add(process.Id))
                        {
                            continue;
                        }

                        yield return (process.Id, process.ProcessName);
                    }
                }
            }
        }

        private void ProcessTarget((int Id, string Name) target)
        {
            bool useExplicitEnvironment = Environment != null;
            string mode = useExplicitEnvironment ? "Explicit" : "Refresh";
            var targetDescription = $"PID={target.Id} Name={target.Name}";
            string action = useExplicitEnvironment
                ? "Set explicit process environment variables"
                : "Refresh process environment from registry";

            if (!ShouldProcess(targetDescription, action))
            {
                WriteObject(new PSHostProcessEnvironmentUpdateResult
                {
                    ProcessId = target.Id,
                    ProcessName = target.Name,
                    Status = "Skipped",
                    Mode = mode,
                    Reason = "Skipped by ShouldProcess"
                });
                return;
            }

            PSSession? session = null;

            try
            {
                session = ConnectSession(target.Id);
                var operationResult = useExplicitEnvironment
                    ? InvokeExplicitEnvironment(session, Environment!)
                    : InvokeRefresh(session);

                string? psVersion = GetStringProperty(operationResult, "PSVersion");

                if (useExplicitEnvironment)
                {
                    int appliedCount = GetInt32Property(operationResult, "AppliedCount");

                    WriteObject(new PSHostProcessEnvironmentUpdateResult
                    {
                        ProcessId = target.Id,
                        ProcessName = target.Name,
                        Status = "Applied",
                        Mode = mode,
                        AppliedCount = appliedCount,
                        PSVersion = psVersion
                    });
                }
                else
                {
                    bool wasUpToDate = GetBooleanProperty(operationResult, "WasUpToDate");
                    int driftCount = GetInt32Property(operationResult, "DriftCount");

                    WriteObject(new PSHostProcessEnvironmentUpdateResult
                    {
                        ProcessId = target.Id,
                        ProcessName = target.Name,
                        Status = wasUpToDate ? "AlreadyCurrent" : "Updated",
                        Mode = mode,
                        WasUpToDate = wasUpToDate,
                        DriftCount = driftCount,
                        PSVersion = psVersion
                    });
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message;
                bool isSkippable = IsSkippableError(message);

                WriteObject(new PSHostProcessEnvironmentUpdateResult
                {
                    ProcessId = target.Id,
                    ProcessName = target.Name,
                    Status = isSkippable ? "Skipped" : "Failed",
                    Mode = mode,
                    Reason = message
                });

                if (!isSkippable)
                {
                    WriteWarning($"Failed to refresh environment for {targetDescription}: {message}");
                }
            }
            finally
            {
                if (session != null)
                {
                    RemoveSession(session);
                }
            }
        }

        private PSSession ConnectSession(int processId)
        {
            using var ps = System.Management.Automation.PowerShell.Create(RunspaceMode.CurrentRunspace);
            ps.AddCommand("Connect-PSHostProcess")
              .AddParameter("Id", processId)
              .AddParameter("OpenTimeout", OpenTimeout)
              .AddParameter("ReadinessTimeout", ReadinessTimeout)
              .AddParameter("ErrorAction", ActionPreference.Stop);

            var sessions = ps.Invoke<PSSession>();
            ThrowIfPipelineErrors(ps, $"Connect-PSHostProcess PID={processId}");

            if (sessions.Count == 0 || sessions[0] == null)
            {
                throw new InvalidOperationException($"Connect-PSHostProcess returned no session for PID {processId}");
            }

            return sessions[0];
        }

        private PSObject? InvokeRefresh(PSSession session)
        {
            using var ps = System.Management.Automation.PowerShell.Create(RunspaceMode.CurrentRunspace);
            ps.AddCommand("Invoke-Command")
              .AddParameter("Session", session)
              .AddParameter("ScriptBlock", ScriptBlock.Create(RefreshEnvironmentScript))
              .AddParameter("ErrorAction", ActionPreference.Stop);

            var results = ps.Invoke();
            ThrowIfPipelineErrors(ps, "Invoke-Command refresh script");

            if (results.Count == 0)
            {
                return null;
            }

            return results[0];
        }

        private PSObject? InvokeExplicitEnvironment(PSSession session, Hashtable variables)
        {
            using var ps = System.Management.Automation.PowerShell.Create(RunspaceMode.CurrentRunspace);
            ps.AddCommand("Invoke-Command")
              .AddParameter("Session", session)
              .AddParameter("ScriptBlock", ScriptBlock.Create(ExplicitEnvironmentScript))
              .AddParameter("ArgumentList", new object[] { variables })
              .AddParameter("ErrorAction", ActionPreference.Stop);

            var results = ps.Invoke();
            ThrowIfPipelineErrors(ps, "Invoke-Command explicit environment script");

            if (results.Count == 0)
            {
                return null;
            }

            return results[0];
        }

        private void RemoveSession(PSSession session)
        {
            try
            {
                using var ps = System.Management.Automation.PowerShell.Create(RunspaceMode.CurrentRunspace);
                ps.AddCommand("Remove-PSSession")
                  .AddParameter("Session", session)
                  .AddParameter("ErrorAction", ActionPreference.Stop);
                ps.Invoke();
            }
            catch
            {
            }
        }

        private static bool GetBooleanProperty(PSObject? obj, string propertyName)
        {
            var value = obj?.Properties[propertyName]?.Value;
            if (value is bool booleanValue)
            {
                return booleanValue;
            }

            if (value != null && bool.TryParse(value.ToString(), out bool parsedBoolean))
            {
                return parsedBoolean;
            }

            return false;
        }

        private static int GetInt32Property(PSObject? obj, string propertyName)
        {
            var value = obj?.Properties[propertyName]?.Value;
            if (value is int intValue)
            {
                return intValue;
            }

            if (value != null && int.TryParse(value.ToString(), out int parsedInt))
            {
                return parsedInt;
            }

            return 0;
        }

        private static string? GetStringProperty(PSObject? obj, string propertyName)
        {
            return obj?.Properties[propertyName]?.Value?.ToString();
        }

        private static bool IsSkippableError(string? message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return false;
            }

            return message.Contains("Access is denied", StringComparison.OrdinalIgnoreCase)
                || message.Contains("has exited", StringComparison.OrdinalIgnoreCase)
                || message.Contains("No process is associated with this object", StringComparison.OrdinalIgnoreCase)
                || message.Contains("Cannot find a process with the process identifier", StringComparison.OrdinalIgnoreCase)
                || message.Contains("The system cannot find the file specified", StringComparison.OrdinalIgnoreCase)
                || message.Contains("pipe", StringComparison.OrdinalIgnoreCase);
        }

        private static void ThrowIfPipelineErrors(System.Management.Automation.PowerShell ps, string operation)
        {
            if (ps.Streams.Error.Count == 0)
            {
                return;
            }

            var error = ps.Streams.Error[0];
            throw new InvalidOperationException(
                $"{operation} failed: {error.Exception?.Message ?? error.ToString()}",
                error.Exception);
        }
    }
}
