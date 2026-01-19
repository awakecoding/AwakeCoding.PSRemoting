using System;
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
    [Cmdlet(VerbsCommon.New, "PSHostSession")]
    [OutputType(typeof(PSSession))]
    public sealed class NewPSHostSessionCommand : PSCmdlet
    {
        private PSHostClientInfo? _connectionInfo;
        private Runspace? _runspace;
        private ManualResetEvent? _openAsync;

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string? ExecutablePath { get; set; }

        [Parameter()]
        public SwitchParameter UseWindowsPowerShell { get; set; }

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string RunspaceName { get; set; } = "PSHostClient";

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string TransportName { get; set; } = "PSHostSession";

        protected override void BeginProcessing()
        {
            string computerName = "localhost";
            string arguments = "-NoLogo -NoProfile -s";

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

            _connectionInfo = new PSHostClientInfo(computerName, executable, arguments);

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
                _openAsync.WaitOne();

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
                // If timeout, the session will be returned in Broken state
                _openAsync.WaitOne(_connectionInfo.OpenTimeout);

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
