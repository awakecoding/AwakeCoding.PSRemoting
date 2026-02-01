#if false
// Deprecated: types moved to PSHostSessionCommands.cs and Transports/*
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Pipes;
using System.Management.Automation;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Runtime.InteropServices;
using System.Threading;

namespace AwakeCoding.PSRemoting.PowerShell
{
    internal sealed class PSHostClientInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        public string Executable { get; set; }

        public string Arguments { get; set; }

        public override PSCredential? Credential
        {
            get { return null; }
            set { throw new NotImplementedException(); }
        }

        public override AuthenticationMechanism AuthenticationMechanism
        {
            get { return AuthenticationMechanism.Default; }
            set { throw new NotImplementedException(); }
        }

        public override string CertificateThumbprint
        {
            get { return string.Empty; }
            set { throw new NotImplementedException(); }
        }

        public PSHostClientInfo(string computerName, string executable, string arguments)
        {
            ComputerName = computerName;
            Executable = executable;
            Arguments = arguments;
        }

        public override RunspaceConnectionInfo Clone()
        {
            var connectionInfo = new PSHostClientInfo(ComputerName, Executable, Arguments);
            return connectionInfo;
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new PSHostClientSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    internal sealed class PSHostClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly PSHostClientInfo _connectionInfo;
        private Process? _process = null;
        private volatile bool _isClosed = false;

        internal PSHostClientSessionTransportMgr(
            PSHostClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
        }

        public override void CreateAsync()
        {
            _process = new Process();
            _process.StartInfo.FileName = _connectionInfo.Executable;
            _process.StartInfo.Arguments = _connectionInfo.Arguments;
            _process.StartInfo.RedirectStandardInput = true;
            _process.StartInfo.RedirectStandardOutput = true;
            _process.StartInfo.RedirectStandardError = true;
            _process.StartInfo.UseShellExecute = false;
            _process.StartInfo.CreateNoWindow = true;

            _process.ErrorDataReceived += ErrorDataReceived;
            _process.OutputDataReceived += OutputDataReceived;
            _process.Start();
            _process.BeginOutputReadLine();
            _process.BeginErrorReadLine();

            SetMessageWriter(_process.StandardInput);
            SendOneItem();
        }

        public override void CloseAsync()
        {
            // Mark as closed to prevent event handler race conditions
            _isClosed = true;

            // Cancel async read operations first
            if (_process != null)
            {
                try
                {
                    _process.CancelOutputRead();
                }
                catch { }

                try
                {
                    _process.CancelErrorRead();
                }
                catch { }
            }

            // Call base implementation
            base.CloseAsync();
        }

        private void ErrorDataReceived(object? sender, DataReceivedEventArgs args)
        {
            // Guard against race conditions during close
            if (_isClosed) return;
            HandleErrorDataReceived(args.Data);
        }

        private void OutputDataReceived(object? sender, DataReceivedEventArgs args)
        {
            // Guard against race conditions during close
            if (_isClosed) return;
            HandleOutputDataReceived(args.Data);
        }

        protected override void Dispose(bool isDisposing)
        {
            if (isDisposing)
            {
                CleanupConnection();
            }

            base.Dispose(isDisposing);
        }

        protected override void CleanupConnection()
        {
            _isClosed = true;

            if (_process != null)
            {
                try
                {
                    // Unsubscribe event handlers first to prevent race conditions
                    _process.ErrorDataReceived -= ErrorDataReceived;
                    _process.OutputDataReceived -= OutputDataReceived;
                }
                catch { }

                try
                {
                    // Cancel async reads if still active
                    _process.CancelOutputRead();
                }
                catch { }

                try
                {
                    _process.CancelErrorRead();
                }
                catch { }

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
                    _process?.Dispose();
                }
            }
            _process = null;
        }
    }

    internal sealed class PSHostNamedPipeInfo : RunspaceConnectionInfo
    {
        public override string ComputerName { get; set; }

        public string PipeName { get; set; }

        public string AppDomainName { get; set; }

        public new int OpenTimeout { get; set; } = 5000; // 5 second default timeout

        public override PSCredential? Credential
        {
            get { return null; }
            set { throw new NotImplementedException(); }
        }

        public override AuthenticationMechanism AuthenticationMechanism
        {
            get { return AuthenticationMechanism.Default; }
            set { throw new NotImplementedException(); }
        }

        public override string CertificateThumbprint
        {
            get { return string.Empty; }
            set { throw new NotImplementedException(); }
        }

        public PSHostNamedPipeInfo(string computerName, string pipeName, string appDomainName = "DefaultAppDomain")
        {
            ComputerName = computerName;
            PipeName = pipeName;
            AppDomainName = appDomainName;
        }

        public override RunspaceConnectionInfo Clone()
        {
            var connectionInfo = new PSHostNamedPipeInfo(ComputerName, PipeName, AppDomainName);
            return connectionInfo;
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            return new PSHostNamedPipeSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
        }
    }

    internal sealed class PSHostNamedPipeSessionTransportMgr : ClientSessionTransportManagerBase
    {
        private readonly PSHostNamedPipeInfo _connectionInfo;
        private NamedPipeClientStream? _pipeStream = null;
        private StreamWriter? _streamWriter = null;
        private StreamReader? _streamReader = null;
        private CancellationTokenSource? _readerCts = null;
        private const string _threadName = "PSHostNamedPipe Reader Thread";

        internal PSHostNamedPipeSessionTransportMgr(
            PSHostNamedPipeInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            if (connectionInfo == null) { throw new PSArgumentException("connectionInfo"); }
            _connectionInfo = connectionInfo;
        }

        public override void CreateAsync()
        {
            // Create a client stream to the local server using the pipe name without prefix
            // Using duplex direction for bidirectional communication
            _pipeStream = new NamedPipeClientStream(
                ".", // local machine
                _connectionInfo.PipeName,
                PipeDirection.InOut,
                PipeOptions.Asynchronous);

            // Use ConnectAsync with CancellationToken for reliable timeout on all platforms
            // On Linux, the synchronous Connect() timeout parameter doesn't work reliably
            using var cts = new CancellationTokenSource(_connectionInfo.OpenTimeout);
            try
            {
                // ConnectAsync respects the CancellationToken properly
                var connectTask = _pipeStream.ConnectAsync(cts.Token);
                if (!connectTask.Wait(_connectionInfo.OpenTimeout))
                {
                    cts.Cancel();
                    throw new TimeoutException($"Named pipe connection timed out after {_connectionInfo.OpenTimeout}ms");
                }
            }
            catch (OperationCanceledException)
            {
                throw new TimeoutException($"Named pipe connection timed out after {_connectionInfo.OpenTimeout}ms");
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
            {
                throw new TimeoutException($"Named pipe connection timed out after {_connectionInfo.OpenTimeout}ms");
            }

            // Create text reader/writer for line-based PSRP protocol
            _streamReader = new StreamReader(_pipeStream, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
            _streamWriter = new StreamWriter(_pipeStream, System.Text.Encoding.UTF8, 4096, leaveOpen: true)
            {
                AutoFlush = true
            };

            // Set up the message writer for outgoing data
            SetMessageWriter(_streamWriter);

            // Create cancellation token for reader thread
            _readerCts = new CancellationTokenSource();

            // Start the reader thread (pattern matches NamedPipeClientSessionTransportManagerBase)
            StartReaderThread();
        }

        public override void CloseAsync()
        {
            // Cancel the reader thread first before attempting to send close packet
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Dispose the pipe stream to unblock any pending reads
            try
            {
                _pipeStream?.Dispose();
            }
            catch { }

            // Call base implementation
            base.CloseAsync();
        }

        private void StartReaderThread()
        {
            Thread readerThread = new Thread(ProcessReaderThread);
            readerThread.Name = _threadName;
            readerThread.IsBackground = true;
            readerThread.Start();
        }

        private void ProcessReaderThread()
        {
            try
            {
                // Send one fragment - must be done inside reader thread
                SendOneItem();

                // Start reader loop with cancellation support
                while (_readerCts != null && !_readerCts.IsCancellationRequested)
                {
                    // Use async read with cancellation token for proper timeout handling
                    var readTask = _streamReader?.ReadLineAsync(_readerCts.Token);
                    if (readTask == null)
                    {
                        break;
                    }

                    string? data;
                    try
                    {
                        data = readTask.Value.AsTask().GetAwaiter().GetResult();
                    }
                    catch (OperationCanceledException)
                    {
                        // Cancellation requested - exit gracefully
                        break;
                    }

                    if (data == null)
                    {
                        // End of stream - pipe closed or error
                        break;
                    }

                    // Process the received data (handles both normal and error messages)
                    HandleDataReceived(data);
                }
            }
            catch (ObjectDisposedException)
            {
                // Normal reader thread end
            }
            catch (IOException)
            {
                // Pipe closed or disconnected
            }
            catch (OperationCanceledException)
            {
                // Cancellation requested
            }
        }

        protected override void Dispose(bool isDisposing)
        {
            // Cleanup first to cancel pending reads before base class tries to wait
            if (isDisposing)
            {
                CleanupConnection();
            }

            base.Dispose(isDisposing);
        }

        protected override void CleanupConnection()
        {
            // Cancel the reader thread first
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Dispose the pipe stream to unblock any pending reads
            // This must be done before disposing readers/writers
            try
            {
                _pipeStream?.Dispose();
            }
            catch { }

            try
            {
                _streamReader?.Dispose();
            }
            catch { }

            try
            {
                _streamWriter?.Dispose();
            }
            catch { }

            _readerCts = null;
            _streamReader = null;
            _streamWriter = null;
            _pipeStream = null;
        }
    }

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
#endif
