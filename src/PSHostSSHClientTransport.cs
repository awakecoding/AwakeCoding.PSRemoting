using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Host;
using System.Management.Automation.Internal;
using System.Management.Automation.Remoting;
using System.Management.Automation.Remoting.Client;
using System.Management.Automation.Runspaces;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Tmds.Ssh;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Connection info for SSH-based PowerShell remoting connections.
    /// Uses Tmds.Ssh for native .NET SSH connectivity to the PowerShell subsystem.
    /// </summary>
    public sealed class SSHClientInfo : RunspaceConnectionInfo
    {
        #region Constants

        private const string DefaultSubsystem = "powershell";
        private const int DefaultSSHPort = 22;
        private const int DefaultConnectingTimeout = -1; // Infinite

        #endregion

        #region Properties

        public override string ComputerName { get; set; }

        /// <summary>
        /// SSH user name (can include domain, e.g., "user@domain")
        /// </summary>
        public string? UserName { get; set; }

        /// <summary>
        /// Path to SSH private key file
        /// </summary>
        public string? KeyFilePath { get; set; }

        /// <summary>
        /// SSH port (default 22)
        /// </summary>
        public int Port { get; set; }

        /// <summary>
        /// SSH subsystem (default "powershell")
        /// </summary>
        public string Subsystem { get; set; }

        /// <summary>
        /// Connection timeout in milliseconds (-1 = infinite)
        /// </summary>
        public int ConnectingTimeout { get; set; }

        /// <summary>
        /// Additional SSH options as hashtable (e.g., @{StrictHostKeyChecking="no"})
        /// </summary>
        public Hashtable? Options { get; set; }

        /// <summary>
        /// Password or credential to use for authentication
        /// </summary>
        public override PSCredential? Credential { get; set; }

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

        /// <summary>
        /// Whether to skip host key verification (not recommended for production)
        /// </summary>
        public bool SkipHostKeyCheck { get; set; }

        /// <summary>
        /// Path to known_hosts file for host key verification.
        /// If not specified, uses default OpenSSH known_hosts locations.
        /// </summary>
        public string? KnownHostsFilePath { get; set; }

        /// <summary>
        /// Whether to automatically add unknown host keys to the known_hosts file
        /// </summary>
        public bool AutoAddHostKey { get; set; }

        /// <summary>
        /// PSHost for interactive prompting (password, host key verification)
        /// </summary>
        internal PSHost? PSHost { get; set; }

        #endregion

        #region Constructors

        public SSHClientInfo(string computerName)
        {
            ComputerName = computerName ?? throw new ArgumentNullException(nameof(computerName));
            Port = DefaultSSHPort;
            Subsystem = DefaultSubsystem;
            ConnectingTimeout = DefaultConnectingTimeout;
        }

        public SSHClientInfo(
            string computerName,
            string? userName,
            string? keyFilePath = null,
            int port = DefaultSSHPort,
            string? subsystem = null,
            int connectingTimeout = DefaultConnectingTimeout,
            Hashtable? options = null,
            string? sshExecutablePath = null) // sshExecutablePath ignored - kept for API compatibility
            : this(computerName)
        {
            UserName = userName;
            KeyFilePath = keyFilePath;
            Port = port > 0 ? port : DefaultSSHPort;
            Subsystem = subsystem ?? DefaultSubsystem;
            ConnectingTimeout = connectingTimeout;
            Options = options;

            // Parse options for known settings
            if (options != null)
            {
                if (options.ContainsKey("StrictHostKeyChecking"))
                {
                    var value = options["StrictHostKeyChecking"]?.ToString();
                    if (value != null && (value.Equals("no", StringComparison.OrdinalIgnoreCase) ||
                                          value.Equals("accept-new", StringComparison.OrdinalIgnoreCase)))
                    {
                        SkipHostKeyCheck = true;
                        AutoAddHostKey = value.Equals("accept-new", StringComparison.OrdinalIgnoreCase);
                    }
                }

                if (options.ContainsKey("UserKnownHostsFile"))
                {
                    KnownHostsFilePath = options["UserKnownHostsFile"]?.ToString();
                }
            }
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Creates Tmds.Ssh client settings for this connection
        /// </summary>
        internal SshClientSettings CreateSshClientSettings()
        {
            var settings = new SshClientSettings(ComputerName)
            {
                Port = Port,
                UserName = UserName ?? Environment.UserName,
                AutoConnect = false,
                AutoReconnect = false
            };

            // Set connection timeout
            if (ConnectingTimeout > 0)
            {
                settings.ConnectTimeout = TimeSpan.FromMilliseconds(ConnectingTimeout);
            }

            // Build credentials list
            var credentials = new List<Credential>();

            // If password credential provided, add it first
            if (Credential != null)
            {
                string password = Credential.GetNetworkCredential().Password;
                credentials.Add(new PasswordCredential(password));
            }
            else
            {
                // No credential provided - add interactive password credential that will prompt
                credentials.Add(new PasswordCredential(() =>
                {
                    return PromptForPassword();
                }));
            }

            // If key file provided, add private key credential
            if (!string.IsNullOrEmpty(KeyFilePath))
            {
                string expandedPath = ExpandKeyFilePath(KeyFilePath);
                if (File.Exists(expandedPath))
                {
                    // Check if key might be encrypted (we'd need a password)
                    if (Credential != null)
                    {
                        string password = Credential.GetNetworkCredential().Password;
                        credentials.Add(new PrivateKeyCredential(expandedPath, password));
                    }
                    else
                    {
                        credentials.Add(new PrivateKeyCredential(expandedPath));
                    }
                }
                else
                {
                    throw new FileNotFoundException($"SSH key file not found: {KeyFilePath}", expandedPath);
                }
            }

            // Add default credentials (SSH agent, default identity files)
            // This allows fallback to standard SSH authentication methods
            if (credentials.Count == 0)
            {
                // Try SSH agent first
                credentials.Add(new SshAgentCredentials());

                // Then try default identity files
                string homeDir = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                string sshDir = Path.Combine(homeDir, ".ssh");
                string[] defaultKeyFiles = { "id_ed25519", "id_ecdsa", "id_rsa", "id_dsa" };

                foreach (var keyFile in defaultKeyFiles)
                {
                    string keyPath = Path.Combine(sshDir, keyFile);
                    if (File.Exists(keyPath))
                    {
                        credentials.Add(new PrivateKeyCredential(keyPath));
                    }
                }
            }

            settings.Credentials = credentials;

            // Configure host key verification
            if (SkipHostKeyCheck)
            {
                // Accept all host keys (not recommended for production)
                settings.HostAuthentication = (KnownHostResult knownHostResult, SshConnectionInfo connectionInfo, CancellationToken cancellationToken) => 
                    ValueTask.FromResult(true);
            }
            else
            {
                // Use custom host authentication with prompting
                if (!string.IsNullOrEmpty(KnownHostsFilePath))
                {
                    settings.UserKnownHostsFilePaths = new List<string> { KnownHostsFilePath };
                }
                
                settings.HostAuthentication = HostAuthenticationCallback;
            }

            return settings;
        }

        public override RunspaceConnectionInfo Clone()
        {
            return new SSHClientInfo(
                ComputerName,
                UserName,
                KeyFilePath,
                Port,
                Subsystem,
                ConnectingTimeout,
                Options,
                null)
            {
                Credential = Credential,
                SkipHostKeyCheck = SkipHostKeyCheck,
                KnownHostsFilePath = KnownHostsFilePath,
                AutoAddHostKey = AutoAddHostKey
            };
        }

        public override BaseClientSessionTransportManager CreateClientSessionTransportManager(
            Guid instanceId,
            string sessionName,
            PSRemotingCryptoHelper cryptoHelper)
        {
            var transportMgr = new SSHClientSessionTransportMgr(
                connectionInfo: this,
                runspaceId: instanceId,
                cryptoHelper: cryptoHelper);
            
            // Set PSHost for interactive prompting if available
            transportMgr.PSHost = this.PSHost;
            
            return transportMgr;
        }

        #endregion

        #region Private Methods

        private static string ExpandKeyFilePath(string keyFilePath)
        {
            if (keyFilePath.StartsWith("~"))
            {
                string home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                return Path.Combine(home, keyFilePath.Substring(1)
                    .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
            }
            return keyFilePath;
        }

        /// <summary>
        /// Prompts for password interactively
        /// </summary>
        private string? PromptForPassword()
        {
            if (PSHost == null)
            {
                // Fallback to Console if PSHost not available
                Console.Write($"{UserName ?? Environment.UserName}@{ComputerName}'s password: ");
                return ReadPasswordFromConsole();
            }

            try
            {
                var prompt = $"{UserName ?? Environment.UserName}@{ComputerName}'s password";
                var credential = PSHost.UI.PromptForCredential(
                    "SSH Authentication",
                    prompt,
                    UserName ?? Environment.UserName,
                    "");
                
                return credential?.GetNetworkCredential().Password;
            }
            catch
            {
                // Fallback to Console prompt
                Console.Write($"{UserName ?? Environment.UserName}@{ComputerName}'s password: ");
                return ReadPasswordFromConsole();
            }
        }

        /// <summary>
        /// Host key verification callback that prompts user for unknown hosts
        /// </summary>
        private ValueTask<bool> HostAuthenticationCallback(
            KnownHostResult knownHostResult, 
            SshConnectionInfo connectionInfo, 
            CancellationToken cancellationToken)
        {
            // If host is trusted, accept immediately
            if (knownHostResult == KnownHostResult.Trusted)
            {
                return ValueTask.FromResult(true);
            }

            // Get the host key information
            string fingerprint = connectionInfo.ServerKey.SHA256FingerPrint ?? "unknown";
            string keyType = connectionInfo.ServerKey.ToString() ?? "unknown";
            
            string message;
            if (knownHostResult == KnownHostResult.Unknown)
            {
                message = $"The authenticity of host '{ComputerName} ({connectionInfo.HostName}:{connectionInfo.Port})' can't be established.\n" +
                         $"Host key fingerprint is {fingerprint}.\n" +
                         "Are you sure you want to continue connecting?";
            }
            else if (knownHostResult == KnownHostResult.Changed || knownHostResult == KnownHostResult.Revoked)
            {
                message = $"@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\n" +
                         $"@    WARNING: REMOTE HOST IDENTIFICATION HAS CHANGED!     @\n" +
                         $"@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@\n" +
                         $"IT IS POSSIBLE THAT SOMEONE IS DOING SOMETHING NASTY!\n" +
                         $"Someone could be eavesdropping on you right now (man-in-the-middle attack)!\n" +
                         $"It is also possible that the host key has just been changed.\n" +
                         $"The fingerprint for the host key sent by the remote host is\n{fingerprint}.\n" +
                         $"Do you want to continue connecting (NOT RECOMMENDED)?";
            }
            else
            {
                // Untrusted - reject by default
                return ValueTask.FromResult(false);
            }

            bool accept = PromptYesNo(message, defaultValue: false);
            
            if (accept && AutoAddHostKey && knownHostResult == KnownHostResult.Unknown)
            {
                // TODO: Add host key to known_hosts file
                // This would require writing to the known_hosts file
            }

            return ValueTask.FromResult(accept);
        }

        /// <summary>
        /// Prompts user for yes/no confirmation
        /// </summary>
        private bool PromptYesNo(string message, bool defaultValue)
        {
            if (PSHost != null)
            {
                try
                {
                    var choices = new System.Collections.ObjectModel.Collection<ChoiceDescription>
                    {
                        new ChoiceDescription("&Yes", "Accept the connection"),
                        new ChoiceDescription("&No", "Reject the connection")
                    };

                    int result = PSHost.UI.PromptForChoice(
                        "SSH Host Key Verification",
                        message,
                        choices,
                        defaultValue ? 0 : 1);

                    return result == 0;
                }
                catch
                {
                    // Fall through to console prompt
                }
            }

            // Fallback to Console prompt
            Console.WriteLine(message);
            Console.Write(defaultValue ? "(yes/no) [yes]: " : "(yes/no) [no]: ");
            string? response = Console.ReadLine()?.Trim().ToLowerInvariant();
            
            if (string.IsNullOrEmpty(response))
                return defaultValue;
            
            return response == "yes" || response == "y";
        }

        /// <summary>
        /// Reads password from console without echoing
        /// </summary>
        private static string? ReadPasswordFromConsole()
        {
            var password = new StringBuilder();
            ConsoleKeyInfo key;

            do
            {
                key = Console.ReadKey(true);

                if (key.Key != ConsoleKey.Enter && key.Key != ConsoleKey.Backspace)
                {
                    password.Append(key.KeyChar);
                }
                else if (key.Key == ConsoleKey.Backspace && password.Length > 0)
                {
                    password.Remove(password.Length - 1, 1);
                }
            }
            while (key.Key != ConsoleKey.Enter);

            Console.WriteLine();
            return password.ToString();
        }

        #endregion
    }

    /// <summary>
    /// Custom TextWriter that wraps RemoteProcess for PSRP message writing.
    /// Provides synchronous Write methods that the PSRP base class requires
    /// by blocking on RemoteProcess's async WriteLineAsync method.
    /// </summary>
    internal sealed class RemoteProcessTextWriter : TextWriter
    {
        private readonly RemoteProcess _process;

        public RemoteProcessTextWriter(RemoteProcess process)
        {
            _process = process ?? throw new ArgumentNullException(nameof(process));
        }

        public override Encoding Encoding => Encoding.UTF8;

        // PSRP uses WriteLine for message sending
        public override void WriteLine(string? value)
        {
            if (value == null) return;
            
            // Block on the async write - PSRP base class expects synchronous completion
            _process.WriteLineAsync(value, CancellationToken.None).AsTask().GetAwaiter().GetResult();
        }

        public override void Write(string? value)
        {
            if (value == null) return;
            
            // Block on the async write
            _process.WriteAsync(value, CancellationToken.None).AsTask().GetAwaiter().GetResult();
        }

        public override void Write(char value)
        {
            _process.WriteAsync(value.ToString(), CancellationToken.None).AsTask().GetAwaiter().GetResult();
        }

        public override void Flush()
        {
            // RemoteProcess auto-flushes, nothing to do
        }
    }

    /// <summary>
    /// SSH transport manager using Tmds.Ssh for native .NET SSH connectivity.
    /// Connects to the PowerShell subsystem and proxies PSRP protocol data.
    /// </summary>
    internal sealed class SSHClientSessionTransportMgr : ClientSessionTransportManagerBase
    {
        #region Data

        private readonly SSHClientInfo _connectionInfo;
        private SshClient? _sshClient;
        private RemoteProcess? _remoteProcess;
        private CancellationTokenSource? _readerCts;
        private volatile bool _isClosed;

        private const string ReaderThreadName = "SSH Transport Reader Thread";

        #endregion

        #region Properties

        /// <summary>
        /// PSHost for interactive prompting
        /// </summary>
        internal PSHost? PSHost { get; set; }

        #endregion

        #region Constructor

        internal SSHClientSessionTransportMgr(
            SSHClientInfo connectionInfo,
            Guid runspaceId,
            PSRemotingCryptoHelper cryptoHelper)
            : base(runspaceId, cryptoHelper)
        {
            _connectionInfo = connectionInfo ?? throw new PSArgumentException(nameof(connectionInfo));
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Create SSH connection using Tmds.Ssh and set up transport reader/writer
        /// </summary>
        public override void CreateAsync()
        {
            // We need to run the async connection synchronously here
            // because the base class expects synchronous completion
            Task.Run(async () =>
            {
                try
                {
                    await CreateConnectionAsync().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    HandleSSHError(new PSRemotingTransportException(
                        $"Failed to establish SSH connection: {ex.Message}", ex));
                }
            }).Wait();
        }

        private async Task CreateConnectionAsync()
        {
            // Create SSH client settings
            var settings = _connectionInfo.CreateSshClientSettings();

            // Create and connect SSH client
            _sshClient = new SshClient(settings);

            // Set up cancellation for timeout
            using var timeoutCts = _connectionInfo.ConnectingTimeout > 0
                ? new CancellationTokenSource(_connectionInfo.ConnectingTimeout)
                : new CancellationTokenSource();

            try
            {
                await _sshClient.ConnectAsync(timeoutCts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (timeoutCts.IsCancellationRequested)
            {
                throw new TimeoutException(
                    $"SSH connection timed out after {_connectionInfo.ConnectingTimeout}ms");
            }

            // Execute the PowerShell subsystem
            _remoteProcess = await _sshClient.ExecuteSubsystemAsync(
                _connectionInfo.Subsystem,
                timeoutCts.Token).ConfigureAwait(false);

            // Create cancellation token for reader
            _readerCts = new CancellationTokenSource();

            // Set up the message writer using our custom TextWriter that wraps RemoteProcess
            // This provides synchronous Write/WriteLine that the PSRP base class requires
            var inputWriter = new RemoteProcessTextWriter(_remoteProcess);
            SetMessageWriter(inputWriter);

            // Start reader thread for PSRP data
            // Note: Reader thread handles both stdout and stderr via ReadLineAsync
            StartReaderThread();
        }

        public override void CloseAsync()
        {
            _isClosed = true;
            
            // Cancel reader thread first to unblock ReadLineAsync
            try
            {
                _readerCts?.Cancel();
            }
            catch { }
            
            // Dispose remote process to close the SSH channel
            // This will cause ReadLineAsync to return null or throw
            try
            {
                _remoteProcess?.Dispose();
            }
            catch { }
            
            base.CloseAsync();
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

            // Cancel reader threads
            try
            {
                _readerCts?.Cancel();
            }
            catch { }

            // Dispose remote process
            try
            {
                _remoteProcess?.Dispose();
            }
            catch { }

            // Dispose SSH client
            try
            {
                _sshClient?.Dispose();
            }
            catch { }

            _readerCts?.Dispose();
            _readerCts = null;
            _remoteProcess = null;
            _sshClient = null;
        }

        #endregion

        #region Private Methods

        private void StartReaderThread()
        {
            var readerThread = new Thread(ProcessReaderThread)
            {
                Name = ReaderThreadName,
                IsBackground = true
            };
            readerThread.Start();
        }

        /// <summary>
        /// Process stdout from SSH subsystem - handle PSRP data
        /// Uses RemoteProcess.ReadLineAsync() to read from the subsystem output.
        /// </summary>
        private void ProcessReaderThread()
        {
            try
            {
                // Send first PSRP fragment
                SendOneItem();

                var cancellationToken = _readerCts?.Token ?? CancellationToken.None;

                // Run async reader loop synchronously in this thread
                Task.Run(async () =>
                {
                    while (!_isClosed && !cancellationToken.IsCancellationRequested)
                    {
                        try
                        {
                            // ReadLineAsync signature: (bool readStdout, bool readStderr, CancellationToken)
                            // Returns: (bool isError, string? line)
                            var (isError, line) = await _remoteProcess!.ReadLineAsync(
                                readStdout: true, 
                                readStderr: true, 
                                cancellationToken).ConfigureAwait(false);

                            if (line == null)
                            {
                                // End of stream - SSH connection closed
                                break;
                            }

                            if (!isError)
                            {
                                // Process PSRP data from stdout
                                HandleOutputDataReceived(line);
                            }
                            else
                            {
                                // Write stderr to console for diagnostics
                                if (line.Length > 0)
                                {
                                    try
                                    {
                                        Console.Error.WriteLine(line);
                                    }
                                    catch (IOException)
                                    {
                                        // Console may not be available
                                    }
                                }
                            }
                        }
                        catch (OperationCanceledException)
                        {
                            break;
                        }
                        catch (IOException)
                        {
                            break;
                        }
                    }
                }).Wait(cancellationToken);
            }
            catch (ObjectDisposedException)
            {
                // Normal reader thread end
            }
            catch (OperationCanceledException)
            {
                // Cancellation requested
            }
            catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
            {
                // Cancellation requested via aggregate
            }
            catch (Exception ex)
            {
                if (!_isClosed)
                {
                    HandleSSHError(new PSRemotingTransportException(
                        $"SSH transport reader thread ended with error: {ex.Message}", ex));
                }
            }
        }

        // Note: Error reading is handled in the reader thread via ReadLineAsync's isError flag
        // No separate error thread needed with Tmds.Ssh RemoteProcess API

        private void HandleSSHError(PSRemotingTransportException psrte)
        {
            RaiseErrorHandler(new TransportErrorOccuredEventArgs(
                psrte, TransportMethodEnum.CloseShellOperationEx));
            CleanupConnection();
        }

        #endregion
    }
}
