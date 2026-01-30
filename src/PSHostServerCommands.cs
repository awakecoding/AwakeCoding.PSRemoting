using System;
using System.Linq;
using System.Management.Automation;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Start-PSHostServer cmdlet - Starts a PowerShell remoting server with specified transport
    /// </summary>
    [Cmdlet(VerbsLifecycle.Start, "PSHostServer")]
    [OutputType(typeof(PSHostTcpServer), typeof(PSHostWebSocketServer), typeof(PSHostNamedPipeServer))]
    public sealed class StartPSHostServerCommand : PSCmdlet
    {
        // Common parameters
        [Parameter(Mandatory = true, Position = 0)]
        [ValidateNotNull()]
        public PSHostTransportType TransportType { get; set; }

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter()]
        [ValidateRange(0, int.MaxValue)]
        public int MaxConnections { get; set; } = 0;

        [Parameter()]
        [ValidateRange(0, int.MaxValue)]
        public int DrainTimeout { get; set; } = 30;

        // TCP and WebSocket parameters
        [Parameter(Position = 1)]
        [ValidateRange(0, 65535)]
        public int? Port { get; set; }

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string ListenAddress { get; set; } = "127.0.0.1";

        // WebSocket-specific parameters
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string Path { get; set; } = "/pwsh";

        [Parameter()]
        public SwitchParameter UseSecureConnection { get; set; }

        // NamedPipe-specific parameters
        [Parameter(Position = 1)]
        [ValidateNotNullOrEmpty()]
        public string? PipeName { get; set; }

        protected override void BeginProcessing()
        {
            PSHostServerBase server;

            try
            {
                switch (TransportType)
                {
                    case PSHostTransportType.TCP:
                        server = CreateTcpServer();
                        break;

                    case PSHostTransportType.WebSocket:
                        server = CreateWebSocketServer();
                        break;

                    case PSHostTransportType.NamedPipe:
                        server = CreateNamedPipeServer();
                        break;

                    default:
                        throw new ArgumentException($"Unsupported transport type: {TransportType}");
                }

                // Register in global registry
                if (!server.Register())
                {
                    throw new InvalidOperationException($"Failed to register server '{server.Name}'");
                }

                try
                {
                    // Start the listener
                    server.StartListenerAsync();

                    // Output the server object
                    WriteObject(server);
                }
                catch
                {
                    // If start fails, unregister the server
                    server.Unregister();
                    throw;
                }
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(
                    ex,
                    "StartPSHostServerFailed",
                    ErrorCategory.InvalidOperation,
                    Name));
            }
        }

        private PSHostTcpServer CreateTcpServer()
        {
            // Validate Port is provided
            if (!Port.HasValue)
            {
                throw new ArgumentException("Port parameter is required for TCP transport");
            }

            int portValue = Port.Value;

            // Generate default name if not provided
            if (string.IsNullOrWhiteSpace(Name))
            {
                Name = portValue == 0 ? $"PSHostTcpServer{Guid.NewGuid().ToString("N").Substring(0, 8)}" : $"PSHostTcpServer{portValue}";
            }

            // Check if server with this name already exists
            var existingServer = PSHostServerBase.GetServer(Name);
            if (existingServer != null)
            {
                throw new InvalidOperationException($"Server with name '{Name}' already exists");
            }

            // Check if another server is already listening on this port
            var serverOnPort = PSHostServerBase.GetServerByPort(portValue);
            if (serverOnPort != null)
            {
                throw new InvalidOperationException($"Server '{serverOnPort.Name}' is already listening on port {portValue}");
            }

            return new PSHostTcpServer(
                name: Name,
                port: portValue,
                listenAddress: ListenAddress,
                maxConnections: MaxConnections,
                drainTimeout: DrainTimeout);
        }

        private PSHostWebSocketServer CreateWebSocketServer()
        {
            // Validate Port is provided
            if (!Port.HasValue)
            {
                throw new ArgumentException("Port parameter is required for WebSocket transport");
            }

            int portValue = Port.Value;

            // Validate port range for WebSocket (must be > 0)
            if (portValue == 0)
            {
                throw new ArgumentException("WebSocket server requires a specific port (cannot be 0)");
            }

            // Generate default name if not provided
            if (string.IsNullOrWhiteSpace(Name))
            {
                Name = $"PSHostWebSocketServer{portValue}";
            }

            // Check if server with this name already exists
            var existingServer = PSHostServerBase.GetServer(Name);
            if (existingServer != null)
            {
                throw new InvalidOperationException($"Server with name '{Name}' already exists");
            }

            // Check if another server is already listening on this port
            var serverOnPort = PSHostServerBase.GetServerByPort(portValue);
            if (serverOnPort != null)
            {
                throw new InvalidOperationException($"Server '{serverOnPort.Name}' is already listening on port {portValue}");
            }

            return new PSHostWebSocketServer(
                name: Name,
                port: portValue,
                listenAddress: ListenAddress,
                path: Path,
                useSecureConnection: UseSecureConnection,
                maxConnections: MaxConnections,
                drainTimeout: DrainTimeout);
        }

        private PSHostNamedPipeServer CreateNamedPipeServer()
        {
            // Generate random pipe name if not provided
            if (string.IsNullOrWhiteSpace(PipeName))
            {
                PipeName = PSHostNamedPipeServer.GenerateRandomPipeName();
            }

            // Generate default server name if not provided
            // Extract enough of the unique portion for a unique name
            if (string.IsNullOrWhiteSpace(Name))
            {
                // PipeName is typically "PSHost_<GUID>", so skip the "PSHost_" prefix
                // and take the first 12 chars of the GUID for uniqueness
                int prefixLen = "PSHost_".Length;
                string uniquePart = PipeName.Length > prefixLen 
                    ? PipeName.Substring(prefixLen, Math.Min(12, PipeName.Length - prefixLen))
                    : PipeName.Substring(0, Math.Min(12, PipeName.Length));
                Name = $"PSHostNamedPipeServer{uniquePart}";
            }

            // Check if server with this name already exists
            var existingServer = PSHostServerBase.GetServer(Name);
            if (existingServer != null)
            {
                throw new InvalidOperationException($"Server with name '{Name}' already exists");
            }

            // Check if another server is already listening on this pipe name
            var serverOnPipe = PSHostServerBase.GetServerByPipeName(PipeName);
            if (serverOnPipe != null)
            {
                throw new InvalidOperationException($"Server '{serverOnPipe.Name}' is already listening on pipe '{PipeName}'");
            }

            return new PSHostNamedPipeServer(
                name: Name,
                pipeName: PipeName,
                maxConnections: MaxConnections,
                drainTimeout: DrainTimeout);
        }
    }

    /// <summary>
    /// Stop-PSHostServer cmdlet - Stops a PowerShell remoting server
    /// </summary>
    [Cmdlet(VerbsLifecycle.Stop, "PSHostServer", DefaultParameterSetName = "ByName")]
    public sealed class StopPSHostServerCommand : PSCmdlet
    {
        [Parameter(ParameterSetName = "ByName", Position = 0, Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter(ParameterSetName = "ByPort", Mandatory = true)]
        [ValidateRange(1, 65535)]
        public int Port { get; set; }

        [Parameter(ParameterSetName = "ByPipeName", Mandatory = true)]
        [ValidateNotNullOrEmpty()]
        public string? PipeName { get; set; }

        [Parameter(ParameterSetName = "ByServer", Mandatory = true, ValueFromPipeline = true)]
        [ValidateNotNull()]
        public PSHostServerBase? Server { get; set; }

        [Parameter()]
        public SwitchParameter Force { get; set; }

        protected override void ProcessRecord()
        {
            PSHostServerBase? server = null;

            try
            {
                // Find the server based on parameter set
                switch (ParameterSetName)
                {
                    case "ByName":
                        server = PSHostServerBase.GetServer(Name!);
                        if (server == null)
                        {
                            WriteError(new ErrorRecord(
                                new ItemNotFoundException($"No server found with name '{Name}'"),
                                "ServerNotFound",
                                ErrorCategory.ObjectNotFound,
                                Name));
                            return;
                        }
                        break;

                    case "ByPort":
                        server = PSHostServerBase.GetServerByPort(Port);
                        if (server == null)
                        {
                            WriteError(new ErrorRecord(
                                new ItemNotFoundException($"No server found listening on port {Port}"),
                                "ServerNotFound",
                                ErrorCategory.ObjectNotFound,
                                Port));
                            return;
                        }
                        break;

                    case "ByPipeName":
                        server = PSHostServerBase.GetServerByPipeName(PipeName!);
                        if (server == null)
                        {
                            WriteError(new ErrorRecord(
                                new ItemNotFoundException($"No server found with pipe name '{PipeName}'"),
                                "ServerNotFound",
                                ErrorCategory.ObjectNotFound,
                                PipeName));
                            return;
                        }
                        break;

                    case "ByServer":
                        server = Server;
                        break;
                }

                if (server != null)
                {
                    server.StopListenerAsync(force: Force);
                }
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(
                    ex,
                    "StopPSHostServerFailed",
                    ErrorCategory.InvalidOperation,
                    server?.Name ?? Name));
            }
        }
    }

    /// <summary>
    /// Get-PSHostServer cmdlet - Retrieves PowerShell remoting servers
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "PSHostServer", DefaultParameterSetName = "All")]
    [OutputType(typeof(PSHostTcpServer), typeof(PSHostWebSocketServer), typeof(PSHostNamedPipeServer))]
    public sealed class GetPSHostServerCommand : PSCmdlet
    {
        [Parameter(ParameterSetName = "ByName", Position = 0, Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter(ParameterSetName = "ByPort", Mandatory = true)]
        [ValidateRange(1, 65535)]
        public int Port { get; set; }

        [Parameter(ParameterSetName = "ByPipeName", Mandatory = true)]
        [ValidateNotNullOrEmpty()]
        public string? PipeName { get; set; }

        [Parameter(ParameterSetName = "All")]
        [ValidateNotNull()]
        public PSHostTransportType? TransportType { get; set; }

        protected override void ProcessRecord()
        {
            try
            {
                switch (ParameterSetName)
                {
                    case "ByName":
                        var server = PSHostServerBase.GetServer(Name!);
                        if (server == null)
                        {
                            WriteError(new ErrorRecord(
                                new ItemNotFoundException($"No server found with name '{Name}'"),
                                "ServerNotFound",
                                ErrorCategory.ObjectNotFound,
                                Name));
                            return;
                        }
                        WriteObject(server);
                        break;

                    case "ByPort":
                        var serverByPort = PSHostServerBase.GetServerByPort(Port);
                        if (serverByPort == null)
                        {
                            WriteError(new ErrorRecord(
                                new ItemNotFoundException($"No server found listening on port {Port}"),
                                "ServerNotFound",
                                ErrorCategory.ObjectNotFound,
                                Port));
                            return;
                        }
                        WriteObject(serverByPort);
                        break;

                    case "ByPipeName":
                        var serverByPipe = PSHostServerBase.GetServerByPipeName(PipeName!);
                        if (serverByPipe == null)
                        {
                            WriteError(new ErrorRecord(
                                new ItemNotFoundException($"No server found with pipe name '{PipeName}'"),
                                "ServerNotFound",
                                ErrorCategory.ObjectNotFound,
                                PipeName));
                            return;
                        }
                        WriteObject(serverByPipe);
                        break;

                    case "All":
                        var allServers = PSHostServerBase.GetAllServers();
                        
                        // Filter by transport type if specified
                        if (TransportType.HasValue)
                        {
                            allServers = TransportType.Value switch
                            {
                                PSHostTransportType.TCP => allServers.OfType<PSHostTcpServer>(),
                                PSHostTransportType.WebSocket => allServers.OfType<PSHostWebSocketServer>(),
                                PSHostTransportType.NamedPipe => allServers.OfType<PSHostNamedPipeServer>(),
                                _ => allServers
                            };
                        }

                        WriteObject(allServers, enumerateCollection: true);
                        break;
                }
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(
                    ex,
                    "GetPSHostServerFailed",
                    ErrorCategory.InvalidOperation,
                    Name));
            }
        }
    }
}
