using System;
using System.Linq;
using System.Management.Automation;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Start-PSHostWebSocketServer cmdlet - Starts a WebSocket-based PowerShell remoting server
    /// </summary>
    [Cmdlet(VerbsLifecycle.Start, "PSHostWebSocketServer")]
    [OutputType(typeof(PSHostWebSocketServer))]
    public sealed class StartPSHostWebSocketServerCommand : PSCmdlet
    {
        [Parameter(Position = 0, Mandatory = true)]
        [ValidateRange(1, 65535)]
        public int Port { get; set; }

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string ListenAddress { get; set; } = "127.0.0.1";

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string Path { get; set; } = "/pwsh";

        [Parameter()]
        [ValidateRange(0, int.MaxValue)]
        public int MaxConnections { get; set; } = 0;

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter()]
        [ValidateRange(0, int.MaxValue)]
        public int DrainTimeout { get; set; } = 30;

        [Parameter()]
        public SwitchParameter UseSecureConnection { get; set; }

        protected override void BeginProcessing()
        {
            // Generate default name if not provided
            if (string.IsNullOrWhiteSpace(Name))
            {
                Name = $"PSHostWebSocketServer{Port}";
            }

            // Check if server with this name already exists
            var existingServer = PSHostServerBase.GetServer(Name);
            if (existingServer != null)
            {
                throw new InvalidOperationException($"Server with name '{Name}' already exists");
            }

            // Check if another server is already listening on this port
            var serverOnPort = PSHostServerBase.GetServerByPort(Port);
            if (serverOnPort != null)
            {
                throw new InvalidOperationException($"Server '{serverOnPort.Name}' is already listening on port {Port}");
            }

            try
            {
                // Create and start the WebSocket server
                var server = new PSHostWebSocketServer(
                    name: Name,
                    port: Port,
                    listenAddress: ListenAddress,
                    path: Path,
                    maxConnections: MaxConnections,
                    drainTimeout: DrainTimeout,
                    useSecureConnection: UseSecureConnection);

                // Register in global registry
                if (!server.Register())
                {
                    throw new InvalidOperationException($"Failed to register server '{Name}'");
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
                    "StartPSHostWebSocketServerFailed",
                    ErrorCategory.InvalidOperation,
                    Name));
            }
        }
    }

    /// <summary>
    /// Stop-PSHostWebSocketServer cmdlet - Stops a WebSocket-based PowerShell remoting server
    /// </summary>
    [Cmdlet(VerbsLifecycle.Stop, "PSHostWebSocketServer", DefaultParameterSetName = "ByName")]
    public sealed class StopPSHostWebSocketServerCommand : PSCmdlet
    {
        [Parameter(ParameterSetName = "ByName", Position = 0, Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter(ParameterSetName = "ByPort", Mandatory = true)]
        [ValidateRange(1, 65535)]
        public int Port { get; set; }

        [Parameter(ParameterSetName = "ByServer", Position = 0, Mandatory = true, ValueFromPipeline = true)]
        [ValidateNotNull()]
        public PSHostServerBase? Server { get; set; }

        [Parameter()]
        public SwitchParameter Force { get; set; }

        protected override void ProcessRecord()
        {
            PSHostServerBase? server = null;

            // Determine which server to stop based on parameter set
            if (ParameterSetName == "ByServer")
            {
                server = Server;
            }
            else if (ParameterSetName == "ByName")
            {
                server = PSHostServerBase.GetServer(Name!);
                if (server == null)
                {
                    WriteError(new ErrorRecord(
                        new InvalidOperationException($"Server '{Name}' not found"),
                        "ServerNotFound",
                        ErrorCategory.ObjectNotFound,
                        Name));
                    return;
                }
            }
            else if (ParameterSetName == "ByPort")
            {
                server = PSHostServerBase.GetServerByPort(Port);
                if (server == null)
                {
                    WriteError(new ErrorRecord(
                        new InvalidOperationException($"No server found listening on port {Port}"),
                        "ServerNotFound",
                        ErrorCategory.ObjectNotFound,
                        Port));
                    return;
                }
            }

            if (server == null)
            {
                return;
            }

            try
            {
                // Stop the server
                server.StopListenerAsync(Force);

                // Unregister from global registry
                server.Unregister();

                WriteVerbose($"Server '{server.Name}' stopped successfully");
            }
            catch (Exception ex)
            {
                WriteError(new ErrorRecord(
                    ex,
                    "StopPSHostWebSocketServerFailed",
                    ErrorCategory.InvalidOperation,
                    server.Name));
            }
        }
    }

    /// <summary>
    /// Get-PSHostWebSocketServer cmdlet - Gets WebSocket-based PowerShell remoting servers
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "PSHostWebSocketServer", DefaultParameterSetName = "All")]
    [OutputType(typeof(PSHostServerBase))]
    public sealed class GetPSHostWebSocketServerCommand : PSCmdlet
    {
        [Parameter(ParameterSetName = "ByName", Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter(ParameterSetName = "ByPort")]
        [ValidateRange(1, 65535)]
        public int Port { get; set; }

        [Parameter(ParameterSetName = "All")]
        public SwitchParameter All { get; set; }

        protected override void ProcessRecord()
        {
            if (ParameterSetName == "ByName")
            {
                var server = PSHostServerBase.GetServer(Name!);
                if (server != null && server is PSHostWebSocketServer)
                {
                    WriteObject(server);
                }
                else if (server == null)
                {
                    WriteError(new ErrorRecord(
                        new InvalidOperationException($"Server '{Name}' not found"),
                        "ServerNotFound",
                        ErrorCategory.ObjectNotFound,
                        Name));
                }
            }
            else if (ParameterSetName == "ByPort")
            {
                var server = PSHostServerBase.GetServerByPort(Port);
                if (server != null && server is PSHostWebSocketServer)
                {
                    WriteObject(server);
                }
                else if (server == null)
                {
                    WriteError(new ErrorRecord(
                        new InvalidOperationException($"No server found listening on port {Port}"),
                        "ServerNotFound",
                        ErrorCategory.ObjectNotFound,
                        Port));
                }
            }
            else // All
            {
                var servers = PSHostServerBase.GetAllServers()
                    .Where(s => s is PSHostWebSocketServer)
                    .ToArray();
                if (servers.Length > 0)
                {
                    WriteObject(servers, enumerateCollection: true);
                }
            }
        }
    }
}
