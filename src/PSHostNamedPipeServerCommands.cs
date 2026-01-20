using System;
using System.Linq;
using System.Management.Automation;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Start-PSHostNamedPipeServer cmdlet - Starts a named pipe-based PowerShell remoting server
    /// </summary>
    [Cmdlet(VerbsLifecycle.Start, "PSHostNamedPipeServer")]
    [OutputType(typeof(PSHostNamedPipeServer))]
    public sealed class StartPSHostNamedPipeServerCommand : PSCmdlet
    {
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        public string? PipeName { get; set; }

        [Parameter()]
        [ValidateRange(0, int.MaxValue)]
        public int MaxConnections { get; set; } = 0;

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter()]
        [ValidateRange(0, int.MaxValue)]
        public int DrainTimeout { get; set; } = 30;

        protected override void BeginProcessing()
        {
            // Generate random pipe name if not provided
            if (string.IsNullOrWhiteSpace(PipeName))
            {
                PipeName = PSHostNamedPipeServer.GenerateRandomPipeName();
            }

            // Generate default server name if not provided
            if (string.IsNullOrWhiteSpace(Name))
            {
                Name = $"PSHostNamedPipeServer_{PipeName.Substring(0, Math.Min(8, PipeName.Length))}";
            }

            // Check if server with this name already exists
            var existingServer = PSHostServerBase.GetServer(Name);
            if (existingServer != null)
            {
                throw new InvalidOperationException($"Server with name '{Name}' already exists");
            }

            // Check if another server is already listening on this pipe name
            var serverOnPipe = PSHostServerBase.GetAllServers()
                .OfType<PSHostNamedPipeServer>()
                .FirstOrDefault(s => s.PipeName.Equals(PipeName, StringComparison.OrdinalIgnoreCase));
            if (serverOnPipe != null)
            {
                throw new InvalidOperationException($"Server '{serverOnPipe.Name}' is already listening on pipe '{PipeName}'");
            }

            try
            {
                // Create and start the named pipe server
                var server = new PSHostNamedPipeServer(
                    name: Name,
                    pipeName: PipeName,
                    maxConnections: MaxConnections,
                    drainTimeout: DrainTimeout);

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
                    "StartPSHostNamedPipeServerFailed",
                    ErrorCategory.InvalidOperation,
                    Name));
            }
        }
    }

    /// <summary>
    /// Stop-PSHostNamedPipeServer cmdlet - Stops a named pipe-based PowerShell remoting server
    /// </summary>
    [Cmdlet(VerbsLifecycle.Stop, "PSHostNamedPipeServer", DefaultParameterSetName = "ByName")]
    public sealed class StopPSHostNamedPipeServerCommand : PSCmdlet
    {
        [Parameter(ParameterSetName = "ByName", Position = 0, Mandatory = true, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter(ParameterSetName = "ByPipeName", Mandatory = true)]
        [ValidateNotNullOrEmpty()]
        public string? PipeName { get; set; }

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
            else if (ParameterSetName == "ByPipeName")
            {
                server = PSHostServerBase.GetAllServers()
                    .OfType<PSHostNamedPipeServer>()
                    .FirstOrDefault(s => s.PipeName.Equals(PipeName, StringComparison.OrdinalIgnoreCase));
                if (server == null)
                {
                    WriteError(new ErrorRecord(
                        new InvalidOperationException($"No server found listening on pipe '{PipeName}'"),
                        "ServerNotFound",
                        ErrorCategory.ObjectNotFound,
                        PipeName));
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
                    "StopPSHostNamedPipeServerFailed",
                    ErrorCategory.InvalidOperation,
                    server.Name));
            }
        }
    }

    /// <summary>
    /// Get-PSHostNamedPipeServer cmdlet - Gets named pipe-based PowerShell remoting servers
    /// </summary>
    [Cmdlet(VerbsCommon.Get, "PSHostNamedPipeServer", DefaultParameterSetName = "All")]
    [OutputType(typeof(PSHostServerBase))]
    public sealed class GetPSHostNamedPipeServerCommand : PSCmdlet
    {
        [Parameter(ParameterSetName = "ByName", Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true)]
        [ValidateNotNullOrEmpty()]
        public string? Name { get; set; }

        [Parameter(ParameterSetName = "ByPipeName")]
        [ValidateNotNullOrEmpty()]
        public string? PipeName { get; set; }

        [Parameter(ParameterSetName = "All")]
        public SwitchParameter All { get; set; }

        protected override void ProcessRecord()
        {
            if (ParameterSetName == "ByName")
            {
                var server = PSHostServerBase.GetServer(Name!);
                if (server != null && server is PSHostNamedPipeServer)
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
            else if (ParameterSetName == "ByPipeName")
            {
                var server = PSHostServerBase.GetAllServers()
                    .OfType<PSHostNamedPipeServer>()
                    .FirstOrDefault(s => s.PipeName.Equals(PipeName, StringComparison.OrdinalIgnoreCase));
                if (server != null)
                {
                    WriteObject(server);
                }
                else
                {
                    WriteError(new ErrorRecord(
                        new InvalidOperationException($"No server found listening on pipe '{PipeName}'"),
                        "ServerNotFound",
                        ErrorCategory.ObjectNotFound,
                        PipeName));
                }
            }
            else // All
            {
                var servers = PSHostServerBase.GetAllServers()
                    .OfType<PSHostNamedPipeServer>()
                    .ToArray();
                if (servers.Length > 0)
                {
                    WriteObject(servers, enumerateCollection: true);
                }
            }
        }
    }
}
