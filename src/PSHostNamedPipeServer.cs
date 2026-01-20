using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Threading;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Named pipe-based PowerShell remoting server
    /// </summary>
    public sealed class PSHostNamedPipeServer : PSHostServerBase
    {
        private const string ThreadName = "PSHostNamedPipe Listener";
        private const int BufferSize = 4096;
        
        /// <summary>
        /// Name of the pipe (without platform-specific prefix)
        /// </summary>
        public string PipeName { get; private set; }
        
        /// <summary>
        /// Full platform-specific pipe path
        /// </summary>
        public string PipeFullPath => GetPipeFullPath(PipeName);

        public PSHostNamedPipeServer(string name, string pipeName, int maxConnections, int drainTimeout)
            : base(name, port: 0, listenAddress: "localhost", maxConnections, drainTimeout)
        {
            PipeName = pipeName;
        }

        public override void StartListenerAsync()
        {
            if (State == ServerState.Running || State == ServerState.Starting)
            {
                throw new InvalidOperationException($"Server '{Name}' is already {State}");
            }

            State = ServerState.Starting;

            try
            {
                // Create cancellation token source
                _serverInstance.CancellationTokenSource = new CancellationTokenSource();

                // Start listener thread
                _serverInstance.ListenerThread = new Thread(ListenerThreadProc)
                {
                    Name = ThreadName,
                    IsBackground = true
                };
                _serverInstance.ListenerThread.Start();

                State = ServerState.Running;
            }
            catch (Exception ex)
            {
                State = ServerState.Failed;
                LastError = ex;
                throw;
            }
        }

        public override void StopListenerAsync(bool force)
        {
            if (State != ServerState.Running && State != ServerState.Starting)
            {
                return;
            }

            State = ServerState.Stopping;

            try
            {
                // Signal cancellation to stop accepting new connections
                _serverInstance.CancellationTokenSource?.Cancel();

                if (!force && ConnectionCount > 0)
                {
                    // Wait for drain timeout for connections to complete
                    var drainStart = DateTime.UtcNow;
                    var timeout = TimeSpan.FromSeconds(DrainTimeout);
                    
                    while (ConnectionCount > 0 && (DateTime.UtcNow - drainStart) < timeout)
                    {
                        Thread.Sleep(100);
                    }
                }

                // Force kill any remaining connections and their subprocesses
                var connections = _serverInstance.ActiveConnections.Values;
                foreach (var connection in connections)
                {
                    if (connection.ProcessId.HasValue)
                    {
                        try
                        {
                            var process = Process.GetProcessById(connection.ProcessId.Value);
                            if (!process.HasExited)
                            {
                                process.Kill();
                                process.WaitForExit(500);
                            }
                        }
                        catch { }
                    }
                }

                // Clear all connections
                _serverInstance.ActiveConnections.Clear();

                // Wait for listener thread to exit
                _serverInstance.ListenerThread?.Join(1000);

                State = ServerState.Stopped;
            }
            catch (Exception ex)
            {
                State = ServerState.Failed;
                LastError = ex;
                throw;
            }
            finally
            {
                _serverInstance.CancellationTokenSource?.Dispose();
                _serverInstance.CancellationTokenSource = null;
            }
        }

        private void ListenerThreadProc()
        {
            var cancellationToken = _serverInstance.CancellationTokenSource?.Token ?? CancellationToken.None;

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    NamedPipeServerStream? pipeServer = null;
                    
                    try
                    {
                        // Check if we've reached max connections
                        if (IsMaxConnectionsReached())
                        {
                            Thread.Sleep(100);
                            continue;
                        }

                        // Create a new pipe server instance for each connection
                        // MaxAllowedServerInstances allows multiple concurrent connections
                        pipeServer = new NamedPipeServerStream(
                            PipeName,
                            PipeDirection.InOut,
                            NamedPipeServerStream.MaxAllowedServerInstances,
                            PipeTransmissionMode.Byte,
                            PipeOptions.Asynchronous,
                            BufferSize,
                            BufferSize);

                        // Wait for connection with cancellation support
                        var connectTask = pipeServer.WaitForConnectionAsync(cancellationToken);
                        
                        try
                        {
                            connectTask.Wait(cancellationToken);
                        }
                        catch (OperationCanceledException)
                        {
                            pipeServer.Dispose();
                            break;
                        }
                        catch (AggregateException ex) when (ex.InnerException is OperationCanceledException)
                        {
                            pipeServer.Dispose();
                            break;
                        }

                        // Connection accepted - handle it
                        AcceptPipeConnection(pipeServer);
                    }
                    catch (IOException)
                    {
                        // Pipe broken or closed
                        pipeServer?.Dispose();
                    }
                    catch (Exception) when (cancellationToken.IsCancellationRequested)
                    {
                        pipeServer?.Dispose();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                if (!cancellationToken.IsCancellationRequested)
                {
                    State = ServerState.Failed;
                    LastError = ex;
                }
            }
        }

        private void AcceptPipeConnection(NamedPipeServerStream pipeServer)
        {
            var connectionId = Guid.NewGuid().ToString();

            try
            {
                // Get client info (limited for named pipes)
                var clientAddress = pipeServer.IsConnected ? "pipe-client" : "unknown";

                // Spawn PowerShell subprocess
                var executable = PowerShellFinder.GetPowerShellPath();
                if (executable == null || !File.Exists(executable))
                {
                    pipeServer.Dispose();
                    return;
                }

                var process = new Process();
                process.StartInfo.FileName = executable;
                process.StartInfo.Arguments = "-NoLogo -NoProfile -s";
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;

                process.Start();

                // Add connection to tracking
                var connectionDetails = new ConnectionDetails(connectionId, clientAddress, process.Id);
                AddConnection(connectionDetails);

                // Start proxy threads in background
                StartProxyThreads(pipeServer, process, connectionId);
            }
            catch (Exception)
            {
                pipeServer.Dispose();
                throw;
            }
        }

        private void StartProxyThreads(NamedPipeServerStream pipeStream, Process process, string connectionId)
        {
            // Thread to proxy pipe → subprocess stdin
            var inThread = new Thread(() => ProxyPipeToProcess(pipeStream, process.StandardInput.BaseStream, connectionId))
            {
                Name = $"PSHostPipe In {connectionId.Substring(0, 8)}",
                IsBackground = true
            };
            inThread.Start();

            // Thread to proxy subprocess stdout → pipe
            var outThread = new Thread(() => ProxyProcessToPipe(process.StandardOutput.BaseStream, pipeStream, connectionId))
            {
                Name = $"PSHostPipe Out {connectionId.Substring(0, 8)}",
                IsBackground = true
            };
            outThread.Start();

            // Thread to read stderr (discard it)
            var errThread = new Thread(() => DiscardProcessErrors(process.StandardError, connectionId))
            {
                Name = $"PSHostPipe Err {connectionId.Substring(0, 8)}",
                IsBackground = true
            };
            errThread.Start();
        }

        private void ProxyPipeToProcess(NamedPipeServerStream source, Stream destination, string connectionId)
        {
            try
            {
                byte[] buffer = new byte[BufferSize];
                int bytesRead;

                while (source.IsConnected && (bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    destination.Write(buffer, 0, bytesRead);
                    destination.Flush();
                }
            }
            catch { }
            finally
            {
                CleanupConnection(connectionId, source);
            }
        }

        private void ProxyProcessToPipe(Stream source, NamedPipeServerStream destination, string connectionId)
        {
            try
            {
                byte[] buffer = new byte[BufferSize];
                int bytesRead;

                while (destination.IsConnected && (bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    destination.Write(buffer, 0, bytesRead);
                    destination.Flush();
                }
            }
            catch { }
            finally
            {
                CleanupConnection(connectionId, destination);
            }
        }

        private void DiscardProcessErrors(StreamReader errorReader, string connectionId)
        {
            try
            {
                while (errorReader.ReadLine() != null)
                {
                    // Discard error output
                }
            }
            catch { }
        }

        private void CleanupConnection(string connectionId, NamedPipeServerStream? pipeStream)
        {
            // Dispose the pipe stream
            try
            {
                pipeStream?.Dispose();
            }
            catch { }

            // Remove connection from tracking and kill subprocess if still running
            if (_serverInstance.ActiveConnections.TryRemove(connectionId, out var connection))
            {
                if (connection.ProcessId.HasValue)
                {
                    try
                    {
                        var process = Process.GetProcessById(connection.ProcessId.Value);
                        if (!process.HasExited)
                        {
                            process.Kill();
                            process.WaitForExit(500);
                        }
                        process.Dispose();
                    }
                    catch { }
                }
            }
        }

        protected override void AcceptConnectionAsync()
        {
            // Not used - pipe connections are handled via AcceptPipeConnection
            throw new NotImplementedException("Pipe connections are handled via AcceptPipeConnection");
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

        /// <summary>
        /// Gets the full platform-specific pipe path
        /// </summary>
        public static string GetPipeFullPath(string pipeName)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return $@"\\.\pipe\{pipeName}";
            }
            else
            {
                return Path.Combine(Path.GetTempPath(), $"CoreFxPipe_{pipeName}");
            }
        }

        /// <summary>
        /// Generates a random pipe name using GUID
        /// </summary>
        public static string GenerateRandomPipeName()
        {
            return $"PSHost_{Guid.NewGuid():N}";
        }
    }
}
