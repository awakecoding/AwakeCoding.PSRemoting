using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// TCP-based PowerShell remoting server
    /// </summary>
    public class PSHostTcpServer : PSHostServerBase
    {
        private TcpListener? _listener;
        private const string ThreadName = "PSHostTcpServer Listener";

        public PSHostTcpServer(string name, int port, string listenAddress, int maxConnections, int drainTimeout)
            : base(name, port, listenAddress, maxConnections, drainTimeout)
        {
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
                // Parse the listen address
                IPAddress ipAddress = IPAddress.Parse(ListenAddress);
                _listener = new TcpListener(ipAddress, Port);
                _listener.Start();

                // If port 0 was specified, update with the actual assigned port
                if (Port == 0)
                {
                    Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
                }

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

                // Stop the TCP listener
                _listener?.Stop();

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
                Unregister();
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
                _listener = null;
            }
        }

        private void ListenerThreadProc()
        {
            var cancellationToken = _serverInstance.CancellationTokenSource?.Token ?? CancellationToken.None;

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    try
                    {
                        // Check if we've reached max connections
                        if (IsMaxConnectionsReached())
                        {
                            Thread.Sleep(100);
                            continue;
                        }

                        // Use Poll to check for pending connections (with timeout to check cancellation)
                        if (_listener?.Pending() == true)
                        {
                            // Accept the connection asynchronously
                            AcceptConnectionAsync();
                        }
                        else
                        {
                            Thread.Sleep(100);
                        }
                    }
                    catch (Exception) when (cancellationToken.IsCancellationRequested)
                    {
                        // Expected during shutdown
                        break;
                    }
                    catch (SocketException)
                    {
                        // Listener was stopped
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                State = ServerState.Failed;
                LastError = ex;
            }
        }

        protected override void AcceptConnectionAsync()
        {
            if (_listener == null)
                return;

            TcpClient? client = null;

            try
            {
                client = _listener.AcceptTcpClient();
                var clientEndpoint = client.Client.RemoteEndPoint?.ToString() ?? "unknown";
                var connectionId = Guid.NewGuid().ToString();

                // Spawn PowerShell subprocess
                var executable = PowerShellFinder.GetPowerShellPath();
                if (executable == null || !File.Exists(executable))
                {
                    client?.Close();
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
                var connectionDetails = new ConnectionDetails(connectionId, clientEndpoint, process.Id);
                AddConnection(connectionDetails);

                // Start proxy threads in background
                var networkStream = client.GetStream();
                StartProxyThreads(networkStream, process, connectionId);
            }
            catch (Exception)
            {
                client?.Close();
                throw;
            }
        }

        private void StartProxyThreads(NetworkStream networkStream, Process process, string connectionId)
        {
            // Thread to proxy network → subprocess stdin
            var inThread = new Thread(() => ProxyNetworkToProcess(networkStream, process.StandardInput.BaseStream, connectionId))
            {
                Name = $"PSHostTcp In {connectionId}",
                IsBackground = true
            };
            inThread.Start();

            // Thread to proxy subprocess stdout → network
            var outThread = new Thread(() => ProxyProcessToNetwork(process.StandardOutput.BaseStream, networkStream, connectionId))
            {
                Name = $"PSHostTcp Out {connectionId}",
                IsBackground = true
            };
            outThread.Start();

            // Thread to read stderr (discard it like PSHostClientSessionTransportMgr does)
            var errThread = new Thread(() => DiscardProcessErrors(process.StandardError, connectionId))
            {
                Name = $"PSHostTcp Err {connectionId}",
                IsBackground = true
            };
            errThread.Start();
        }

        private void ProxyNetworkToProcess(NetworkStream source, Stream destination, string connectionId)
        {
            try
            {
                byte[] buffer = new byte[4096];
                int bytesRead;

                while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    destination.Write(buffer, 0, bytesRead);
                    destination.Flush();
                }
            }
            catch { }
            finally
            {
                CleanupConnection(connectionId);
            }
        }

        private void ProxyProcessToNetwork(Stream source, NetworkStream destination, string connectionId)
        {
            try
            {
                byte[] buffer = new byte[4096];
                int bytesRead;

                while ((bytesRead = source.Read(buffer, 0, buffer.Length)) > 0)
                {
                    destination.Write(buffer, 0, bytesRead);
                    destination.Flush();
                }
            }
            catch { }
            finally
            {
                CleanupConnection(connectionId);
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

        private void CleanupConnection(string connectionId)
        {
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

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _listener?.Stop();
            }

            base.Dispose(disposing);
        }
    }
}
