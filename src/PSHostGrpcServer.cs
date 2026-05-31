using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Threading.Tasks;
using Google.Protobuf;
using Grpc.Core;
using AwakeCoding.PSRemoting.PowerShell.Grpc;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// gRPC-based PowerShell remoting server.
    /// </summary>
    public sealed class PSHostGrpcServer : PSHostServerBase
    {
        private const int BufferSize = 8192;

        private Server? _grpcServer;

        public bool UseSecureConnection { get; private set; }

        public bool AllowUnencrypted { get; private set; }

        public string GrpcUri
        {
            get
            {
                string scheme = UseSecureConnection ? "grpcs" : "grpc";
                return $"{scheme}://{ListenAddress}:{Port}";
            }
        }

        public PSHostGrpcServer(
            string name,
            int port,
            string listenAddress,
            bool useSecureConnection,
            bool allowUnencrypted,
            int maxConnections,
            int drainTimeout)
            : base(name, port, listenAddress, maxConnections, drainTimeout)
        {
            UseSecureConnection = useSecureConnection;
            AllowUnencrypted = allowUnencrypted;
        }

        public override void StartListenerAsync()
        {
            if (State == ServerState.Running || State == ServerState.Starting)
            {
                throw new InvalidOperationException($"Server '{Name}' is already {State}");
            }

            if (UseSecureConnection)
            {
                throw new NotSupportedException("gRPC TLS server credentials are not implemented yet. Use loopback plaintext gRPC for this module transport.");
            }

            if (!AllowUnencrypted && !IsLoopbackAddress(ListenAddress))
            {
                throw new InvalidOperationException("Plaintext gRPC is allowed only on loopback by default. Specify -AllowUnencrypted to bind plaintext gRPC to a non-loopback address.");
            }

            State = ServerState.Starting;

            try
            {
                _serverInstance.CancellationTokenSource = new CancellationTokenSource();

                if (Port == 0)
                {
                    Port = GetAvailableTcpPort(ListenAddress);
                }

                var service = new PSHostGrpcTransportService(this, _serverInstance.CancellationTokenSource.Token);
                var serverPort = new ServerPort(
                    NormalizeListenAddress(ListenAddress),
                    Port,
                    ServerCredentials.Insecure);

                _grpcServer = new Server
                {
                    Services = { PSHostGrpcTransport.BindService(service) },
                    Ports = { serverPort }
                };

                _grpcServer.Start();
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
            if (!TryBeginStopping())
            {
                return;
            }

            if (State != ServerState.Running && State != ServerState.Starting)
            {
                ResetStoppingFlag();
                return;
            }

            State = ServerState.Stopping;

            try
            {
                _serverInstance.CancellationTokenSource?.Cancel();

                if (!force && ConnectionCount > 0)
                {
                    var drainStart = DateTime.UtcNow;
                    var timeout = TimeSpan.FromSeconds(DrainTimeout);

                    while (ConnectionCount > 0 && (DateTime.UtcNow - drainStart) < timeout)
                    {
                        Thread.Sleep(100);
                    }
                }

                var connections = _serverInstance.ActiveConnections.Values.ToArray();
                bool hadRemainingConnections = connections.Length > 0;
                foreach (var connection in connections)
                {
                    if (!connection.ProcessId.HasValue)
                    {
                        continue;
                    }

                    try
                    {
                        var process = Process.GetProcessById(connection.ProcessId.Value);
                        if (!process.HasExited)
                        {
                            process.Kill();
                            process.WaitForExit(ProcessKillWaitTimeoutMs);
                        }
                    }
                    catch { }
                }

                _serverInstance.ActiveConnections.Clear();

                if (_grpcServer != null)
                {
                    if (force || hadRemainingConnections)
                    {
                        _grpcServer.KillAsync().Wait(ListenerThreadJoinTimeoutMs);
                    }
                    else
                    {
                        _grpcServer.ShutdownAsync().Wait(ListenerThreadJoinTimeoutMs);
                    }
                }

                Unregister();
                State = ServerState.Stopped;
            }
            catch (Exception ex)
            {
                Unregister();
                State = ServerState.Failed;
                LastError = ex;
                throw;
            }
            finally
            {
                _serverInstance.CancellationTokenSource?.Dispose();
                _serverInstance.CancellationTokenSource = null;
                _grpcServer = null;
                ResetStoppingFlag();
            }
        }

        protected override void AcceptConnectionAsync()
        {
            throw new NotImplementedException("gRPC connections are accepted by the gRPC server runtime.");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _grpcServer != null)
            {
                try
                {
                    _grpcServer.KillAsync().Wait(ListenerThreadJoinTimeoutMs);
                }
                catch { }
            }

            base.Dispose(disposing);
        }

        private static bool IsLoopbackAddress(string listenAddress)
        {
            if (listenAddress.Equals("localhost", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IPAddress.TryParse(listenAddress, out var address) && IPAddress.IsLoopback(address);
        }

        private static string NormalizeListenAddress(string listenAddress)
        {
            if (listenAddress == "*" || listenAddress == "+")
            {
                return "0.0.0.0";
            }

            return listenAddress;
        }

        private static int GetAvailableTcpPort(string listenAddress)
        {
            IPAddress address;
            if (listenAddress == "*" || listenAddress == "+" || listenAddress == "0.0.0.0")
            {
                address = IPAddress.Any;
            }
            else if (listenAddress.Equals("localhost", StringComparison.OrdinalIgnoreCase))
            {
                address = IPAddress.Loopback;
            }
            else if (!IPAddress.TryParse(listenAddress, out address!))
            {
                address = IPAddress.Loopback;
            }

            var listener = new TcpListener(address, 0);
            listener.Start();
            try
            {
                return ((IPEndPoint)listener.LocalEndpoint).Port;
            }
            finally
            {
                listener.Stop();
            }
        }

        private sealed class PSHostGrpcTransportService : PSHostGrpcTransport.PSHostGrpcTransportBase
        {
            private readonly PSHostGrpcServer _owner;
            private readonly CancellationToken _serverToken;

            public PSHostGrpcTransportService(PSHostGrpcServer owner, CancellationToken serverToken)
            {
                _owner = owner;
                _serverToken = serverToken;
            }

            public override async Task Connect(
                IAsyncStreamReader<PSHostGrpcFrame> requestStream,
                IServerStreamWriter<PSHostGrpcFrame> responseStream,
                ServerCallContext context)
            {
                if (_owner.IsMaxConnectionsReached())
                {
                    throw new RpcException(new Status(StatusCode.ResourceExhausted, "Server at maximum capacity"));
                }

                var executable = PowerShellFinder.GetPowerShellPath();
                if (executable == null || !File.Exists(executable))
                {
                    throw new RpcException(new Status(StatusCode.Internal, "The PowerShell executable path could not be found"));
                }

                var connectionId = Guid.NewGuid().ToString();
                Process? process = null;

                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(_serverToken, context.CancellationToken);
                using var writeLock = new SemaphoreSlim(1, 1);

                try
                {
                    process = new Process();
                    process.StartInfo.FileName = executable;
                    process.StartInfo.Arguments = "-NoLogo -NoProfile -s";
                    process.StartInfo.RedirectStandardInput = true;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.RedirectStandardError = true;
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.CreateNoWindow = true;

                    process.Start();

                    var clientAddress = context.Peer ?? "grpc-client";
                    _owner.AddConnection(new ConnectionDetails(connectionId, clientAddress, process.Id));

                    Task inputTask = ProxyGrpcToProcessAsync(requestStream, process.StandardInput.BaseStream, linkedCts.Token);
                    Task outputTask = ProxyProcessToGrpcAsync(process.StandardOutput.BaseStream, responseStream, PSHostGrpcFrameKind.Output, writeLock, linkedCts.Token);
                    Task errorTask = ProxyProcessToGrpcAsync(process.StandardError.BaseStream, responseStream, PSHostGrpcFrameKind.Error, writeLock, linkedCts.Token);

                    await Task.WhenAny(inputTask, outputTask, errorTask).ConfigureAwait(false);
                    linkedCts.Cancel();

                    await ObserveTaskAsync(inputTask).ConfigureAwait(false);
                    await ObserveTaskAsync(outputTask).ConfigureAwait(false);
                    await ObserveTaskAsync(errorTask).ConfigureAwait(false);
                }
                finally
                {
                    _owner.RemoveConnection(connectionId);

                    if (process != null)
                    {
                        try
                        {
                            if (!process.HasExited)
                            {
                                process.Kill();
                                process.WaitForExit(ProcessKillWaitTimeoutMs);
                            }
                        }
                        catch { }
                        finally
                        {
                            process.Dispose();
                        }
                    }
                }
            }

            private static async Task ProxyGrpcToProcessAsync(
                IAsyncStreamReader<PSHostGrpcFrame> requestStream,
                Stream destination,
                CancellationToken cancellationToken)
            {
                while (await requestStream.MoveNext(cancellationToken).ConfigureAwait(false))
                {
                    var frame = requestStream.Current;
                    if (frame.Kind == PSHostGrpcFrameKind.Close)
                    {
                        break;
                    }

                    if (frame.Kind != PSHostGrpcFrameKind.Input || frame.Payload.Length == 0)
                    {
                        continue;
                    }

                    byte[] buffer = frame.Payload.ToByteArray();
                    await destination.WriteAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
                    await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
                }
            }

            private static async Task ProxyProcessToGrpcAsync(
                Stream source,
                IServerStreamWriter<PSHostGrpcFrame> responseStream,
                PSHostGrpcFrameKind kind,
                SemaphoreSlim writeLock,
                CancellationToken cancellationToken)
            {
                byte[] buffer = new byte[BufferSize];

                while (true)
                {
                    int bytesRead = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
                    if (bytesRead <= 0)
                    {
                        break;
                    }

                    var frame = new PSHostGrpcFrame
                    {
                        Kind = kind,
                        Payload = ByteString.CopyFrom(buffer, 0, bytesRead)
                    };

                    await writeLock.WaitAsync(cancellationToken).ConfigureAwait(false);
                    try
                    {
                        await responseStream.WriteAsync(frame).ConfigureAwait(false);
                    }
                    finally
                    {
                        writeLock.Release();
                    }
                }
            }

            private static async Task ObserveTaskAsync(Task task)
            {
                try
                {
                    await task.ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                }
                catch (IOException)
                {
                }
                catch (ObjectDisposedException)
                {
                }
                catch (RpcException ex) when (ex.StatusCode == StatusCode.Cancelled)
                {
                }
            }
        }
    }
}
