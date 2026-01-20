using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace AwakeCoding.PSRemoting.PowerShell
{
    /// <summary>
    /// Server state enum
    /// </summary>
    public enum ServerState
    {
        Stopped,
        Starting,
        Running,
        Stopping,
        Failed
    }

    /// <summary>
    /// Connection details for tracking active connections
    /// </summary>
    public class ConnectionDetails
    {
        public string ConnectionId { get; set; }
        public string? ClientAddress { get; set; }
        public DateTime ConnectedAt { get; set; }
        public int? ProcessId { get; set; }
        public TimeSpan Uptime => DateTime.UtcNow - ConnectedAt;

        public ConnectionDetails(string connectionId, string? clientAddress, int? processId = null)
        {
            ConnectionId = connectionId;
            ClientAddress = clientAddress;
            ConnectedAt = DateTime.UtcNow;
            ProcessId = processId;
        }
    }

    /// <summary>
    /// Helper class for managing server instance state
    /// </summary>
    public class ServerInstance
    {
        public Thread? ListenerThread { get; set; }
        public CancellationTokenSource? CancellationTokenSource { get; set; }
        public ConcurrentDictionary<string, ConnectionDetails> ActiveConnections { get; set; }

        public ServerInstance()
        {
            ActiveConnections = new ConcurrentDictionary<string, ConnectionDetails>();
        }
    }

    /// <summary>
    /// Abstract base class for PowerShell remoting servers
    /// </summary>
    public abstract class PSHostServerBase : IDisposable
    {
        // Global registry of all active servers
        private static readonly ConcurrentDictionary<string, PSHostServerBase> _servers 
            = new ConcurrentDictionary<string, PSHostServerBase>(StringComparer.OrdinalIgnoreCase);

        protected ServerInstance _serverInstance;
        private object _stateLock = new object();
        private ServerState _state = ServerState.Stopped;
        private Exception? _lastError;

        /// <summary>
        /// Unique server name
        /// </summary>
        public string Name { get; protected set; }

        /// <summary>
        /// Port the server is listening on
        /// </summary>
        public int Port { get; protected set; }

        /// <summary>
        /// IP address the server is bound to
        /// </summary>
        public string ListenAddress { get; protected set; }

        /// <summary>
        /// Maximum number of concurrent connections (0 = unlimited)
        /// </summary>
        public int MaxConnections { get; protected set; }

        /// <summary>
        /// Current number of active connections
        /// </summary>
        public int ConnectionCount => _serverInstance?.ActiveConnections.Count ?? 0;

        /// <summary>
        /// Current server state
        /// </summary>
        public ServerState State 
        { 
            get 
            { 
                lock (_stateLock) 
                { 
                    return _state; 
                } 
            } 
            protected set 
            { 
                lock (_stateLock) 
                { 
                    _state = value; 
                } 
            } 
        }

        /// <summary>
        /// Last error that occurred (if State is Failed)
        /// </summary>
        public Exception? LastError 
        { 
            get 
            { 
                lock (_stateLock) 
                { 
                    return _lastError; 
                } 
            } 
            protected set 
            { 
                lock (_stateLock) 
                { 
                    _lastError = value; 
                } 
            } 
        }

        /// <summary>
        /// Drain timeout for graceful shutdown (seconds)
        /// </summary>
        public int DrainTimeout { get; set; }

        protected PSHostServerBase(string name, int port, string listenAddress, int maxConnections, int drainTimeout)
        {
            Name = name;
            Port = port;
            ListenAddress = listenAddress;
            MaxConnections = maxConnections;
            DrainTimeout = drainTimeout;
            _serverInstance = new ServerInstance();
        }

        /// <summary>
        /// Start the server listener
        /// </summary>
        public abstract void StartListenerAsync();

        /// <summary>
        /// Stop the server listener
        /// </summary>
        /// <param name="force">If true, immediately kill all connections. If false, wait for drain timeout.</param>
        public abstract void StopListenerAsync(bool force);

        /// <summary>
        /// Accept a new connection (called by listener thread)
        /// </summary>
        protected abstract void AcceptConnectionAsync();

        /// <summary>
        /// Get list of active connections
        /// </summary>
        public List<ConnectionDetails> GetConnections()
        {
            return _serverInstance?.ActiveConnections.Values.ToList() ?? new List<ConnectionDetails>();
        }

        /// <summary>
        /// Add a connection to tracking
        /// </summary>
        protected void AddConnection(ConnectionDetails details)
        {
            _serverInstance?.ActiveConnections.TryAdd(details.ConnectionId, details);
        }

        /// <summary>
        /// Remove a connection from tracking
        /// </summary>
        protected void RemoveConnection(string connectionId)
        {
            _serverInstance?.ActiveConnections.TryRemove(connectionId, out _);
        }

        /// <summary>
        /// Check if max connections limit is reached
        /// </summary>
        protected bool IsMaxConnectionsReached()
        {
            return MaxConnections > 0 && ConnectionCount >= MaxConnections;
        }

        /// <summary>
        /// Register this server in the global registry
        /// </summary>
        public bool Register()
        {
            return _servers.TryAdd(Name, this);
        }

        /// <summary>
        /// Unregister this server from the global registry
        /// </summary>
        public void Unregister()
        {
            _servers.TryRemove(Name, out _);
        }

        /// <summary>
        /// Get a server by name from the global registry
        /// </summary>
        public static PSHostServerBase? GetServer(string name)
        {
            _servers.TryGetValue(name, out var server);
            return server;
        }

        /// <summary>
        /// Get all servers from the global registry
        /// </summary>
        public static IEnumerable<PSHostServerBase> GetAllServers()
        {
            return _servers.Values;
        }

        /// <summary>
        /// Get a server by port from the global registry
        /// </summary>
        public static PSHostServerBase? GetServerByPort(int port)
        {
            return _servers.Values.FirstOrDefault(s => s.Port == port);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (State == ServerState.Running || State == ServerState.Starting)
                {
                    StopListenerAsync(force: true);
                }
                Unregister();
            }
        }
    }
}
