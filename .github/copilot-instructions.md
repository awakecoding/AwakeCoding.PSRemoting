# AwakeCoding.PSRemoting - Copilot Instructions

## Project Overview

This is a **hybrid .NET/PowerShell module** that provides PowerShell remoting capabilities without network transport. It supports:
1. **Client Sessions**: Creating PSSession objects connected to local PowerShell subprocesses via stdio streams
2. **Server Infrastructure**: Hosting PowerShell remoting endpoints via TCP, WebSocket, and Named Pipe transports

**Architecture**: C# binary module (compiled to DLL) → loaded by PowerShell → exposes client and server cmdlets

## Key Components

### Client-Side Components
- [`src/PSHostSessionCommandBase.cs`](../src/PSHostSessionCommandBase.cs) - Abstract base class for session cmdlets
  - Shares parameters and connection logic across `New-PSHostSession` and `Enter-PSHostSession`
  - Supports 6 parameter sets: Subprocess, TCP, WebSocket, NamedPipe, ProcessId, SSH
  - Provides `CreateAndOpenRunspace()` method and transport-specific connection factories
- [`src/PSHostSessionCommands.cs`](../src/PSHostSessionCommands.cs) - Client cmdlets for creating sessions
  - `NewPSHostSessionCommand` - Creates PSSessions to local subprocesses (extends base class)
  - `EnterPSHostSessionCommand` - Creates and enters sessions interactively (extends base class)
  - `ConnectPSHostProcessCommand` - Connects to existing PowerShell processes via named pipes
- [`src/PSHostClientTransport.cs`](../src/PSHostClientTransport.cs) - Client transport using Process stdio
- [`src/PSHostNamedPipeTransport.cs`](../src/PSHostNamedPipeTransport.cs) - Named pipe client transport
- [`src/PSHostTcpClientTransport.cs`](../src/PSHostTcpClientTransport.cs) - TCP client transport
- [`src/PSHostWebSocketClientTransport.cs`](../src/PSHostWebSocketClientTransport.cs) - WebSocket client transport
- [`src/PSHostSSHClientTransport.cs`](../src/PSHostSSHClientTransport.cs) - SSH client transport

### Server-Side Components
- [`src/PSHostServerCommands.cs`](../src/PSHostServerCommands.cs) - **Unified server cmdlets** (3 cmdlets for all transports)
  - `Start-PSHostServer -TransportType <TCP|WebSocket|NamedPipe>` - Start remoting server
  - `Stop-PSHostServer` - Stop server by name/port/pipe/object
  - `Get-PSHostServer` - Query running servers with optional transport filter
- [`src/PSHostServerBase.cs`](../src/PSHostServerBase.cs) - Abstract base class for all server implementations
- [`src/PSHostTcpServer.cs`](../src/PSHostTcpServer.cs) - TCP transport server (TcpListener)
- [`src/PSHostWebSocketServer.cs`](../src/PSHostWebSocketServer.cs) - WebSocket transport server (HttpListener)
- [`src/PSHostNamedPipeServer.cs`](../src/PSHostNamedPipeServer.cs) - Named Pipe transport server (NamedPipeServerStream)
- [`src/PSHostTransportType.cs`](../src/PSHostTransportType.cs) - Enum defining transport types

### Shared Components
- [`src/PowerShellFinder.cs`](../src/PowerShellFinder.cs) - Utilities to locate PowerShell executables and generate pipe names
- [`src/PSHostTcpServerTransport.cs`](../src/PSHostTcpServerTransport.cs) - Internal TCP transport connection info and manager (server-side)
- [`AwakeCoding.PSRemoting/AwakeCoding.PSRemoting.psd1`](../AwakeCoding.PSRemoting/AwakeCoding.PSRemoting.psd1) - Module manifest (exports 6 cmdlets total)

### Deprecated Files
- [`src/PSHostSession.cs`](../src/PSHostSession.cs) - Deprecated code wrapped in `#if false` (types moved to other files, kept for reference)

## Build & Development Workflow

Build the module using the provided script:
```powershell
.\build.ps1
```

This script:
1. Compiles the C# project: `dotnet build -c Release -f net8.0`
2. Copies the built DLL from `src/bin/Release/net8.0/` to `AwakeCoding.PSRemoting/` folder

**Testing locally**: After building, import the module from the workspace:
```powershell
Import-Module .\AwakeCoding.PSRemoting\AwakeCoding.PSRemoting.psd1

# Client usage - create local subprocess session
New-PSHostSession | Enter-PSSession

# Server usage - start TCP server
$server = Start-PSHostServer -TransportType TCP -Port 8080
# Connect from another PowerShell session
Enter-PSSession -HostName localhost -Port 8080
```

**Run tests**:
```powershell
.\test.ps1
```

## Project Conventions

- **Target Framework**: .NET 8.0 (net8.0)
- **PowerShell Version**: Requires PowerShell 7.2+ (CompatiblePSEditions = Core only)
- **Namespace**: All C# code lives in `AwakeCoding.PSRemoting.PowerShell`
- **Cmdlet Naming**: PowerShell approved verbs only (e.g., `New-PSHostSession`)
- **Transport Arguments**: Host processes always use `-NoLogo -NoProfile -s` (server mode)

## Critical Implementation Details

**Custom Transport Layer**: This project bypasses PowerShell's standard network-based remoting (WinRM/SSH) by implementing a custom `ClientSessionTransportManagerBase`. The transport reads/writes remoting protocol data over Process stdio streams instead of sockets.

**Process Lifecycle**: The spawned PowerShell subprocess is managed through `Process.Start()` with redirected stdio. It runs with `-s` flag (server mode) which makes it listen for remoting protocol on stdin/stdout. When the PSSession is closed/removed, the process is killed.

**PowerShell Discovery**: 
- If running from pwsh/powershell, uses the same executable for the subprocess
- Otherwise searches PATH for `pwsh.exe`/`pwsh` or `powershell.exe` (Windows only)
- Can be overridden with `-ExecutablePath` parameter

**Server Infrastructure**: Servers inherit from `PSHostServerBase` and maintain a global registry. Each transport spawns pwsh subprocesses in server mode (`-s`) and proxies bidirectional data. Servers track state (Stopped/Starting/Running/Stopping/Failed), connection count, and active connections. All servers properly unregister from the global registry when stopped.

**Transport Types**:
- **TCP**: Uses `TcpListener`, supports port 0 for auto-assignment, thread-based connection handling
- **WebSocket**: Uses `HttpListener`, requires specific port (>0), supports custom path and SSL, async connection handling
- **NamedPipe**: Uses `NamedPipeServerStream` with `MaxAllowedServerInstances` for concurrent connections, platform-specific paths (Windows: `\\.\pipe\{name}`, Linux: `/tmp/CoreFxPipe_{name}`)

## Common Patterns

### Shared Base Class Architecture

Client session cmdlets (`New-PSHostSession` and `Enter-PSHostSession`) inherit from `PSHostSessionCommandBase` to share:
- **All parameters**: Subprocess, TCP, WebSocket, NamedPipe, ProcessId, and SSH parameter sets
- **Connection logic**: Transport-specific connection info factory methods
- **Runspace opening**: `CreateAndOpenRunspace()` handles async open, timeout, and readiness checks
- **Derived implementation**: Each cmdlet implements `ProcessSession(Runspace)` for specific behavior
  - `NewPSHostSessionCommand.ProcessSession()` → Creates and outputs PSSession object
  - `EnterPSHostSessionCommand.ProcessSession()` → Creates PSSession and calls built-in `Enter-PSSession`

This pattern eliminates ~283 lines of duplicated code and ensures parameter consistency.

### Adding New Cmdlet Parameters

For client session cmdlets:
- Add properties to `PSHostSessionCommandBase` to share across both cmdlets
- Mark with `[Parameter()]` attribute and specify applicable parameter sets
- Use `[ValidateNotNullOrEmpty()]` for mandatory string parameters
- Access via properties in `CreateAndOpenRunspace()` or connection factory methods

For server cmdlets:
- Unified cmdlets use `-TransportType` enum to select transport
- Transport-specific parameters (Port, PipeName, Path, etc.) are conditionally validated
- Common parameters (Name, MaxConnections, DrainTimeout) apply to all transports

### Modifying Server Behavior

Adding new server features:
1. Add property/method to `PSHostServerBase` if shared across all transports
2. Override in specific transport class if transport-specific
3. Update unified cmdlet to expose new parameter if user-facing
4. Remember to update `StopListenerAsync()` to cleanup new resources

Transport connection handling:
- TCP/NamedPipe: Thread-based proxy with synchronous read/write loops
- WebSocket: Async/await pattern with `WebSocketReceiveResult`
- All: Spawn subprocess with `pwsh -NoLogo -NoProfile -s`, proxy stdio to transport

### Server Lifecycle

1. **Start**: `StartListenerAsync()` → creates listener → starts background thread → sets state to Running → registers in global registry
2. **Accept**: Background thread waits for connections → spawns pwsh subprocess → creates proxy threads (in/out/err)
3. **Stop**: `StopListenerAsync()` → cancels listener → drains connections (unless Force) → kills subprocesses → **unregisters from registry**
4. **Cleanup**: `CleanupConnection()` disposes transport resources, kills subprocess, removes from connection tracking

### Client Transport Behavior

When modifying transport behavior:
- Edit `PSHostClientSessionTransportMgr.CreateAsync()` for connection setup
- Override `CleanupConnection()` to ensure proper resource disposal
- Remember to handle both normal and error data streams from Process

## Dependencies

- `System.Management.Automation` (PowerShell SDK 7.x) - Required for cmdlet development
- No network/remoting dependencies - this is entirely local process communication
