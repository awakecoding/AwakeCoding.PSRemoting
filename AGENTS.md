# AwakeCoding.PSRemoting - Agent Instructions

## Project Overview

This is a **hybrid .NET/PowerShell module** that provides PowerShell remoting capabilities without requiring a WinRM listener for the common local and custom transport scenarios. It supports:
1. **Client Sessions**: Creating PSSession objects connected to local PowerShell subprocesses via stdio streams or remote endpoints via custom transports
2. **Server Infrastructure**: Hosting PowerShell remoting endpoints via TCP, WebSocket, Named Pipe, WinRM, and gRPC transports

**Architecture**: C# binary module (compiled to DLL) -> loaded by PowerShell -> exposes client and server cmdlets.

## Key Components

### Client-Side Components
- [`src/PSHostSessionCommandBase.cs`](src/PSHostSessionCommandBase.cs) - Abstract base class for session cmdlets
  - Shares parameters and connection logic across `New-PSHostSession` and `Enter-PSHostSession`
  - Supports parameter sets for Subprocess, TCP, WebSocket, NamedPipe, ProcessId, SSH, WinRM (`ComputerName` and `ConnectionUri`), and gRPC
  - Provides `CreateAndOpenRunspace()` and transport-specific connection factory methods
- [`src/PSHostSessionCommands.cs`](src/PSHostSessionCommands.cs) - Client cmdlets for creating sessions
  - `NewPSHostSessionCommand` - Creates PSSessions (extends base class)
  - `EnterPSHostSessionCommand` - Creates and enters sessions interactively (extends base class)
  - `ConnectPSHostProcessCommand` - Connects to existing PowerShell processes via named pipes
- [`src/PSHostClientTransport.cs`](src/PSHostClientTransport.cs) - Client transport using process stdio
- [`src/PSHostTcpClientTransport.cs`](src/PSHostTcpClientTransport.cs) - TCP client transport
- [`src/PSHostWebSocketClientTransport.cs`](src/PSHostWebSocketClientTransport.cs) - WebSocket client transport
- [`src/PSHostNamedPipeTransport.cs`](src/PSHostNamedPipeTransport.cs) - Named pipe client transport
- [`src/PSHostSSHClientTransport.cs`](src/PSHostSSHClientTransport.cs) - SSH client transport
- [`src/PSHostWinRMClientTransport.cs`](src/PSHostWinRMClientTransport.cs) - WinRM/WSMan client transport
- [`src/PSHostGrpcClientTransport.cs`](src/PSHostGrpcClientTransport.cs) - gRPC client transport

### Server-Side Components
- [`src/PSHostServerCommands.cs`](src/PSHostServerCommands.cs) - Unified server cmdlets
  - `Start-PSHostServer -TransportType <TCP|WebSocket|NamedPipe|WinRM|Grpc>` - Start remoting server
  - `Stop-PSHostServer` - Stop server by name, port, pipe, or server object
  - `Get-PSHostServer` - Query running servers with optional transport filter
- [`src/PSHostServerBase.cs`](src/PSHostServerBase.cs) - Abstract base class for server implementations
- [`src/PSHostTcpServer.cs`](src/PSHostTcpServer.cs) - TCP transport server (`TcpListener`)
- [`src/PSHostWebSocketServer.cs`](src/PSHostWebSocketServer.cs) - WebSocket transport server (`HttpListener`)
- [`src/PSHostNamedPipeServer.cs`](src/PSHostNamedPipeServer.cs) - Named Pipe transport server (`NamedPipeServerStream`)
- [`src/PSHostWinRMServer.cs`](src/PSHostWinRMServer.cs) - WinRM/WSMan listener backed by a PowerShell subprocess
- [`src/PSHostGrpcServer.cs`](src/PSHostGrpcServer.cs) - gRPC transport server
- [`src/PSHostTransportType.cs`](src/PSHostTransportType.cs) - Enum defining transport types

### Shared Components
- [`src/PowerShellFinder.cs`](src/PowerShellFinder.cs) - Utilities to locate PowerShell executables and generate pipe names
- [`src/PSHostTcpServerTransport.cs`](src/PSHostTcpServerTransport.cs) - Internal TCP transport connection info and manager (server-side)
- [`src/PSHostGrpcPlatform.cs`](src/PSHostGrpcPlatform.cs) - gRPC platform support guard
- [`src/Protos/pshostgrpc.proto`](src/Protos/pshostgrpc.proto) - gRPC transport service contract
- [`src/WinRMSspiAuth.cs`](src/WinRMSspiAuth.cs) - WinRM SSPI authentication helpers
- [`src/AwakeCoding.PSRemoting.psd1`](src/AwakeCoding.PSRemoting.psd1) - Module manifest

### Deprecated Files
- [`src/PSHostSession.cs`](src/PSHostSession.cs) - Deprecated code wrapped in `#if false` (types moved to other files, kept for reference)

## Build & Development Workflow

Build the module using the provided script:
```powershell
.\build.ps1
```

This script:
1. Compiles the C# project: `dotnet build -c Release -f net8.0`
2. Copies the built DLL and required dependency DLLs to the module output folder

**Testing locally**: After building, import the module from the workspace:
```powershell
Import-Module .\src\AwakeCoding.PSRemoting.psd1

# Client usage - create local subprocess session
New-PSHostSession | Enter-PSSession

# Server usage - start TCP server
$server = Start-PSHostServer -TransportType TCP -Port 8080
Enter-PSSession -HostName localhost -Port 8080

# Server usage - start gRPC server
$grpcServer = Start-PSHostServer -TransportType Grpc -Port 0
$session = New-PSHostSession -GrpcUri $grpcServer.GrpcUri
```

**Run tests**:
```powershell
.\test.ps1
```

## Project Conventions

- **Target Framework**: .NET 8.0 (net8.0)
- **PowerShell Version**: Requires PowerShell 7.2+ (CompatiblePSEditions = Core only)
- **Namespace**: All C# code lives in `AwakeCoding.PSRemoting.PowerShell`
- **Cmdlet Naming**: PowerShell approved verbs only (for example, `New-PSHostSession`)
- **Transport Arguments**: Host subprocesses use `-NoLogo -NoProfile -s` for PSRP server mode unless a transport requires a different PowerShell server mode
- **Docs**: Keep README transport examples in sync when adding or changing user-facing transport parameters

## Critical Implementation Details

**Custom Transport Layer**: This project bypasses PowerShell's standard network-based remoting stack for custom transports by implementing `ClientSessionTransportManagerBase`. Transport implementations read and write PSRP data over process stdio, TCP sockets, WebSockets, named pipes, SSH, WinRM/WSMan, or gRPC streams.

**Process Lifecycle**: The spawned PowerShell subprocess is managed through `Process.Start()` with redirected stdio. For stdio, TCP, WebSocket, Named Pipe, WinRM server, and gRPC server flows, the child process runs in server mode with `-s` and listens for remoting protocol data. When the PSSession is closed or removed, the process is cleaned up.

**PowerShell Discovery**:
- If running from pwsh/powershell, uses the same executable for subprocesses
- Otherwise searches PATH for `pwsh.exe`/`pwsh` or `powershell.exe` on Windows
- Can be overridden with `-ExecutablePath` for subprocess sessions

**Server Infrastructure**: Servers inherit from `PSHostServerBase` and maintain a global registry. Servers track state (`Stopped`, `Starting`, `Running`, `Stopping`, `Failed`), connection count, and active connections. All servers must unregister from the global registry when stopped.

**Transport Types**:
- **TCP**: Uses `TcpListener`, supports port 0 for auto-assignment, thread-based connection handling
- **WebSocket**: Uses `HttpListener`, requires a specific port (>0), supports custom path and SSL switch, async connection handling
- **NamedPipe**: Uses `NamedPipeServerStream` with `MaxAllowedServerInstances` for concurrent connections, platform-specific paths (Windows: `\\.\pipe\{name}`, Linux: `/tmp/CoreFxPipe_{name}`)
- **WinRM**: Uses a lightweight WSMan listener/client implementation with SOAP over HTTP(S); supports explicit credentials and SSPI-backed Negotiate/Kerberos paths
- **gRPC**: Uses `Grpc.Core` duplex streaming, supports port 0 auto-assignment and exposes `GrpcUri`; plaintext is loopback-only by default unless `-AllowUnencrypted` is specified; TLS server credentials are not implemented; macOS arm64 is unsupported by the native runtime

## Common Patterns

### Shared Base Class Architecture

Client session cmdlets (`New-PSHostSession` and `Enter-PSHostSession`) inherit from `PSHostSessionCommandBase` to share:
- **All parameters**: Subprocess, TCP, WebSocket, NamedPipe, ProcessId, SSH, WinRM, and gRPC parameter sets
- **Connection logic**: Transport-specific connection info factory methods
- **Runspace opening**: `CreateAndOpenRunspace()` handles async open, timeout, and readiness checks
- **Derived implementation**: Each cmdlet implements `ProcessSession(Runspace)` for specific behavior
  - `NewPSHostSessionCommand.ProcessSession()` -> Creates and outputs a PSSession object
  - `EnterPSHostSessionCommand.ProcessSession()` -> Creates a PSSession and calls built-in `Enter-PSSession`

### Adding New Cmdlet Parameters

For client session cmdlets:
- Add shared properties to `PSHostSessionCommandBase`
- Mark with `[Parameter()]` and specify applicable parameter sets
- Use `[ValidateNotNullOrEmpty()]` for mandatory string parameters
- Access through `CreateAndOpenRunspace()` or connection factory methods

For server cmdlets:
- Unified cmdlets use `-TransportType` to select transport
- Transport-specific parameters (Port, PipeName, Path, WinRMPath, AllowUnencrypted, etc.) are conditionally validated in server factory methods
- Common parameters (`Name`, `MaxConnections`, `DrainTimeout`) apply to all transports

### Modifying Server Behavior

Adding new server features:
1. Add property/method to `PSHostServerBase` if shared across all transports
2. Override in specific transport class if transport-specific
3. Update unified cmdlet to expose a new parameter if user-facing
4. Update README sample usage for user-facing behavior
5. Ensure `StopListenerAsync()` cleans up new resources

Transport connection handling:
- TCP/NamedPipe: Thread-based proxy with synchronous read/write loops
- WebSocket: Async/await pattern with `WebSocketReceiveResult`
- WinRM: HTTP/WSMan request handling around a subprocess-backed PSRP stream
- gRPC: Duplex streaming frames between the gRPC stream and subprocess stdio
- All subprocess-backed server transports: proxy stdio to the selected transport and clean up child processes reliably

### Server Lifecycle

1. **Start**: `StartListenerAsync()` creates listener/runtime resources, starts accepting connections, sets state to `Running`, and keeps the server registered
2. **Accept**: Listener accepts connections, spawns PowerShell subprocesses as needed, and creates proxy tasks/threads
3. **Stop**: `StopListenerAsync()` cancels listener/runtime resources, drains connections unless forced, kills subprocesses when needed, and unregisters from the registry
4. **Cleanup**: Connection cleanup disposes transport resources, kills subprocesses, and removes connection tracking

### Client Transport Behavior

When modifying transport behavior:
- Edit the transport-specific `CreateAsync()`/setup path for connection establishment
- Override cleanup paths to ensure process, socket, pipe, stream, channel, or call resources are disposed
- Preserve both output and error stream handling
- Keep `OpenTimeout` behavior consistent with other transports

## Dependencies

- `System.Management.Automation` (PowerShell SDK 7.x) - Required for cmdlet development
- `Tmds.Ssh` - SSH transport support
- `Grpc.Core` and `Grpc.Core.Api` - gRPC transport support
- `Devolutions.Sspi` runtime dependencies - WinRM Negotiate/Kerberos support
