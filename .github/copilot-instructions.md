# AwakeCoding.PSRemoting - Copilot Instructions

## Project Overview

This is a **hybrid .NET/PowerShell module** that creates PSSession objects connected to PowerShell host subprocesses. It enables PowerShell remoting without network transport by spawning local PowerShell processes (pwsh.exe or powershell.exe) and communicating via standard input/output streams.

**Architecture**: C# binary module (compiled to DLL) → loaded by PowerShell → exposes `New-PSHostSession` cmdlet

## Key Components

- [`src/PSHostSession.cs`](../src/PSHostSession.cs) - Core cmdlet implementation with custom transport manager
  - `NewPSHostSessionCommand` - The main cmdlet that creates PSSessions
  - `PSHostClientSessionTransportMgr` - Custom transport using Process stdio instead of network
  - `PSHostClientInfo` - Connection info storing executable path and arguments
- [`src/PowerShellFinder.cs`](../src/PowerShellFinder.cs) - Logic to locate PowerShell executables in PATH
- [`AwakeCoding.PSRemoting/AwakeCoding.PSRemoting.psd1`](../AwakeCoding.PSRemoting/AwakeCoding.PSRemoting.psd1) - Module manifest defining metadata and exports

## Build & Development Workflow

Build the module using the provided script:
```powershell
.\build.ps1
```

This script:
1. Compiles the C# project: `dotnet build -c Release -f net9.0`
2. Copies the built DLL from `src/bin/Release/net9.0/` to `AwakeCoding.PSRemoting/` folder

**Testing locally**: After building, import the module from the workspace:
```powershell
Import-Module .\AwakeCoding.PSRemoting\AwakeCoding.PSRemoting.psd1
New-PSHostSession | Enter-PSSession
```

## Project Conventions

- **Target Framework**: .NET 9.0 (net9.0)
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

## Common Patterns

When adding new cmdlet parameters:
- Add as properties to `NewPSHostSessionCommand` class
- Mark with `[Parameter()]` attribute
- Use `[ValidateNotNullOrEmpty()]` for mandatory string parameters
- Access via properties in `BeginProcessing()` method

When modifying transport behavior:
- Edit `PSHostClientSessionTransportMgr.CreateAsync()` for connection setup
- Override `CleanupConnection()` to ensure proper resource disposal
- Remember to handle both normal and error data streams from Process

## Dependencies

- `System.Management.Automation` (PowerShell SDK 7.x) - Required for cmdlet development
- No network/remoting dependencies - this is entirely local process communication
