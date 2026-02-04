# AwakeCoding PSRemoting Extensions

PowerShell remoting without network transport - create local PowerShell sessions via subprocess stdio, and host remoting endpoints via TCP, WebSocket, or Named Pipes.

## Installation

```powershell
Install-Module AwakeCoding.PSRemoting
```

## Features

- **Client Sessions**: Create PSSession objects connected to local PowerShell subprocesses
- **Server Infrastructure**: Host PowerShell remoting endpoints on TCP, WebSocket, or Named Pipe transports
- **Process Connection**: Connect to existing PowerShell processes via named pipes
- **Cross-Platform**: Works on Windows, Linux, and macOS

## New-PSHostSession - Local Subprocess Sessions

The `New-PSHostSession` cmdlet creates a [PSSession object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_pssessions) connected to a PowerShell host subprocess:

## Enter-PSHostSession - Create and Enter Sessions Directly

The `Enter-PSHostSession` cmdlet combines session creation and interactive entry in a single command. It accepts the same parameters as `New-PSHostSession` but automatically enters the remote session:

```powershell
# Create subprocess session and enter it interactively
PS > Enter-PSHostSession
[localhost] PS > $PID
24228
[localhost] PS > exit
PS > $PID
19704
```

This is equivalent to `New-PSHostSession | Enter-PSSession` but more convenient. All transport types are supported:

```powershell
# Subprocess (default)
Enter-PSHostSession

# TCP connection
Enter-PSHostSession -HostName remote-server -Port 8080

# WebSocket connection
Enter-PSHostSession -Uri ws://localhost:8080/pwsh

# Named pipe connection
Enter-PSHostSession -PipeName MyPipe

# Connect to existing process
Enter-PSHostSession -ProcessId 12345

# SSH connection
Enter-PSHostSession -SSHTransport -HostName remote-server
```

## Creating Sessions Without Entering

Use `New-PSHostSession` when you want to create a session object for later use with `Invoke-Command` or other remoting cmdlets:

```powershell
PS > $PSSession = New-PSHostSession
PS > $PSSession

 Id Name            Transport ComputerName    ComputerType    State         ConfigurationName     Availability
 -- ----            --------- ------------    ------------    -----         -----------------     ------------
  1 PSHostClient    PSHostSe… localhost       RemoteMachine   Opened                                 Available
```

Confirm that execution happens in a *separate* process by comparing the process IDs between the client and server:

```powershell
PS > $PID
19704
PS > Invoke-Command -Session $PSSession { $PID }
24228
```

You can also enter the PowerShell host process to execute commands directly, then exit it to get back into the original process:

```powershell
PS > $PSSession | Enter-PSSession
[localhost] PS > $PID
24228
[localhost] PS > exit
PS > $PID
19704
```

PSSession objects are persistent, here is how you can list and cleanup the ones created by `New-PSHostSession`:

```powershell
Get-PSSession | Where-Object { $_.Name -eq 'PSHostClient' } | Remove-PSSession
```

## Windows PowerShell Compatibility

By default, `New-PSHostSession` starts PowerShell 7 (pwsh). To launch a session using Windows PowerShell instead, use the `-UseWindowsPowerShell` switch:

```powershell
PS > New-PSHostSession -UseWindowsPowerShell | Enter-PSSession
[localhost]: PS > $PSVersionTable

Name                           Value
----                           -----
PSVersion                      5.1.26100.4202
PSEdition                      Desktop
PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0...}
BuildVersion                   10.0.26100.4202
CLRVersion                     4.0.30319.42000
WSManStackVersion              3.0
PSRemotingProtocolVersion      2.3
SerializationVersion           1.1.0.1
```

For obvious reasons, this is only supported on Windows. You may have heard that PowerShell 7 can load Windows PowerShell modules through a [special compatibility layer](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_windows_powershell_compatibility), like this:

```powershell
PS > Import-Module ScheduledTasks -UseWindowsPowerShell
WARNING: Module ScheduledTasks is loaded in Windows PowerShell using WinPSCompatSession remoting session; please note that all input and output of commands from this module will be deserialized objects. If you want to load this module into PowerShell please use 'Import-Module -SkipEditionCheck' syntax.
```

You can achieve the same result manually by using the [Import-Module](https://learn.microsoft.com/en-us/powershell/scripting/developer/module/importing-a-powershell-module) `-PSSession` parameter:

```powershell
PS > $WinPSSession = New-PSHostSession -UseWindowsPowerShell
Import-Module -Name ScheduledTasks -PSSession $WinPSSession
```

Just like the built-in Windows compatibility, this results in automatic proxy functions created for the entire PowerShell module, such that when you call the cmdlets in the client process, they get executed through PSRemoting in a Windows PowerShell subprocess.

## Explicit PowerShell Path

The `New-PSHostSession` cmdlet finds the first PowerShell executable in the PATH, unless the parent executable is PowerShell, in which case it will use the same one. If you wish to launch a specific version of PowerShell, just specify the explicit executable path:

```powershell
PS > $PSSession = New-PSHostSession -ExecutablePath 'C:\PowerShell-7.2.11-win-x64\pwsh.exe'
PS > $PSSession | Enter-PSSession
[localhost]: PS > $PSVersionTable

Name                           Value
----                           -----
PSVersion                      7.2.11
PSEdition                      Core
GitCommitId                    7.2.11
OS                             Microsoft Windows 10.0.26100
Platform                       Win32NT
PSCompatibleVersions           {1.0, 2.0, 3.0, 4.0…}
PSRemotingProtocolVersion      2.3
SerializationVersion           1.1.0.1
WSManStackVersion              3.0
```

This can be useful to have a script running using the latest PowerShell that can drive executions in multiple older versions of PowerShell through local PSRemoting.

## SSH Transport - Remote PowerShell Over SSH

The `New-PSHostSession` cmdlet supports SSH transport for connecting to remote PowerShell instances over SSH. This provides an alternative to WinRM-based remoting that works across platforms.

### Basic SSH Connection

Connect to a remote PowerShell instance using SSH:

```powershell
# Basic SSH connection (defaults to port 22)
$session = New-PSHostSession -SSHTransport -HostName 'remote-server'

# Specify custom port
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' -Port 2222

# Enter the remote session
$session | Enter-PSSession
```

### Key-Based Authentication

Use SSH key files for authentication:

```powershell
# Connect using SSH private key
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' `
    -UserName 'admin' `
    -KeyFilePath '~/.ssh/id_rsa'
```

### Password Authentication

Use credentials for password-based authentication:

```powershell
# Prompt for credentials
$cred = Get-Credential -UserName 'admin'
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' -Credential $cred

# Interactive password prompt (if no credential provided)
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' -UserName 'admin'
```

### Advanced SSH Options

Customize SSH connection behavior:

```powershell
# Skip host key verification (useful for testing, not recommended for production)
$session = New-PSHostSession -SSHTransport -HostName 'test-server' -SkipHostKeyCheck

# Custom SSH subsystem (default is 'powershell')
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' `
    -Subsystem 'custom-subsystem'

# Connection timeout (in milliseconds, -1 for infinite)
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' `
    -ConnectingTimeout 10000

# Custom SSH options (passed directly to SSH)
$sshOptions = @{
    'StrictHostKeyChecking' = 'no'
    'UserKnownHostsFile' = '/dev/null'
}
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' -Options $sshOptions

# Use custom SSH executable path
$session = New-PSHostSession -SSHTransport -HostName 'remote-server' `
    -SSHExecutablePath '/usr/local/bin/ssh'
```

### SSH Requirements

To use SSH transport:
- **Remote machine** must have PowerShell installed and configured for SSH remoting
- **SSH server** must be configured with PowerShell as an SSH subsystem
- **Tmds.Ssh library** handles SSH protocol (included with module)

Configure PowerShell SSH subsystem on remote server:
```bash
# Linux/macOS - Add to /etc/ssh/sshd_config
Subsystem powershell /usr/bin/pwsh -sshs -NoLogo -NoProfile

# Windows - Add to sshd_config
Subsystem powershell C:\Program Files\PowerShell\7\pwsh.exe -sshs -NoLogo -NoProfile
```

## How does New-PSHostSession work?

So, what does `New-PSHostSession` even do? It creates a PowerShell subprocess in *server mode*, meaning it will expect the [PSRemoting protocol](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-psrp/58ff7ff6-8078-45e2-a6ff-d496e188c39a) over standard input/output instead the regular command-line interface. The cmdlet then does the internal plumbing to connect the PSRemoting client to the standard input/output streams of the subprocess, the same way it would be done for [named pipe transports](https://awakecoding.com/posts/powershell-host-ipc-for-any-dotnet-application/).

This is similar to [Start-Job](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/start-job) except that it works with the PowerShell SDK and it returns a [PSSession object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_pssessions) instead of a [Job object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_jobs). The interface is different, but if you're used to PSRemoting scenarios, it feels more natural.

## Why use New-PSHostSession at all?

Process isolation to avoid assembly loading conflicts, especially inside a parent .NET process which uses Microsoft Graph or Azure-related assemblies. Using WinRM on localhost can provide a similar result, but it requires configuring the WinRM listener, and you may hit problems with NTLM. There is obviously an overhead involved in launching a complete new PowerShell process to run something isolated, but sometimes it's the best (and only) solution to the problem.

## Server Infrastructure - Hosting Remoting Endpoints

The module provides unified server cmdlets to host PowerShell remoting endpoints using different transport types.

### Start-PSHostServer - TCP Transport

Start a TCP server on a specific port (or use port 0 for auto-assignment):

```powershell
# Start TCP server on port 8080
$server = Start-PSHostServer -TransportType TCP -Port 8080

# Connect from another PowerShell session
Enter-PSSession -HostName localhost -Port 8080
```

### Start-PSHostServer - WebSocket Transport

Start a WebSocket server with custom path and optional SSL:

```powershell
# Start WebSocket server
$server = Start-PSHostServer -TransportType WebSocket -Port 8081 -Path '/pwsh'

# With SSL (requires certificate configuration)
$server = Start-PSHostServer -TransportType WebSocket -Port 8082 -UseSecureConnection
```

### Start-PSHostServer - Named Pipe Transport

Start a named pipe server (generates random pipe name by default):

```powershell
# Start with auto-generated pipe name
$server = Start-PSHostServer -TransportType NamedPipe

# Start with custom pipe name
$server = Start-PSHostServer -TransportType NamedPipe -PipeName 'MyCustomPipe'
```

### Managing Servers

```powershell
# List all running servers
Get-PSHostServer

# Filter by transport type
Get-PSHostServer -TransportType TCP

# Stop server by name
Stop-PSHostServer -Name 'PSHostTcpServer8080'

# Stop server by port
Stop-PSHostServer -Port 8080

# Stop server by pipe name
Stop-PSHostServer -PipeName 'MyCustomPipe'

# Stop server forcefully (kills active connections immediately)
Stop-PSHostServer -Name 'MyServer' -Force
```

### Server Properties

Servers track state and connection information:

```powershell
$server = Start-PSHostServer -TransportType TCP -Port 8080

$server.State            # Running, Stopped, Starting, Stopping, Failed
$server.Port             # Listening port
$server.ConnectionCount  # Number of active connections
$server.Connections      # Array of connection details
$server.MaxConnections   # Maximum allowed connections (0 = unlimited)
$server.DrainTimeout     # Graceful shutdown timeout in seconds
```

## Connect-PSHostProcess - Attach to Running PowerShell

Connect to an existing PowerShell process via named pipes:

```powershell
# Connect by process ID
$session = Connect-PSHostProcess -Id 12345

# Connect by process name
$session = Connect-PSHostProcess -Name 'pwsh'

# Connect by Process object
$process = Get-Process -Name 'pwsh' | Select-Object -First 1
$session = Connect-PSHostProcess -Process $process

# Enter the connected session
$session | Enter-PSSession
```

## Building and running

Run the build.ps1 script to build the C# project:

```powershell
.\build.ps1
```

Run tests:

```powershell
.\test.ps1
```

Launch a new PowerShell process with the local copy of the module loaded:

```powershell
.\build.ps1 && pwsh -Command { Import-Module .\AwakeCoding.PSRemoting } -NoExit
```

Since this is a binary PowerShell module, it cannot be unloaded or reloaded, which is why you need to launch a new PowerShell process every time.
