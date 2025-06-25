# AwakeCoding PSRemoting Extensions

## Installation

```powershell
Install-Module AwakeCoding.PSRemoting
```

## New-PSHostSession usage

The `New-PSHostSession` cmdlet creates a [PSSession object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_pssessions) connected to a PowerShell host subprocess:

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
PS > $PSSession | Enter-PSession
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

## How does New-PSHostSession work?

So, what does `New-PSHostSession` even do? It creates a PowerShell subprocess in *server mode*, meaning it will expect the [PSRemoting protocol](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-psrp/58ff7ff6-8078-45e2-a6ff-d496e188c39a) over standard input/output instead the regular command-line interface. The cmdlet then does the internal plumbing to connect the PSRemoting client to the standard input/output streams of the subprocess, the same way it would be done for [named pipe transports](https://awakecoding.com/posts/powershell-host-ipc-for-any-dotnet-application/).

This is similar to [Start-Job](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/start-job) except that it works with the PowerShell SDK and it returns a [PSSession object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_pssessions) instead of a [Job object](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_jobs). The interface is different, but if you're used to PSRemoting scenarios, it feels more natural.

## Why use New-PSHostSession at all?

Process isolation to avoid assembly loading conflicts, especially inside a parent .NET process which uses Microsoft Graph or Azure-related assemblies. Using WinRM on localhost can provide a similar result, but it requires configuring the WinRM listener, and you may hit problems with NTLM. There is obviously an overhead involved in launching a complete new PowerShell process to run something isolated, but sometimes it's the best (and only) solution to the problem.

## Building and running

Run the build.ps1 script to build the C# project, then launch a new PowerShell process with the local copy of the module loaded:

```powershell
.\build.ps1 && pwsh -Command { Import-Module .\AwakeCoding.PSRemoting } -NoExit
```

Since this is a binary PowerShell module, it cannot be unloaded or reloaded, which is why you need to launch a new PowerShell process every time.
