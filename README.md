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
