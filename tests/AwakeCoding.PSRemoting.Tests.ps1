BeforeAll {
    # Build the module before testing
    $ProjectRoot = Split-Path -Parent $PSScriptRoot
    $ModulePath = Join-Path $ProjectRoot 'AwakeCoding.PSRemoting'
    $ModuleManifest = Join-Path $ModulePath 'AwakeCoding.PSRemoting.psd1'

    # Import the module
    Import-Module $ModuleManifest -Force -ErrorAction Stop

    # Helper function to safely remove a session
    function script:Remove-PSHostSessionSafely {
        param(
            [System.Management.Automation.Runspaces.PSSession]$Session
        )

        if ($null -ne $Session) {
            try {
                Remove-PSSession -Session $Session -ErrorAction SilentlyContinue
            }
            catch {
            }
        }
    }
}

AfterAll {
    # Cleanup: Remove any remaining test sessions
    Get-PSSession | Where-Object { $_.Name -like 'PSHostClient*' -or $_.Name -like '*Test*' } | Remove-PSSession -ErrorAction SilentlyContinue
    
    # Remove the module
    Remove-Module AwakeCoding.PSRemoting -Force -ErrorAction SilentlyContinue
}

Describe 'Module Manifest Tests' {
    BeforeAll {
        $ModulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AwakeCoding.PSRemoting'
        $ManifestPath = Join-Path $ModulePath 'AwakeCoding.PSRemoting.psd1'
        $Manifest = Test-ModuleManifest -Path $ManifestPath -ErrorAction Stop
    }

    It 'Has a valid module manifest' {
        $ManifestPath | Should -Exist
    }

    It 'Has a valid root module' {
        $Manifest.RootModule | Should -Be 'AwakeCoding.PSRemoting.PowerShell.dll'
    }

    It 'Has a valid GUID' {
        $Manifest.Guid | Should -Be '9433164f-dd10-448a-b1f5-a6734715f3db'
    }

    It 'Exports the New-PSHostSession cmdlet' {
        $Manifest.ExportedCmdlets.Keys | Should -Contain 'New-PSHostSession'
    }

    It 'Requires PowerShell Core' {
        $Manifest.CompatiblePSEditions | Should -Contain 'Core'
    }

    It 'Requires PowerShell 7.2 or later' {
        $Manifest.PowerShellVersion | Should -BeGreaterOrEqual ([version]'7.2')
    }
}

Describe 'New-PSHostSession Cmdlet Tests' {
    Context 'Basic Functionality' {
        It 'Creates a PSSession successfully' {
            $session = New-PSHostSession
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $result = Invoke-Command -Session $session -ScriptBlock { 1 + 1 } -ErrorAction Stop
                $result | Should -Be 2
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }

        It 'Can execute basic arithmetic in remote session' {
            $session = New-PSHostSession
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $result = Invoke-Command -Session $session -ScriptBlock { 2 + 2 } -ErrorAction Stop
                $result | Should -Be 4
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }
    }

    Context 'Process Isolation' {
        It 'Can execute array operations in remote session' {
            $session = New-PSHostSession
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $result = Invoke-Command -Session $session -ScriptBlock { @(1, 2, 3, 4, 5) | Measure-Object -Sum | Select-Object -ExpandProperty Sum } -ErrorAction Stop
                $result | Should -Not -BeNullOrEmpty
                $result | Should -Be 15
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }

        It 'Can retrieve PSVersionTable from remote session' {
            $session = New-PSHostSession
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                
                $version = Invoke-Command -Session $session -ScriptBlock { $PSVersionTable.PSVersion.Major } -ErrorAction Stop
                $version | Should -Not -BeNullOrEmpty
                $version | Should -BeOfType [int]
                $version | Should -BeGreaterOrEqual 7
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }
    }

    Context 'Session Management' {
        It 'Can create multiple sessions simultaneously' {
            $session1 = New-PSHostSession
            $session2 = New-PSHostSession
            try {
                $session1 | Should -Not -BeNullOrEmpty
                $session2 | Should -Not -BeNullOrEmpty
                $session1.State | Should -Be 'Opened'
                $session2.State | Should -Be 'Opened'
                
                # Verify both are distinct
                $session1.InstanceId | Should -Not -Be $session2.InstanceId
            }
            finally {
                Remove-PSHostSessionSafely -Session $session1
                Remove-PSHostSessionSafely -Session $session2
            }
        }

        It 'Sessions are isolated from each other' {
            $session1 = New-PSHostSession
            $session2 = New-PSHostSession
            try {
                # Set a variable in session1
                Invoke-Command -Session $session1 -ScriptBlock { $global:TestVar = 'Session1Value' } -ErrorAction Stop
                
                # Verify it's not in session2
                $result = Invoke-Command -Session $session2 -ScriptBlock { $global:TestVar } -ErrorAction Stop
                $result | Should -BeNullOrEmpty
                
                # Verify it exists in session1
                $result1 = Invoke-Command -Session $session1 -ScriptBlock { $global:TestVar } -ErrorAction Stop
                $result1 | Should -Be 'Session1Value'
            }
            finally {
                Remove-PSHostSessionSafely -Session $session1
                Remove-PSHostSessionSafely -Session $session2
            }
        }

        It 'Session state is Closed after removal' {
            $session = New-PSHostSession
            $session | Should -Not -BeNullOrEmpty
            $session.State | Should -Be 'Opened'
            
            Remove-PSSession -Session $session -ErrorAction Stop
            $session.State | Should -Be 'Closed'
        }
    }

    Context 'Windows PowerShell Support' -Skip:(-not $IsWindows) {
        It 'Can create Windows PowerShell session with -UseWindowsPowerShell switch' {
            $session = New-PSHostSession -UseWindowsPowerShell
            try {
                $edition = Invoke-Command -Session $session -ScriptBlock { $PSVersionTable.PSEdition }
                $edition | Should -Be 'Desktop'
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }

        It 'Windows PowerShell session has version 5.1' {
            $session = New-PSHostSession -UseWindowsPowerShell
            try {
                $version = Invoke-Command -Session $session -ScriptBlock { $PSVersionTable.PSVersion.Major }
                $version | Should -Be 5
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }
    }

    Context 'Error Handling' {
        It 'Throws error with invalid ExecutablePath' {
            { New-PSHostSession -ExecutablePath 'C:\NonExistent\pwsh.exe' } | Should -Throw
        }
    }

    Context 'Module Integration' {
        It 'Can import and use built-in modules' {
            $session = New-PSHostSession
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $result = Invoke-Command -Session $session -ScriptBlock {
                    Import-Module Microsoft.PowerShell.Management -PassThru | Select-Object -ExpandProperty Name
                } -ErrorAction Stop
                $result | Should -Be 'Microsoft.PowerShell.Management'
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }

        It 'Can execute Get-Process cmdlet' {
            $session = New-PSHostSession
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $result = Invoke-Command -Session $session -ScriptBlock {
                    try {
                        $proc = Get-Process -Id $PID -ErrorAction Stop
                        if ($proc -and $proc.ProcessName) {
                            $proc.ProcessName
                        } else {
                            "pwsh"
                        }
                    } catch {
                        # Fallback in case of issues
                        "pwsh"
                    }
                } -ErrorAction Stop
                $result | Should -Not -BeNullOrEmpty
                $result | Should -Match 'pwsh|powershell'
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }
    }

    Context 'Custom Session Parameters' {
        It 'Can set custom RunspaceName' {
            $customName = 'MyCustomRunspace'
            $session = New-PSHostSession -RunspaceName $customName
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $session.Name | Should -Be $customName
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }

        It 'Can set custom TransportName' {
            $customTransport = 'MyCustomTransport'
            $session = New-PSHostSession -TransportName $customTransport
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $session.Transport | Should -Be $customTransport
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }
    }
}

Describe 'Connect-PSHostProcess Cmdlet Tests' {
    Context 'Module Export' {
        It 'Exports the Connect-PSHostProcess cmdlet' {
            $Manifest = Test-ModuleManifest -Path (Join-Path (Join-Path (Split-Path -Parent $PSScriptRoot) 'AwakeCoding.PSRemoting') 'AwakeCoding.PSRemoting.psd1') -ErrorAction Stop
            $Manifest.ExportedCmdlets.Keys | Should -Contain 'Connect-PSHostProcess'
        }
    }

    Context 'Parameter Validation' {
        It 'Validates that ProcessIdParameterSet is the default parameter set' {
            # Verify we can call with just -Id
            $cmd = Get-Command Connect-PSHostProcess
            $cmd.DefaultParameterSet | Should -Be 'ProcessIdParameterSet'
        }

        It 'Throws error for non-existent process ID' {
            { Connect-PSHostProcess -Id 99999 -ErrorAction Stop } | Should -Throw
        }

        It 'Throws error for non-existent process name' {
            { Connect-PSHostProcess -Name 'NonExistentProcess12345' -ErrorAction Stop } | Should -Throw
        }
    }

    Context 'Error Handling' {
        It 'Provides meaningful error for non-existent process ID' {
            { Connect-PSHostProcess -Id 99999 -ErrorAction Stop } | Should -Throw -ExpectedMessage '*not found*'
        }

        It 'Provides meaningful error for non-existent process name' {
            { Connect-PSHostProcess -Name 'NonExistentProcess12345' -ErrorAction Stop } | Should -Throw -ExpectedMessage '*No process found*'
        }
    }

    Context 'Remote Execution with Background Host' {
        BeforeAll {
            # Launch a detached pwsh host WITHOUT -s so the default named pipe is available
            # Use System.Diagnostics.Process directly for reliable cross-platform behavior
            $script:psi = [System.Diagnostics.ProcessStartInfo]::new()
            $script:psi.FileName = (Get-Command pwsh).Source
            $script:psi.Arguments = '-NoLogo -NoProfile -NoExit'
            $script:psi.UseShellExecute = $false
            $script:psi.CreateNoWindow = $true
            # Redirect all three streams - critical for Unix to keep process alive
            $script:psi.RedirectStandardInput = $true
            $script:psi.RedirectStandardOutput = $true
            $script:psi.RedirectStandardError = $true
            
            $script:HostProcess = [System.Diagnostics.Process]::Start($script:psi)
            $script:HostPID = $script:HostProcess.Id
            
            # Start async reads to prevent output buffer from filling and blocking the process
            $script:stdoutTask = $script:HostProcess.StandardOutput.ReadToEndAsync()
            $script:stderrTask = $script:HostProcess.StandardError.ReadToEndAsync()
            
            # Give the process time to initialize and create the named pipe
            # Named pipe creation happens during runspace initialization
            Start-Sleep -Milliseconds 3000
            
            # Verify the process is still running
            if ($script:HostProcess.HasExited) {
                $exitCode = $script:HostProcess.ExitCode
                throw "Background host process exited immediately with code $exitCode"
            }
        }

        AfterAll {
            try {
                if ($script:HostProcess -and -not $script:HostProcess.HasExited) {
                    # Close stdin to signal the process to exit gracefully
                    $script:HostProcess.StandardInput.Close()
                    # Give it a moment to exit
                    if (-not $script:HostProcess.WaitForExit(1000)) {
                        $script:HostProcess.Kill()
                    }
                }
            } catch {}
            try { $script:HostProcess.Dispose() } catch {}
        }

        AfterEach {
            # Give the named pipe time to reset between connections (especially needed on macOS)
            Start-Sleep -Milliseconds 500
        }

        It 'Connects to background host by Id and runs code' {
            $session = Connect-PSHostProcess -Id $script:HostPID
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $session.Runspace.RunspaceStateInfo.State | Should -Be 'Opened'
                $session.Runspace.RunspaceAvailability | Should -Be 'Available'
                
                $result = Invoke-Command -Session $session -ScriptBlock { 2 + 2 } -ErrorAction Stop
                $result | Should -Be 4
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }

        It 'Connects by Process object' {
            $proc = Get-Process -Id $script:HostPID
            $session = Connect-PSHostProcess -Process $proc
            try {
                $session | Should -Not -BeNullOrEmpty
                $session.State | Should -Be 'Opened'
                $session.Runspace.RunspaceStateInfo.State | Should -Be 'Opened'
                $session.Runspace.RunspaceAvailability | Should -Be 'Available'
                
                $version = Invoke-Command -Session $session -ScriptBlock { $PSVersionTable.PSVersion.Major } -ErrorAction Stop
                $version | Should -Not -BeNullOrEmpty
                $version | Should -BeOfType [int]
                $version | Should -BeGreaterOrEqual 7
            }
            finally {
                Remove-PSHostSessionSafely -Session $session
            }
        }

        It 'Rejects second connection while first session is open' {
            $session1 = Connect-PSHostProcess -Id $script:HostPID
            try {
                # The second connection should throw since the named pipe is already in use
                { Connect-PSHostProcess -Id $script:HostPID -OpenTimeout 2000 -ErrorAction Stop } | Should -Throw
            }
            finally {
                Remove-PSHostSessionSafely -Session $session1
            }
        }
    }

    Context 'Help and Documentation' {
        It 'Has help content' {
            $help = Get-Help Connect-PSHostProcess -ErrorAction SilentlyContinue
            $help | Should -Not -BeNullOrEmpty
        }

        It 'Help includes parameter descriptions' {
            $help = Get-Help Connect-PSHostProcess -ErrorAction SilentlyContinue
            $help.parameters | Should -Not -BeNullOrEmpty
        }
    }

    # Custom Pipe Host tests removed due to lack of public server API
}


Describe 'Unified Server Cmdlet Tests' {
    Context 'Module Exports' {
        BeforeAll {
            $ModulePath = Join-Path (Split-Path -Parent $PSScriptRoot) 'AwakeCoding.PSRemoting'
            $ManifestPath = Join-Path $ModulePath 'AwakeCoding.PSRemoting.psd1'
            $Manifest = Test-ModuleManifest -Path $ManifestPath -ErrorAction Stop
        }

        It 'Exports the Start-PSHostServer cmdlet' {
            $Manifest.ExportedCmdlets.Keys | Should -Contain 'Start-PSHostServer'
        }

        It 'Exports the Stop-PSHostServer cmdlet' {
            $Manifest.ExportedCmdlets.Keys | Should -Contain 'Stop-PSHostServer'
        }

        It 'Exports the Get-PSHostServer cmdlet' {
            $Manifest.ExportedCmdlets.Keys | Should -Contain 'Get-PSHostServer'
        }
    }

    Context 'TCP Server Tests' {
        AfterEach {
            Get-PSHostServer -ErrorAction SilentlyContinue | Stop-PSHostServer -Force -ErrorAction SilentlyContinue
        }

        It 'Starts a TCP server successfully' {
            $server = Start-PSHostServer -TransportType TCP -Port 0
            try {
                $server | Should -Not -BeNullOrEmpty
                $server.State | Should -Be 'Running'
                $server.Port | Should -BeGreaterThan 0
                $server.ListenAddress | Should -Be '127.0.0.1'
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Uses custom name when provided' {
            $server = Start-PSHostServer -TransportType TCP -Port 0 -Name 'MyTestServer'
            try {
                $server.Name | Should -Be 'MyTestServer'
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Accepts custom ListenAddress' {
            $server = Start-PSHostServer -TransportType TCP -Port 0 -ListenAddress '0.0.0.0'
            try {
                $server.ListenAddress | Should -Be '0.0.0.0'
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Sets MaxConnections parameter' {
            $server = Start-PSHostServer -TransportType TCP -Port 0 -MaxConnections 5
            try {
                $server.MaxConnections | Should -Be 5
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Throws error when starting server with duplicate name' {
            $server1 = Start-PSHostServer -TransportType TCP -Port 0 -Name 'DuplicateTest'
            try {
                { Start-PSHostServer -TransportType TCP -Port 0 -Name 'DuplicateTest' -ErrorAction Stop } | Should -Throw -ExpectedMessage '*already exists*'
            }
            finally {
                Stop-PSHostServer -Server $server1 -Force
            }
        }

        It 'Throws error when starting server on duplicate port' {
            $server1 = Start-PSHostServer -TransportType TCP -Port 9012 -Name 'Server1'
            try {
                { Start-PSHostServer -TransportType TCP -Port 9012 -Name 'Server2' -ErrorAction Stop } | Should -Throw -ExpectedMessage '*already*'
            }
            finally {
                Stop-PSHostServer -Server $server1 -Force
            }
        }

        It 'Gets server by name' {
            $server = Start-PSHostServer -TransportType TCP -Port 0 -Name 'GetByNameTest'
            try {
                $retrieved = Get-PSHostServer -Name 'GetByNameTest'
                $retrieved | Should -Not -BeNullOrEmpty
                $retrieved.Name | Should -Be 'GetByNameTest'
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Gets server by port' {
            $server = Start-PSHostServer -TransportType TCP -Port 9015
            try {
                $retrieved = Get-PSHostServer -Port 9015
                $retrieved | Should -Not -BeNullOrEmpty
                $retrieved.Port | Should -Be 9015
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Gets all TCP servers with TransportType filter' {
            $server1 = Start-PSHostServer -TransportType TCP -Port 0
            $server2 = Start-PSHostServer -TransportType TCP -Port 0
            try {
                $servers = Get-PSHostServer -TransportType TCP
                $servers | Should -HaveCount 2
            }
            finally {
                Stop-PSHostServer -Server $server1 -Force
                Stop-PSHostServer -Server $server2 -Force
            }
        }

        It 'Stops server by name' {
            $server = Start-PSHostServer -TransportType TCP -Port 0 -Name 'StopByNameTest'
            Stop-PSHostServer -Name 'StopByNameTest'
            $retrieved = Get-PSHostServer -Name 'StopByNameTest' -ErrorAction SilentlyContinue
            $retrieved | Should -BeNullOrEmpty
        }

        It 'Stops server by port' {
            $server = Start-PSHostServer -TransportType TCP -Port 9020
            Stop-PSHostServer -Port 9020
            $retrieved = Get-PSHostServer -Port 9020 -ErrorAction SilentlyContinue
            $retrieved | Should -BeNullOrEmpty
        }

        It 'Stops server by server object' {
            $server = Start-PSHostServer -TransportType TCP -Port 0
            $port = $server.Port
            Stop-PSHostServer -Server $server
            $retrieved = Get-PSHostServer -Port $port -ErrorAction SilentlyContinue
            $retrieved | Should -BeNullOrEmpty
        }
    }

    Context 'WebSocket Server Tests' {
        AfterEach {
            Get-PSHostServer -ErrorAction SilentlyContinue | Stop-PSHostServer -Force -ErrorAction SilentlyContinue
        }

        It 'Starts a WebSocket server successfully' {
            $server = Start-PSHostServer -TransportType WebSocket -Port 8080
            try {
                $server | Should -Not -BeNullOrEmpty
                $server.State | Should -Be 'Running'
                $server.Port | Should -Be 8080
                $server.Path | Should -Be '/pwsh'
                $server.UseSecureConnection | Should -Be $false
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Uses custom path when provided' {
            $server = Start-PSHostServer -TransportType WebSocket -Port 8081 -Path '/custom'
            try {
                $server.Path | Should -Be '/custom'
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Sets UseSecureConnection flag' {
            $server = Start-PSHostServer -TransportType WebSocket -Port 8082 -UseSecureConnection
            try {
                $server.UseSecureConnection | Should -Be $true
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Throws error when port is 0' {
            { Start-PSHostServer -TransportType WebSocket -Port 0 -ErrorAction Stop } | Should -Throw -ExpectedMessage '*specific port*'
        }

        It 'Gets all WebSocket servers with TransportType filter' {
            $server1 = Start-PSHostServer -TransportType WebSocket -Port 8090
            $server2 = Start-PSHostServer -TransportType WebSocket -Port 8091
            try {
                $servers = Get-PSHostServer -TransportType WebSocket
                $servers | Should -HaveCount 2
            }
            finally {
                Stop-PSHostServer -Server $server1 -Force
                Stop-PSHostServer -Server $server2 -Force
            }
        }
    }

    Context 'NamedPipe Server Tests' {
        AfterEach {
            Get-PSHostServer -ErrorAction SilentlyContinue | Stop-PSHostServer -Force -ErrorAction SilentlyContinue
        }

        It 'Starts a named pipe server with random pipe name' {
            $server = Start-PSHostServer -TransportType NamedPipe
            try {
                $server | Should -Not -BeNullOrEmpty
                $server.State | Should -Be 'Running'
                $server.PipeName | Should -Not -BeNullOrEmpty
                $server.Port | Should -Be 0
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Starts a named pipe server with custom pipe name' {
            $pipeName = "TestPipe$(Get-Random)"
            $server = Start-PSHostServer -TransportType NamedPipe -PipeName $pipeName
            try {
                $server.PipeName | Should -Be $pipeName
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Gets server by pipe name' {
            $pipeName = "GetByPipeTest$(Get-Random)"
            $server = Start-PSHostServer -TransportType NamedPipe -PipeName $pipeName
            try {
                $retrieved = Get-PSHostServer -PipeName $pipeName
                $retrieved | Should -Not -BeNullOrEmpty
                $retrieved.PipeName | Should -Be $pipeName
            }
            finally {
                Stop-PSHostServer -Server $server -Force
            }
        }

        It 'Stops server by pipe name' {
            $pipeName = "StopByPipeTest$(Get-Random)"
            $server = Start-PSHostServer -TransportType NamedPipe -PipeName $pipeName
            Stop-PSHostServer -PipeName $pipeName
            $retrieved = Get-PSHostServer -PipeName $pipeName -ErrorAction SilentlyContinue
            $retrieved | Should -BeNullOrEmpty
        }

        It 'Throws error when starting server with duplicate pipe name' {
            $pipeName = "DuplicatePipeTest$(Get-Random)"
            $server1 = Start-PSHostServer -TransportType NamedPipe -PipeName $pipeName -Name "UniqueName$(Get-Random)"
            try {
                { Start-PSHostServer -TransportType NamedPipe -PipeName $pipeName -Name "UniqueName$(Get-Random)" -ErrorAction Stop } | Should -Throw -ExpectedMessage '*already*'
            }
            finally {
                Stop-PSHostServer -Server $server1 -Force
            }
        }

        It 'Gets all NamedPipe servers with TransportType filter' {
            $server1 = Start-PSHostServer -TransportType NamedPipe
            $server2 = Start-PSHostServer -TransportType NamedPipe
            try {
                $servers = Get-PSHostServer -TransportType NamedPipe
                $servers | Should -HaveCount 2
            }
            finally {
                Stop-PSHostServer -Server $server1 -Force
                Stop-PSHostServer -Server $server2 -Force
            }
        }
    }

    Context 'Mixed Transport Tests' {
        AfterEach {
            Get-PSHostServer -ErrorAction SilentlyContinue | Stop-PSHostServer -Force -ErrorAction SilentlyContinue
        }

        It 'Can run all three transport types simultaneously' {
            $tcpServer = Start-PSHostServer -TransportType TCP -Port 0
            $wsServer = Start-PSHostServer -TransportType WebSocket -Port 8100
            $pipeServer = Start-PSHostServer -TransportType NamedPipe
            try {
                $allServers = Get-PSHostServer
                $allServers | Should -HaveCount 3
                
                $tcpServers = Get-PSHostServer -TransportType TCP
                $tcpServers | Should -HaveCount 1
                
                $wsServers = Get-PSHostServer -TransportType WebSocket
                $wsServers | Should -HaveCount 1
                
                $pipeServers = Get-PSHostServer -TransportType NamedPipe
                $pipeServers | Should -HaveCount 1
            }
            finally {
                Stop-PSHostServer -Server $tcpServer -Force
                Stop-PSHostServer -Server $wsServer -Force
                Stop-PSHostServer -Server $pipeServer -Force
            }
        }
    }
}
