# SSH Test Helper Module
# Provides Docker-based SSH server for testing

$script:TestContainerName = "psssh-test-$(Get-Random)"
$script:TestImageName = "awakecoding/psssh-test"

function Test-DockerAvailable {
    <#
    .SYNOPSIS
    Check if Docker is available and running
    #>
    try {
        # Check if docker command exists
        $dockerCmd = Get-Command docker -ErrorAction SilentlyContinue
        if (-not $dockerCmd) {
            return $false
        }
        
        # Try to run docker version and check output
        $output = docker version 2>&1 | Out-String
        
        # Check if we got valid output (contains "Version:")
        return $output -match 'Version:'
    }
    catch {
        return $false
    }
}

function Build-SSHTestImage {
    <#
    .SYNOPSIS
    Build the Docker image for SSH testing
    #>
    [CmdletBinding()]
    param(
        [switch]$Force
    )

    $dockerfilePath = Join-Path $PSScriptRoot "docker"
    
    # Check if image already exists
    if (-not $Force) {
        $existingImage = docker images -q $script:TestImageName 2>$null
        if ($existingImage) {
            Write-Verbose "Image $script:TestImageName already exists"
            return $true
        }
    }

    Write-Host "Building SSH test Docker image..." -ForegroundColor Cyan
    
    $buildOutput = docker build -t $script:TestImageName $dockerfilePath 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to build Docker image: $buildOutput"
        return $false
    }

    Write-Host "Docker image built successfully" -ForegroundColor Green
    return $true
}

function Start-SSHTestContainer {
    <#
    .SYNOPSIS
    Start the SSH test container
    #>
    [CmdletBinding()]
    param(
        [int]$Port = 0  # 0 = find available port
    )

    # Ensure image is built
    if (-not (Build-SSHTestImage)) {
        throw "Failed to build Docker image"
    }

    # Find available port if not specified
    if ($Port -eq 0) {
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, 0)
        $listener.Start()
        $Port = $listener.LocalEndpoint.Port
        $listener.Stop()
    }

    # Remove any existing container with same name
    $existing = docker ps -a -q -f "name=$script:TestContainerName" 2>$null
    if ($existing) {
        docker rm -f $script:TestContainerName 2>&1 | Out-Null
    }

    Write-Host "Starting SSH test container on port $Port..." -ForegroundColor Cyan

    # Start container
    $containerId = docker run -d `
        --name $script:TestContainerName `
        -p "${Port}:22" `
        $script:TestImageName 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to start container: $containerId"
        return $null
    }

    # Wait for SSH to be ready
    $maxAttempts = 10
    $attempt = 0
    $ready = $false

    Write-Host "Waiting for SSH server to be ready..." -ForegroundColor Cyan

    while ($attempt -lt $maxAttempts -and -not $ready) {
        Start-Sleep -Seconds 1
        
        # Check if container is running
        $containerStatus = docker inspect -f '{{.State.Running}}' $script:TestContainerName 2>$null
        if ($containerStatus -eq 'true') {
            # Try to connect to SSH port
            try {
                $tcpClient = [System.Net.Sockets.TcpClient]::new()
                $tcpClient.Connect('127.0.0.1', $Port)
                $tcpClient.Close()
                $ready = $true
            }
            catch {
                $attempt++
            }
        }
        else {
            break
        }
    }

    if (-not $ready) {
        Write-Error "SSH server did not become ready within timeout"
        docker logs $script:TestContainerName
        Stop-SSHTestContainer
        return $null
    }

    Write-Host "SSH test container ready" -ForegroundColor Green

    return [PSCustomObject]@{
        ContainerId = $containerId.Trim()
        ContainerName = $script:TestContainerName
        Port = $Port
        Host = 'localhost'
        UserName = 'testuser'
        Password = 'testpass'
    }
}

function Stop-SSHTestContainer {
    <#
    .SYNOPSIS
    Stop and remove the SSH test container
    #>
    [CmdletBinding()]
    param()

    $existing = docker ps -a -q -f "name=$script:TestContainerName" 2>$null
    if ($existing) {
        Write-Host "Stopping SSH test container..." -ForegroundColor Cyan
        docker stop $script:TestContainerName 2>&1 | Out-Null
        docker rm $script:TestContainerName 2>&1 | Out-Null
        Write-Host "Container stopped and removed" -ForegroundColor Green
    }
}

function Get-SSHTestCredential {
    <#
    .SYNOPSIS
    Get credentials for SSH test server
    #>
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Server
    )

    return [PSCredential]::new(
        $Server.UserName,
        (ConvertTo-SecureString $Server.Password -AsPlainText -Force)
    )
}

Export-ModuleMember -Function @(
    'Test-DockerAvailable',
    'Build-SSHTestImage',
    'Start-SSHTestContainer',
    'Stop-SSHTestContainer',
    'Get-SSHTestCredential'
)
