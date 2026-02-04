# Build and test the module in a fresh PowerShell process

# Build the module first using build.ps1
Write-Host "Building module..." -ForegroundColor Cyan
& "$PSScriptRoot\build.ps1"

if ($LASTEXITCODE -ne 0) {
    Write-Error "Build failed with exit code $LASTEXITCODE"
    exit $LASTEXITCODE
}

Write-Host "Build completed successfully" -ForegroundColor Green
Write-Host ""
Write-Host "Running tests in fresh PowerShell process..." -ForegroundColor Cyan
Write-Host ""

# Run tests in a fresh pwsh.exe process
$testScript = @"
# Import the local module
Import-Module '$PSScriptRoot\AwakeCoding.PSRemoting\AwakeCoding.PSRemoting.psd1' -Force

# Check if Pester is available
if (-not (Get-Module -ListAvailable -Name Pester | Where-Object Version -ge 5.0.0)) {
    Write-Host 'Installing Pester 5.x...' -ForegroundColor Yellow
    Install-Module -Name Pester -MinimumVersion 5.0.0 -Force -SkipPublisherCheck -Scope CurrentUser
}

# Configure and run Pester tests
`$config = New-PesterConfiguration
`$config.Run.Path = '$PSScriptRoot\tests'
`$config.Run.Exit = `$true
`$config.Output.Verbosity = 'Detailed'
`$config.CodeCoverage.Enabled = `$false

Invoke-Pester -Configuration `$config
"@

pwsh -NoProfile -Command $testScript
exit $LASTEXITCODE
