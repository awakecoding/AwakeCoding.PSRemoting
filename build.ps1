
dotnet build -c Release -f net8.0

$OutputPath = "$PSScriptRoot\src\bin\Release\net8.0"
$ModulePath = "$PSScriptRoot\AwakeCoding.PSRemoting"

# Ensure module path exists
New-Item -ItemType Directory -Path $ModulePath -Force | Out-Null

# Copy the module manifest
Copy-Item "$PSScriptRoot\src\AwakeCoding.PSRemoting.psd1" $ModulePath -Force

# Copy the main module DLL
Copy-Item "$OutputPath\AwakeCoding.PSRemoting.PowerShell.dll" $ModulePath -Force

# Copy required dependency DLLs (Tmds.Ssh and its dependencies)
$Dependencies = @(
    "Tmds.Ssh.dll",
    "BouncyCastle.Cryptography.dll",
    "Microsoft.Extensions.Logging.Abstractions.dll",
    "Microsoft.Extensions.DependencyInjection.Abstractions.dll"
)

foreach ($dll in $Dependencies) {
    $sourcePath = Join-Path $OutputPath $dll
    if (Test-Path $sourcePath) {
        Copy-Item $sourcePath $ModulePath -Force
        Write-Host "Copied: $dll"
    } else {
        Write-Warning "Dependency not found: $dll"
    }
}

Write-Host "`nModule files copied to: $ModulePath" -ForegroundColor Green
