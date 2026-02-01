
dotnet build -c Release -f net8.0

$OutputPath = "$PSScriptRoot\src\bin\Release\net8.0"
$ModulePath = "$PSScriptRoot\AwakeCoding.PSRemoting"
Copy-Item "$OutputPath\AwakeCoding.PSRemoting.PowerShell.dll" "$ModulePath\AwakeCoding.PSRemoting.PowerShell.dll" -Force
