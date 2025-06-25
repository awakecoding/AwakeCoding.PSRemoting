
dotnet build -c Release -f net9.0

$OutputPath = "$PSScriptRoot\src\bin\Release\net9.0"
$ModulePath = "$PSScriptRoot\AwakeCoding.PSRemoting"
Copy-Item "$OutputPath\AwakeCoding.PSRemoting.PowerShell.dll" "$ModulePath\AwakeCoding.PSRemoting.PowerShell.dll" -Force
