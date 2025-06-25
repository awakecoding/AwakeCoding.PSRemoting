
dotnet build -c Release -f net9.0

Copy-Item "$PSScriptRoot\src\bin\Release\net9.0\AwakeCoding.PSRemoting.PowerShell.dll" "$PSScriptRoot\AwakeCoding.PSRemoting\AwakeCoding.PSRemoting.PowerShell.dll" -Force
