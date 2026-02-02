function New-ModulePackage
{
    [CmdletBinding()]
	param(
        [Parameter(Mandatory=$true,Position=0)]
        [string] $InputPath,
        [Parameter(Mandatory=$true,Position=1)]
        [string] $OutputPath,
        [string] $TempPath
    )

    $UniqueId = New-Guid

    if ([string]::IsNullOrEmpty($TempPath)) {
        $TempPath = [System.IO.Path]::GetTempPath()
    }

    $PSRepoName = "psrepo-$UniqueId"
    $PSRepoPath = Join-Path $TempPath $UniqueId

    if (-Not (Test-Path -Path $InputPath -PathType 'Container')) {
        throw "`"$InputPath`" does not exist"
    }

    $PSModulePath = $InputPath
    $PSManifestFile = $(@(Get-ChildItem -Path $PSModulePath -Depth 1 -Filter "*.psd1")[0]).FullName
    $PSManifest = Import-PowerShellDataFile -Path $PSManifestFile
    $PSModuleName = $(Get-Item $PSManifestFile).BaseName
    $PSModuleVersion = $PSManifest.ModuleVersion

    # https://docs.microsoft.com/en-us/nuget/concepts/package-versioning#normalized-version-numbers
    $NugetVersion = $PSModuleVersion -Replace "^(\d+)\.(\d+)\.(\d+)(\.0)$", "`$1.`$2.`$3"

    New-Item -Path $PSRepoPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null

    $Params = @{
        Name = $PSRepoName;
        SourceLocation = $PSRepoPath;
        PublishLocation = $PSRepoPath;
        InstallationPolicy = "Trusted";
    }

    Register-PSRepository @Params | Out-Null

    $OutputFileName = "${PSModuleName}.${NugetVersion}.nupkg"
    $PSModulePackage = Join-Path $PSRepoPath $OutputFileName
    Remove-Item -Path $PSModulePackage -ErrorAction 'SilentlyContinue'
    Publish-Module -Path $PSModulePath -Repository $PSRepoName

    Unregister-PSRepository -Name $PSRepoName | Out-Null

    New-Item -Path $OutputPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
    $OutputFile = Join-Path $OutputPath $OutputFileName
    Copy-Item $PSModulePackage $OutputFile

    Remove-Item $PSmodulePackage
    Remove-Item -Path $PSRepoPath

    $OutputFile
}
