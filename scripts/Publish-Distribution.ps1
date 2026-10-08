[CmdletBinding()]
param(
    [string]$OutputDirectory = "artifacts/distribution",
    [string]$DotnetPath = "dotnet"
)

$ErrorActionPreference = "Stop"
$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$projectPath = Join-Path $repositoryRoot "src/AutomaticTestPrinting.App/AutomaticTestPrinting.App.csproj"
$solutionPath = Join-Path $repositoryRoot "AutomaticTestPrinting.sln"
$readmePath = Join-Path $repositoryRoot "packaging/README-ja.txt"
$propertiesPath = Join-Path $repositoryRoot "Directory.Build.props"

[xml]$properties = [IO.File]::ReadAllText($propertiesPath, [Text.Encoding]::UTF8)
$version = [string]$properties.Project.PropertyGroup.Version
if ([string]::IsNullOrWhiteSpace($version)) {
    throw "Version is missing from Directory.Build.props."
}
$versionReadmePath = Join-Path $repositoryRoot "packaging/README-ja-v$version.txt"
if (Test-Path -LiteralPath $versionReadmePath) { $readmePath = $versionReadmePath }

if ([IO.Path]::IsPathRooted($OutputDirectory)) {
    $distributionRoot = [IO.Path]::GetFullPath($OutputDirectory)
}
else {
    $distributionRoot = [IO.Path]::GetFullPath((Join-Path $repositoryRoot $OutputDirectory))
}

$packageName = "AutomaticTestPrinting-v$version-win-x64"
$publishDirectory = Join-Path $repositoryRoot "artifacts/publish/$packageName"
$stageDirectory = Join-Path $distributionRoot $packageName
$zipPath = Join-Path $distributionRoot "$packageName.zip"
$checksumPath = Join-Path $distributionRoot "$packageName.sha256.txt"

& $DotnetPath test $solutionPath -c Release
if ($LASTEXITCODE -ne 0) {
    throw "Tests failed. Distribution was not created."
}

& $DotnetPath publish $projectPath `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=None `
    -p:DebugSymbols=false `
    -o $publishDirectory
if ($LASTEXITCODE -ne 0) {
    throw "Publishing the Windows x64 application failed."
}

New-Item -ItemType Directory -Force -Path $distributionRoot | Out-Null

foreach ($target in @($stageDirectory, $zipPath, $checksumPath)) {
    $resolvedTarget = [IO.Path]::GetFullPath($target)
    if (-not $resolvedTarget.StartsWith($distributionRoot.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove a path outside the distribution directory: $resolvedTarget"
    }

    if (Test-Path -LiteralPath $resolvedTarget) {
        Remove-Item -LiteralPath $resolvedTarget -Recurse -Force
    }
}

New-Item -ItemType Directory -Force -Path $stageDirectory | Out-Null
Copy-Item -LiteralPath (Join-Path $publishDirectory "AutomaticTestPrinting.App.exe") `
    -Destination $stageDirectory
foreach ($catalog in @("materials.json", "shared-material-catalog.json", "additional-excel-catalog.json", "additional-stage-catalog.json")) {
    $catalogPath = Join-Path $publishDirectory $catalog
    if (-not (Test-Path -LiteralPath $catalogPath)) { throw "Missing distribution catalog: $catalog" }
    Copy-Item -LiteralPath $catalogPath -Destination $stageDirectory
}
Copy-Item -LiteralPath $readmePath -Destination (Join-Path $stageDirectory "README-ja.txt")
$releaseNotesPath = Join-Path $repositoryRoot "packaging/RELEASE_NOTES-v$version.md"
if (Test-Path -LiteralPath $releaseNotesPath) {
    Copy-Item -LiteralPath $releaseNotesPath -Destination $stageDirectory
}

Compress-Archive -LiteralPath $stageDirectory -DestinationPath $zipPath -CompressionLevel Optimal

$hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
$checksumLine = "$hash  $packageName.zip" + [Environment]::NewLine
[IO.File]::WriteAllText($checksumPath, $checksumLine, [Text.UTF8Encoding]::new($false))

Write-Output $zipPath
Write-Output $checksumPath
