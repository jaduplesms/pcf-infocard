[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot '..\artifacts'),
    [string]$ExpectedVersion
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ArchiveXml {
    param($Archive, [string]$Path)
    $entry = $Archive.GetEntry($Path)
    if ($null -eq $entry) { throw "Package is missing $Path." }
    $reader = [System.IO.StreamReader]::new($entry.Open())
    try { return [xml]$reader.ReadToEnd() }
    finally { $reader.Dispose() }
}

$root = Split-Path $PSScriptRoot -Parent
$project = Join-Path $root 'InfoCardSolution\InfoCardSolution.cdsproj'
$source = Join-Path $root 'InfoCardSolution\src'
$manifest = [xml](Get-Content (Join-Path $root 'InfoCardControl\InfoCard\ControlManifest.Input.xml') -Raw)
$version = [string]$manifest.manifest.control.version
$solution = [xml](Get-Content (Join-Path $source 'Other\Solution.xml') -Raw)
$package = Get-Content (Join-Path $root 'InfoCardControl\package.json') -Raw | ConvertFrom-Json
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw "Invalid control version: $version." }
if ($ExpectedVersion -and $ExpectedVersion -ne $version) { throw "Expected $ExpectedVersion, found $version." }
if ($package.version -ne $version -or $solution.ImportExportXml.SolutionManifest.Version -ne "$version.0") {
    throw 'Control, npm package and solution versions must match.'
}
if ($solution.ImportExportXml.SolutionManifest.UniqueName -ne 'InfoCardSample') {
    throw 'This release script only packages the InfoCardSample shipping solution.'
}

Push-Location (Join-Path $root 'InfoCardControl')
try {
    & npm run build -- --buildMode production
    if ($LASTEXITCODE -ne 0) { throw "Production control build failed ($LASTEXITCODE)." }
}
finally { Pop-Location }
$bundlePath = Join-Path $root 'InfoCardControl\out\controls\InfoCard\bundle.js'
$expectedBundleHash = (Get-FileHash -LiteralPath $bundlePath -Algorithm SHA256).Hash

$output = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$stage = Join-Path ([System.IO.Path]::GetTempPath()) ("infocard-release-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $stage | Out-Null
$artifacts = @()
try {
    foreach ($type in @('Unmanaged', 'Managed')) {
        $typeRoot = Join-Path $stage $type
        $stagedSource = Join-Path $typeRoot 'src'
        New-Item -ItemType Directory -Path $typeRoot | Out-Null
        Copy-Item -LiteralPath $source -Destination $stagedSource -Recurse
        $solutionPath = Join-Path $stagedSource 'Other\Solution.xml'
        $stagedSolution = [xml](Get-Content $solutionPath -Raw)
        $managed = if ($type -eq 'Managed') { '1' } else { '0' }
        $stagedSolution.ImportExportXml.SolutionManifest.Managed = $managed
        $stagedSolution.Save($solutionPath)
        $stagedProjectPath = Join-Path $typeRoot 'InfoCardSolution.cdsproj'
        $stagedProject = [xml](Get-Content $project -Raw)
        $projectReference = $stagedProject.SelectSingleNode("//*[local-name()='ProjectReference']")
        $projectReference.SetAttribute('Include', (Join-Path $root 'InfoCardControl\InfoCardControl.pcfproj'))
        $stagedProject.Save($stagedProjectPath)
        $zipPath = Join-Path $typeRoot 'bin\Release\InfoCardSolution.zip'

        & dotnet build $stagedProjectPath --configuration Release "/p:SolutionPackageType=$type"
        if ($LASTEXITCODE -ne 0) { throw "$type solution build failed ($LASTEXITCODE)." }
        if (-not (Test-Path -LiteralPath $zipPath)) { throw "Build did not produce $zipPath." }

        $archive = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
        try {
            $packedSolution = Read-ArchiveXml $archive 'solution.xml'
            $packedManifest = Read-ArchiveXml $archive 'Controls/smp_Sample.InfoCard/ControlManifest.xml'
            $customizations = Read-ArchiveXml $archive 'customizations.xml'
            $sm = $packedSolution.ImportExportXml.SolutionManifest
            if ($sm.UniqueName -ne 'InfoCardSample' -or $sm.Version -ne "$version.0" -or $sm.Managed -ne $managed) {
                throw "$type package solution identity, version or managed flag is incorrect."
            }
            if ($packedManifest.manifest.control.version -ne $version) {
                throw "$type package control version is incorrect."
            }
            $components = @($packedSolution.SelectNodes('//RootComponent'))
            if ($components.Count -ne 1 -or $components[0].type -ne '66' -or
                $components[0].schemaName -ne 'smp_Sample.InfoCard') {
                throw "$type package must contain only the InfoCard root component."
            }
            if ($customizations.SelectNodes('//Entity | //FormXml').Count -gt 0) {
                throw "$type package contains entity or form customizations."
            }
            $registrations = @($customizations.SelectNodes('/ImportExportXml/CustomControls/CustomControl'))
            if ($registrations.Count -ne 1 -or $registrations[0].Name -ne 'smp_Sample.InfoCard' -or
                $registrations[0].FileName -ne '/Controls/smp_Sample.InfoCard/ControlManifest.xml') {
                throw "$type package must register the InfoCard control manifest for clean installation."
            }
            $bundle = $archive.GetEntry('Controls/smp_Sample.InfoCard/bundle.js')
            if ($null -eq $bundle -or $bundle.Length -eq 0) { throw "$type package has no control bundle." }
            $stream = $bundle.Open()
            $sha = [System.Security.Cryptography.SHA256]::Create()
            try { $actualBundleHash = [System.BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
            finally { $sha.Dispose(); $stream.Dispose() }
            if ($actualBundleHash -ne $expectedBundleHash) { throw "$type package contains a stale control bundle." }
            Write-Host "$type verified: v$version; bundle $($bundle.Length) bytes."
        }
        finally { $archive.Dispose() }
        $artifacts += [pscustomobject]@{
            Source = $zipPath
            Name = "InfoCardSolution_${version}_$($type.ToLowerInvariant()).zip"
        }
    }
    foreach ($artifact in $artifacts) {
        Copy-Item -LiteralPath $artifact.Source -Destination (Join-Path $output $artifact.Name) -Force
    }
    $checksums = foreach ($artifact in $artifacts) {
        $hash = Get-FileHash -LiteralPath (Join-Path $output $artifact.Name) -Algorithm SHA256
        "$($hash.Hash.ToLowerInvariant())  $($artifact.Name)"
    }
    $checksums | Set-Content -LiteralPath (Join-Path $output "SHA256SUMS_${version}.txt") -Encoding ascii
    Write-Host "Verified release packages: $output"
}
finally {
    Remove-Item -LiteralPath $stage -Recurse -Force
}
