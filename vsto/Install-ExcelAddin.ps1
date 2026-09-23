param([switch]$Silent)
$ErrorActionPreference = "Stop"

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    $visibleExcel = @($excel | Where-Object { $_.MainWindowHandle -ne 0 })
    $backgroundExcel = @($excel | Where-Object { $_.MainWindowHandle -eq 0 })
    if ($visibleExcel.Count -gt 0) {
        Write-Host "Please close all Microsoft Excel windows, then run this installer again." -ForegroundColor Yellow
        if (-not $Silent) { Read-Host "Press Enter to exit" }
        exit 1
    }
    if ($backgroundExcel.Count -gt 0) {
        Write-Host "Closing leftover background Excel processes..."
        $backgroundExcel | Stop-Process -Force
        Start-Sleep -Milliseconds 500
    }
}

$installerCandidates = @(
    (Join-Path $env:CommonProgramFiles "Microsoft Shared\VSTO\10.0\VSTOInstaller.exe"),
    (Join-Path ${env:CommonProgramFiles(x86)} "Microsoft Shared\VSTO\10.0\VSTOInstaller.exe")
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $installerCandidates) {
    Write-Host "Microsoft Visual Studio Tools for Office Runtime is not installed." -ForegroundColor Red
    if (-not $Silent) { Read-Host "Press Enter to exit" }
    exit 2
}

$installer = $installerCandidates[0]
$manifest = Join-Path $PSScriptRoot "CellImageBridgeVsto.vsto"
if (-not (Test-Path -LiteralPath $manifest)) {
    throw "CellImageBridgeVsto.vsto was not found. Extract the complete ZIP before installation."
}

$registryPaths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\CellImageBridgeVsto",
    "HKCU:\Software\WOW6432Node\Microsoft\Office\Excel\Addins\CellImageBridgeVsto"
)

$manifestsToRemove = @($manifest)
foreach ($registryPath in $registryPaths) {
    if (-not (Test-Path $registryPath)) { continue }
    $oldManifest = (Get-ItemProperty -Path $registryPath -Name Manifest -ErrorAction SilentlyContinue).Manifest
    if (-not $oldManifest) { continue }
    $oldManifest = $oldManifest -replace "\|vstolocal$", ""
    $manifestsToRemove += $oldManifest
}

foreach ($downloadRoot in @((Join-Path $env:USERPROFILE "Downloads"), "D:\Users\Administrator\Downloads")) {
    if (-not (Test-Path -LiteralPath $downloadRoot)) { continue }
    $manifestsToRemove += Get-ChildItem -LiteralPath $downloadRoot -Recurse -Filter "CellImageBridgeVsto.vsto" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch "\\Application Files\\" } |
        Select-Object -ExpandProperty FullName
}

$uninstallRoot = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
if (Test-Path $uninstallRoot) {
    foreach ($entry in Get-ChildItem $uninstallRoot -ErrorAction SilentlyContinue) {
        $properties = Get-ItemProperty $entry.PSPath -ErrorAction SilentlyContinue
        if ($properties.DisplayName -ne "CellImageBridgeVsto" -or -not $properties.UrlUpdateInfo) { continue }
        $manifestsToRemove += $properties.UrlUpdateInfo
    }
}

function Invoke-VstoInstaller([string]$operation, [string]$sourceManifest, [bool]$silentMode) {
    $arguments = @($operation, ('"' + $sourceManifest + '"'))
    if ($silentMode) { $arguments += "/Silent" }
    $process = Start-Process -FilePath $installer -ArgumentList $arguments -Wait -PassThru
    return $process.ExitCode
}

Write-Host "Removing any previously installed CellImageBridgeVsto..."
foreach ($oldManifest in ($manifestsToRemove | Select-Object -Unique)) {
    # VSTO can retain an application in the ClickOnce cache without an Excel
    # Addins registry key. Uninstalling by the current manifest identity also
    # removes that copy, even when it was originally installed from another path.
    try {
        $oldUri = [Uri]$oldManifest
        if ($oldUri.IsFile -and -not (Test-Path -LiteralPath $oldUri.LocalPath) -and $oldUri.LocalPath -match "CellImageBridgeVsto-1\.0\.12") {
            $legacyArchive = Join-Path $PSScriptRoot "Legacy-CellImageBridgeVsto-1.0.12.zip"
            if (Test-Path -LiteralPath $legacyArchive) {
                $legacyFolder = Split-Path -Parent $oldUri.LocalPath
                New-Item -ItemType Directory -Path $legacyFolder -Force | Out-Null
                Expand-Archive -LiteralPath $legacyArchive -DestinationPath $legacyFolder -Force
            }
        }
    } catch { }
    [void](Invoke-VstoInstaller "/Uninstall" $oldManifest $true)
}

Write-Host "Opening the Microsoft Office add-in installer..."
$installExitCode = Invoke-VstoInstaller "/Install" $manifest ([bool]$Silent)

function Test-InstalledManifest {
    foreach ($registryPath in $registryPaths) {
        if (-not (Test-Path $registryPath)) { continue }
        $installed = (Get-ItemProperty -Path $registryPath -Name Manifest -ErrorAction SilentlyContinue).Manifest
        if (-not $installed) { continue }
        try {
            $installedPath = ([Uri]($installed -replace "\|vstolocal$", "")).LocalPath
            if ([string]::Equals($installedPath, $manifest, [stringcomparison]::OrdinalIgnoreCase)) { return $true }
        } catch { }
    }
    return $false
}

if ($Silent -and -not (Test-InstalledManifest)) {
    Write-Host "Silent installation needs Office confirmation. Opening the installer..."
    $installExitCode = Invoke-VstoInstaller "/Install" $manifest $false
}
if ($installExitCode -ne 0 -or -not (Test-InstalledManifest)) {
    Write-Host "VSTO installation did not register the new manifest." -ForegroundColor Red
    exit 1
}
exit 0
