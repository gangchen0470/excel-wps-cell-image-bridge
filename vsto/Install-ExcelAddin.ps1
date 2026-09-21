$ErrorActionPreference = "Stop"

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
    $visibleExcel = @($excel | Where-Object { $_.MainWindowHandle -ne 0 })
    $backgroundExcel = @($excel | Where-Object { $_.MainWindowHandle -eq 0 })
    if ($visibleExcel.Count -gt 0) {
        Write-Host "Please close all Microsoft Excel windows, then run this installer again." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
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
    Read-Host "Press Enter to exit"
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

Write-Host "Removing any previously installed CellImageBridgeVsto..."
foreach ($oldManifest in ($manifestsToRemove | Select-Object -Unique)) {
    # VSTO can retain an application in the ClickOnce cache without an Excel
    # Addins registry key. Uninstalling by the current manifest identity also
    # removes that copy, even when it was originally installed from another path.
    & $installer /Uninstall $oldManifest /Silent 2>$null
}

Write-Host "Opening the Microsoft Office add-in installer..."
& $installer /Install $manifest
exit $LASTEXITCODE
