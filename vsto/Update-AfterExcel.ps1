param([int]$ExcelProcessId = 0)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

function Show-Result([string]$message, [bool]$success) {
    $icon = if ($success) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Error }
    [void][System.Windows.Forms.MessageBox]::Show($message, "Excel / WPS 图片修复更新", [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
}

try {
    $manifest = Join-Path $PSScriptRoot "CellImageBridgeVsto.vsto"
    if (-not (Test-Path -LiteralPath $manifest)) { throw "安装包缺少 CellImageBridgeVsto.vsto。" }
    [xml]$deployment = Get-Content -LiteralPath $manifest
    $identity = $deployment.assembly.assemblyIdentity
    if (-not $identity.version) { throw "无法读取新版版本号。" }
    $version = ([version]$identity.version).ToString(3)

    $registryPaths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins\CellImageBridgeVsto",
        "HKCU:\Software\WOW6432Node\Microsoft\Office\Excel\Addins\CellImageBridgeVsto"
    )
    $installedManifest = $null
    foreach ($registryPath in $registryPaths) {
        if (-not (Test-Path $registryPath)) { continue }
        $value = (Get-ItemProperty -Path $registryPath -Name Manifest -ErrorAction SilentlyContinue).Manifest
        if (-not $value) { continue }
        try {
            $uri = [Uri]($value -replace "\|vstolocal$", "")
            if ($uri.IsFile) { $installedManifest = $uri.LocalPath; break }
        } catch { }
    }
    if (-not $installedManifest) { throw "未找到本地插件来源，请运行安装包中的 Install-ExcelAddin.cmd。" }
    $installRoot = Split-Path -Parent $installedManifest
    if (-not (Test-Path -LiteralPath $installRoot)) { throw "原插件安装目录不存在，请重新安装插件。" }

    Copy-Item -LiteralPath (Join-Path $PSScriptRoot "Application Files") -Destination $installRoot -Recurse -Force
    Get-ChildItem -LiteralPath $PSScriptRoot -File | Where-Object { $_.Name -ne "CellImageBridgeVsto.vsto" } | Copy-Item -Destination $installRoot -Force
    Copy-Item -LiteralPath $manifest -Destination $installedManifest -Force
    Set-Content -LiteralPath (Join-Path $installRoot "update-ready.txt") -Value ("version=" + $version + "`r`nprepared=" + (Get-Date).ToString("O"))
    Show-Result ("新版 " + $version + " 已准备完成。当前 Excel 可以继续使用；下次启动时自动加载新版。") $true
} catch {
    Show-Result ("自动更新失败：" + $_.Exception.Message + "`n安装包位置：" + $PSScriptRoot) $false
    exit 1
}
