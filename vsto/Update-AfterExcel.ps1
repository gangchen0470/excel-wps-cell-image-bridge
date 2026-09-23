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

    $installRoot = Join-Path $env:LOCALAPPDATA ("CellImageBridgeVsto\installed\" + $version)
    $staging = $installRoot + ".staging-" + [guid]::NewGuid().ToString("N")
    New-Item -ItemType Directory -Path $staging | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot "Application Files") -Destination $staging -Recurse
    Get-ChildItem -LiteralPath $PSScriptRoot -File | Copy-Item -Destination $staging
    if (Test-Path -LiteralPath $installRoot) { Remove-Item -LiteralPath $installRoot -Recurse -Force }
    Move-Item -LiteralPath $staging -Destination $installRoot

    $installedManifest = Join-Path $installRoot "CellImageBridgeVsto.vsto"
    $manifestUri = ([Uri]$installedManifest).AbsoluteUri + "|vstolocal"
    $registryPaths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins\CellImageBridgeVsto",
        "HKCU:\Software\WOW6432Node\Microsoft\Office\Excel\Addins\CellImageBridgeVsto"
    )
    $updated = $false
    foreach ($registryPath in $registryPaths) {
        if (-not (Test-Path $registryPath)) { continue }
        Set-ItemProperty -Path $registryPath -Name Manifest -Value $manifestUri
        Set-ItemProperty -Path $registryPath -Name LoadBehavior -Value 3 -Type DWord
        $updated = $true
    }
    if (-not $updated) { throw "未找到插件注册信息，请运行安装包中的 Install-ExcelAddin.cmd。" }
    Set-Content -LiteralPath (Join-Path $installRoot "update-ready.txt") -Value ("version=" + $version + "`r`nprepared=" + (Get-Date).ToString("O"))
    Show-Result ("新版 " + $version + " 已准备完成。当前 Excel 可以继续使用；下次启动时自动加载新版。") $true
} catch {
    Show-Result ("自动更新失败：" + $_.Exception.Message + "`n安装包位置：" + $PSScriptRoot) $false
    exit 1
}
