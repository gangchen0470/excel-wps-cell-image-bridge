param([Parameter(Mandatory=$true)][int]$ExcelProcessId)
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

function Show-Result([string]$message, [bool]$success) {
    $icon = if ($success) { [System.Windows.Forms.MessageBoxIcon]::Information } else { [System.Windows.Forms.MessageBoxIcon]::Error }
    [void][System.Windows.Forms.MessageBox]::Show($message, "Excel / WPS 图片修复更新", [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
}

try {
    $deadline = (Get-Date).AddMinutes(15)
    while ((Get-Process -Id $ExcelProcessId -ErrorAction SilentlyContinue) -or
           @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }).Count -gt 0) {
        if ((Get-Date) -ge $deadline) { throw "等待 Excel 关闭超时。请手动运行安装包中的 Install-ExcelAddin.cmd。" }
        Start-Sleep -Seconds 2
    }

    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "Install-ExcelAddin.ps1") -Silent
    if ($LASTEXITCODE -ne 0) { throw "安装程序返回错误代码 $LASTEXITCODE。请手动运行安装包中的 Install-ExcelAddin.cmd。" }
    Show-Result "插件已更新。请重新打开 Excel。" $true
} catch {
    Show-Result ("自动更新失败：" + $_.Exception.Message + "`n安装包位置：" + $PSScriptRoot) $false
    exit 1
}
