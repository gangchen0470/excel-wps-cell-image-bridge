$ErrorActionPreference = "Stop"

$visibleExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
if ($visibleExcel.Count -gt 0) { throw "请先关闭所有可见的 Excel 窗口，然后重新运行修复程序。" }
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force

$searchRoots = @(
    (Join-Path $env:USERPROFILE "Downloads"),
    "D:\Users\Administrator\Downloads"
) | Where-Object { Test-Path -LiteralPath $_ }

$candidates = foreach ($root in $searchRoots) {
    Get-ChildItem -LiteralPath $root -Recurse -Filter "CellImageBridgeVsto.vsto" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -notmatch "\\Application Files\\" } |
        ForEach-Object {
            if ($_.Directory.Name -match "CellImageBridgeVsto-(\d+\.\d+\.\d+)") {
                [pscustomobject]@{ Version = [version]$Matches[1]; Path = $_.FullName }
            }
        }
}
$source = $candidates | Where-Object { $_.Version -le [version]"1.0.12" } | Sort-Object Version -Descending | Select-Object -First 1
if (-not $source) { throw "未找到以前下载并解压的插件。请重新下载发布包并运行 Install-ExcelAddin.cmd。" }

$registryPaths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\CellImageBridgeVsto",
    "HKCU:\Software\WOW6432Node\Microsoft\Office\Excel\Addins\CellImageBridgeVsto"
)
$repaired = $false
foreach ($registryPath in $registryPaths) {
    if (-not (Test-Path $registryPath)) { continue }
    $uri = ([Uri]$source.Path).AbsoluteUri + "|vstolocal"
    Set-ItemProperty -Path $registryPath -Name Manifest -Value $uri
    Set-ItemProperty -Path $registryPath -Name LoadBehavior -Value 3 -Type DWord
    $repaired = $true
}
if (-not $repaired) { throw "未找到插件注册信息，请运行发布包中的 Install-ExcelAddin.cmd。" }
Write-Host ("Registration repaired: " + $source.Path) -ForegroundColor Green
Write-Host "Open Excel, click Check Update, and install the latest version." -ForegroundColor Green
