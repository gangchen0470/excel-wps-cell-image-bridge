# Excel VSTO 插件

本工程采用与“Excel 催化剂”相同的 VSTO 部署机制：Excel 注册项指向签名后的 `.vsto` 部署清单，由 Microsoft VSTO Runtime 负责加载，而不是手工注册普通 COM DLL。

开发测试：完全退出 Excel，双击 `安装VSTO插件.cmd`。脚本使用 Visual Studio 的 MSBuild 发布 ClickOnce 包，并调用 `VSTOInstaller.exe` 安装。

当前使用本地开发测试证书。正式分发前应替换为可信代码签名证书，并重新发布 ClickOnce 清单。

`生成Git发布包.cmd` 会生成签名后的完整安装目录。将目录完整复制到仓库根目录 `vsto/`，再压缩并作为 GitHub Release 附件发布。普通用户应下载 Release ZIP、完整解压并从本地双击 `CellImageBridgeVsto.vsto`；不要只下载单个部署清单，因为安装还需要相邻的 `Application Files` 目录。

当前版本下载地址：

`https://github.com/gangchen0470/excel-wps-cell-image-bridge/releases/download/v1.0.1/CellImageBridgeVsto-1.0.1.zip`

发布 ZIP 还包含 `Install-ExcelAddin.cmd` 和 `Install-ExcelAddin.ps1`。安装器会读取 Excel COM 加载注册项，卸载从其他本地路径安装的旧版，解决 VSTO 的“不能从该位置升级”错误。

开发私钥 `.pfx` 不得上传。公开发布前应改用可信代码签名证书。
