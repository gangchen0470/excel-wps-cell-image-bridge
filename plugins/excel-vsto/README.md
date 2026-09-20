# Excel VSTO 插件

本工程采用与“Excel 催化剂”相同的 VSTO 部署机制：Excel 注册项指向签名后的 `.vsto` 部署清单，由 Microsoft VSTO Runtime 负责加载，而不是手工注册普通 COM DLL。

开发测试：完全退出 Excel，双击 `安装VSTO插件.cmd`。脚本使用 Visual Studio 的 MSBuild 发布 ClickOnce 包，并调用 `VSTOInstaller.exe` 安装。

当前使用本地开发测试证书。正式分发前应替换为可信代码签名证书，并重新发布 ClickOnce 清单。

`生成Git发布包.cmd` 会生成以 GitHub Raw 为安装和更新地址的签名包。发布目录需完整上传到仓库根目录 `vsto/`，安装入口为：

`https://raw.githubusercontent.com/gangchen0470/excel-wps-cell-image-bridge/v1/wps-floating-foundation/vsto/CellImageBridgeVsto.vsto`

开发私钥 `.pfx` 不得上传。公开发布前应改用可信代码签名证书。
