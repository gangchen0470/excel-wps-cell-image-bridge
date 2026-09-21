# Excel / WPS 单元格图片兼容工具

## Excel 插件下载与安装

Windows 和 Microsoft Excel 用户请下载：

**[CellImageBridgeVsto-1.0.6.zip](https://github.com/gangchen0470/excel-wps-cell-image-bridge/releases/download/v1.0.6/CellImageBridgeVsto-1.0.6.zip)**

1. 将 ZIP **完整解压**到本地文件夹，不要直接在压缩包内运行。
2. 关闭所有 Microsoft Excel 窗口。
3. 双击 `Install-ExcelAddin.cmd`；它会先卸载从其他目录安装的旧版，再打开新版安装程序。首次安装也可直接双击 `CellImageBridgeVsto.vsto`。
4. 在“Microsoft Office 自定义项安装程序”中点击“安装”。
5. 重新打开 Excel，功能区中应出现 **图片修复** 选项卡。

请不要下载绿色 **Code → Download ZIP**，那是项目源码，不是 Excel 插件安装包。

运行要求：Windows、Microsoft Excel、.NET Framework 4.8 和 Microsoft Visual Studio Tools for Office Runtime。当前发布包使用项目开发证书签名，Windows 可能显示证书或发布者确认提示。

### 在 Excel 中使用

1. 先用 Excel 打开从 WPS 保存的本地 `.xlsx` 原文件；转换前不要用 Excel 覆盖保存原文件。
2. 打开 **图片修复** 选项卡，可先点击 **检测 WPS 图片**。
3. 点击 **一键转换**，在当前工作簿中检查图片效果。
4. 点击 **另存为兼容版**，使用新文件名保存 `.xlsx` 副本。

插件同时识别 WPS `DISPIMG` 和新版 Excel“置于单元格”图片。**修复并另存兼容版**会先将两类图片转换为普通浮动图片，再保存新副本；浮动图片可在 WPS 和不支持单元格图片的旧版 Excel 中显示。原文件不会被覆盖。

点击功能区中的 **检查更新**，插件会读取仓库根目录的 `update.json`。发现新版本时可直接打开对应安装包下载地址。

如果安装后没有出现选项卡，请在 Excel 中打开“文件 → 选项 → 加载项 → COM 加载项”，确认 `CellImageBridgeVsto` 已启用。

如果直接双击 `.vsto` 时提示“已安装的自定义项不能从该位置升级”，说明下载的仍是 `1.0.1` 旧包。请下载 `1.0.4` 或更高版本；新版部署清单已提高版本号，可从新目录升级。

ZIP 内含完整的 VSTO 安装文件；仓库的 [`vsto/`](https://github.com/gangchen0470/excel-wps-cell-image-bridge/tree/main/vsto) 目录保留同一份内容。`Application Files` 目录必须和 `CellImageBridgeVsto.vsto` 一起保留。

V1 主线：**WPS DISPIMG / Excel Place in Cell → Excel 和 WPS 可见的标准浮动图片**。项目包含 Python 3.10+ OOXML 转换器与浏览器/NAS 页面。

## 恢复规则

- 保持原图比例、完整显示、水平和垂直居中。
- 默认 2px 内边距（96 DPI），可用 `--margin` 调整。
- 绑定原单元格；合并单元格以整个合并区域计算。
- 写入 `twoCellAnchor editAs="twoCell"`，声明随单元格移动和缩放，并锁定宽高比。
- 保留行列尺寸；沿用现有行高锁定逻辑，避免清除公式后的自动重算。
- 另存副本，拒绝输入输出同路径。指定的已有输出文件会被替换。

转换时的比例、居中、边距已做结构验证。静态锚点不能承诺任意行列调整后自动重新等比居中或保持固定边距；实际表现须在 Excel 验证，必要时后续插件重新布局。

## 运行（仓库根目录，Windows PowerShell）

```powershell
$env:PYTHONPATH = "src"
python -m cell_image_compat.cli inspect "input.xlsx"
python -m cell_image_compat.cli compatible "input.xlsx" "compatible.xlsx" --log "conversion.json"
python -m unittest discover -s tests -v
```

可选 `--margin 2 --sheet "Sheet1" --range "B2:B500"`。macOS/Linux 使用 `PYTHONPATH=src python3 -m cell_image_compat.cli ...`。也可 `python -m pip install -e .` 安装命令行入口。

## 目录

```text
src/cell_image_compat/
  xlsx_parser.py     # ZIP/XML、关系路径、工作表和内容类型
  wps_dispimg.py     # DISPIMG → 图片 ID → 媒体文件
  layout.py         # 默认规则、合并区域、锚点坐标归一化
  core.py           # 图片尺寸、DrawingML 转换和包写回；保留旧能力
  model.py          # 图片模型和异常
  cli.py            # 检测及转换入口
plugins/
  excel/            # 已有 Excel 加载项清单
  excel-vsto/       # Microsoft Excel VSTO 插件源码和发布脚本
  wps/              # 已有 WPS 外壳
  shared/           # 已有共享任务窗格
  service.py        # 本地开发服务
tests/
  fixtures/         # 结构夹具生成器、真实样本约定
  output/           # 人工测试输出目录
  test_v1.py        # 无外部样本依赖的回归测试
  test_samples.py   # 既有单元测试及可选真实样本测试
docs/DEVELOPMENT.md  # 开发路线和客户端验收
vsto/                # 已签名的 VSTO 安装目录
```

沿用已有包结构，不建立重复的空目录。仓库已有原生双向互转实验能力，本轮暂不扩展，不作为 V1 验收目标。

## 验证与边界

当前 Python 测试共 20 项：14 项通过，6 项因缺少真实样本跳过；另有 2 项页面交互测试通过。覆盖合并区域、横竖图、重复引用、已有 Drawing 追加、默认边距、比例、源文件和媒体保留、缺失媒体、异常公式、JPEG 宽高解析。

只接受独立 `DISPIMG("id",1)`，支持 `_xlfn.` 前缀、大小写和空格。嵌套表达式、动态 ID 和缺失图片会报错终止。部分转换保留未转换图片的索引。隐藏行列、复杂公式和特殊字体尺寸仍待完善。

自动化实机测试已用 Excel 完成 3 张合成图片的检测、转换、位置属性检查以及保存重开；下一步使用含 5–10 张图片的真实 WPS 文件做人工视觉验收。详见 [开发说明](docs/DEVELOPMENT.md)。

## 本地页面与插件任务窗格

```powershell
$env:PYTHONPATH = "src"
python plugins/service.py --port 3000
```

打开 http://127.0.0.1:3000 ，选择已保存的 `.xlsx`：

1. **检测 WPS 图片**：列出数量及单元格，不生成文件。
2. **一键修复**：生成兼容副本，完成后启用保存按钮。
3. **另存为兼容版**：发起浏览器下载；保存位置由浏览器设置控制。

可以直接修复，不强制先检测。更换文件或处理失败会清除上次结果，防止误保存。页面最多接受 25 MB 文件，仅处理 WPS 图片；既有 Excel 原生图片保留。原生互转仍可通过旧命令行入口调用，页面不再提供。

同一页面供浏览器和旧任务窗格外壳加载，统一使用文件选择器。需要直接操作活动工作簿时，请使用上面的 Windows 原生插件。

页面测试（开发时需要 Node.js）：`node --test tests/test_ui.cjs`。普通运行只需要 Python。
