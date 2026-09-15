# WPS & Excel 单元格图片兼容核心

当前版本提供一个零第三方运行时依赖的 OOXML 核心，可识别：

- Microsoft Excel Rich Data 单元格图片
- WPS `DISPIMG` / `cellimages.xml` 单元格图片

并可将两者转换为标准 DrawingML 浮动图片，原媒体文件直接复用、不重新压缩。输出图片按原单元格等比例缩放、居中，默认使用完整单元格可用空间；可通过 `--margin` 增加边距。

多图片场景中，每张图片均按自己的原始宽高比和所在单元格独立计算：横图受宽度限制、竖图受高度限制、正方形图片受较短边限制。转换不会为了某张图片修改整列宽度或所在行高度。

尺寸解析覆盖 PNG、JPEG（基线/渐进及常见 SOF 类型）、GIF、BMP 与 WebP（VP8/VP8L/VP8X）。

## 使用

```bash
PYTHONPATH=src python3 -m cell_image_compat.cli inspect "input.xlsx"
PYTHONPATH=src python3 -m cell_image_compat.cli compatible "input.xlsx" "output.xlsx"
PYTHONPATH=src python3 -m cell_image_compat.cli compatible "input.xlsx" "output.xlsx" --margin 5
PYTHONPATH=src python3 -m cell_image_compat.cli compatible "input.xlsx" "output.xlsx" --sheet "Sheet1" --range "B2:B500" --log "conversion.json"
PYTHONPATH=src python3 -m cell_image_compat.cli native "wps-input.xlsx" "excel-output.xlsx" --target excel
PYTHONPATH=src python3 -m cell_image_compat.cli native "excel-input.xlsx" "wps-output.xlsx" --target wps
PYTHONPATH=src python3 -m cell_image_compat.cli floating-native "input.xlsx" "output.xlsx" --target wps --sheet "Sheet1" --range "B2:B500"
```

工具拒绝用相同路径覆盖源文件。

## 插件开发外壳

启动共享本地服务：

```bash
PYTHONPATH=src python3 plugins/service.py --port 3000
```

- Excel 开发清单：`plugins/excel/manifest.xml`
- WPS 功能区：`plugins/wps/ribbon.xml` 与 `plugins/wps/main.js`
- 两端共用：`plugins/shared/` 任务窗格

当前为本地开发外壳，输出以新文件下载，不覆盖活动工作簿。正式发布前需要配置 HTTPS、签名/发布流程和服务鉴权。

## 正式发布目标

正式版以 Windows 为首要平台，最终提供普通用户可直接运行的安装程序：

- 安装程序内置转换核心和运行环境，用户无需安装 Python、Node.js 或命令行工具。
- 安装后分别在 Microsoft Excel 与 WPS 表格中显示“图片兼容工具”。
- 后台转换组件随用户登录自动启动，仅监听本机回环地址。
- Excel 使用正式 HTTPS 任务窗格资源及受信任加载项目录；WPS 使用 `wpsjs publish` 生成的发布包。
- 安装程序负责注册、升级和卸载，不覆盖用户工作簿。

仓库中的 `plugins/` 当前是开发外壳；Windows 可安装包、代码签名和两端发布包仍属于后续发布阶段。

通用兼容版默认处理整个工作簿，也可按工作表及 A1 区域筛选。JSON 日志记录扫描、成功、失败、跳过数量和完成时间。部分范围转换会保留源格式索引，未选中的单元格图片继续维持原格式。

交互规则：嵌入单元格图片转浮动图片默认整本处理，不要求选择；浮动图片转原生单元格图片必须同时提供工作表、选择区域和目标格式。浮动图片按中心点归属单元格，只有中心点位于选区中的图片会被转换。

## 当前边界

- 已实现识别和“通用兼容版”输出。
- 已支持在无既有 Drawing 的样本中生成标准浮动图片；存在 Drawing 时会追加。使用单单元格锚点并显式写入图片变换尺寸，以避免不同客户端重算双锚点偏移。
- 已实现 Excel Rich Data 与 WPS CellImage 的原生格式互转写回，当前通过结构闭环测试，仍需 Microsoft Excel 与 WPS 目标客户端实机确认。
- 合并单元格、隐藏行列、多图冲突和损坏关系的逐项容错将在下一阶段实现。
