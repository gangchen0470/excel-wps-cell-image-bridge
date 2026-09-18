# V1 开发说明

## 现状与模块边界

上游已有 Python OOXML 核心、命令行、Excel/WPS 插件外壳和实验性原生互转。本轮沿用结构，将 ZIP/XML/关系解析拆到 `xlsx_parser.py`，WPS 映射拆到 `wps_dispimg.py`，合并区域与锚点坐标拆到 `layout.py`；转换编排仍在 `core.py`，公共 API 不变。

本地分支 `v1/wps-floating-foundation`。浏览器/NAS 页面已统一检测和转换 WPS DISPIMG 与 Excel Place in Cell；原生 Excel 插件试验未达到可发布标准，不纳入本分支。

## 转换流程

1. 读取工作簿与关系，支持包内绝对路径；忽略外部关系。
2. 解析独立 DISPIMG，按 cellimages 中的名称和图片关系定位媒体；无法完整映射时明确报错。
3. 读取原图尺寸，定位普通或合并区域，计算完整显示的等比尺寸及居中偏移。
4. 追加或创建 Drawing，复用媒体，写双单元格锚点、宽高比锁和显式图形尺寸。图片 ID 从已有最大值递增。
5. 清除已恢复单元格的公式和缓存值；部分转换保留原生索引，全量完成再清理。
6. 校验 ZIP 后写出副本，拒绝源文件路径。

## 锚点决策及限制

`twoCellAnchor editAs="twoCell"` 对应随单元格移动和缩放，依据 [Microsoft Open XML 文档](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.editasvalues?view=openxml-3.0.1)。

历史交接曾因客户端重新解释双锚点改用 oneCellAnchor；本轮按 V1 要求改回，必须实机验证该回归点。标准锚点和宽高比锁不代表客户端在任意单边尺寸变化后都自动重新居中、等比适配。必要时后续插件监听尺寸变化重新布局。

列宽沿用已有 7px 字符宽近似，未按工作簿字体测量；不同字体和 DPI 可能有偏差。小于两倍边距的区域优先保证正尺寸，无法保留完整 2px 留白。隐藏行列、共享/数组公式、多图冲突、严格 OOXML 命名空间、任意扩展命名空间无损回写、宏和加密工作簿不在本轮保证范围。图片头尺寸解析不等于图像解码验证，也不保证目标客户端支持所有图片格式。

## 自动测试

```powershell
$env:PYTHONPATH = "src"
python -m unittest discover -s tests -v
```

结构夹具由 `tests/fixtures/build_wps.py` 在临时目录生成，不依赖用户文件。可选真实样本：

```powershell
$env:WPS_SAMPLE = "C:\samples\wps-single-A1.xlsx"
$env:EXCEL_SAMPLE = "C:\samples\excel-single-A1.xlsx"
python -m unittest discover -s tests -v
```

旧样本测试约定 Sheet1!A1 单张图片；多图需单独增加预期。目前 Python 测试 14 项通过、6 项缺样本跳过，另有 2 项 Node 页面交互测试通过；结构测试不替代客户端视觉验收。

## 下一里程碑：真实 WPS → Excel

准备 5–10 张置于单元格图片的 WPS `.xlsx`：横竖图、普通/合并单元格、不同列宽行高、重复图片、已有浮动图。保留原文件，记录 WPS 和 Excel 版本。

- 核对每个图片 ID、媒体、单元格及总数。
- Excel 打开无修复提示；图片完整、等比、居中、默认 2px 留白。
- 原浮动图、公式、文字、行列尺寸与其他工作表正常。
- 插入行列，单独及同时调整行高列宽，记录移动、缩放、比例和居中表现。
- 保存关闭再打开，检查数量和显示。
- 分范围转换后补转剩余图片，确认无重复或丢失。

三个按钮已接通：检测 WPS 图片（只读）、一键修复（生成副本）、另存为兼容版（发起浏览器下载）。页面采用文件选择器，浏览器和任务窗格共用，不依赖活动工作簿导出 API。服务仅处理上传字节，已移除按客户端提供路径读取本地文件的旧服务入口。服务绑定 127.0.0.1，校验 Host/Origin；不提供外部部署。活动工作簿导出、HTTPS 和插件安装验收尚未完成。
