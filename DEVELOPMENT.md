# PDFToWord 开发文档

更新日期：2026-10-23。改动历史见 `CHANGELOG.md`。

## 架构和转换路线

`src/main.cpp` 提供 Win32 界面、命令行、内置资源释放和 Windows OCR。自动模式固定使用内置可编辑引擎；Word 模式需用户明确选择。Word COM 注册不代表 PDF 导入可靠。

`editable_converter.py` 组织排版、字体数字修复、扫描识别、图片表格替换、DOCX 关系重绑定和报告。

| 输入 | 主要模块及行为 |
| --- | --- |
| 原生文字 PDF | pdf2docx 排版、字体修复、数字完整性检查 |
| 规则扫描表格 | `scan_tables.py` 检测线网和合并格，Windows OCR 后由 `cell_ocr.py` 复核，生成真实 Word 单元格 |
| 文字页内的图片表格 | 按图片边界使用同一识别流程，替换原图片 |
| 折皱拍照表格 | `photo_tables.py` 矫正弯曲网格，`photo_ocr.py` 使用内置 PP-OCR |
| 无法可靠重建的图像 | 保留图片，视 OCR 可用性提供文字层和警告 |

## 单元格复核规则

`cell_ocr.py` 不依赖客户文件名、固定行号或金额。

1. 从表头识别序号、数量、单价、金额、合计、总价列；至少两个数字列才开启复核。跳过表头、合并区域、合计文字标签。
2. 从去除线条和印章干扰的图像裁出实际单元格墨迹，生成原裁片及 1.5 倍缩放加白边裁片。空白裁片跳过。
3. 复用 `photo_ocr.recognizer()` 的离线模型，最多 64 个裁片一批；两次读数相同、置信度均不低于 0.98 才考虑替换。
4. 数字必须满足严格十进制语法，支持多位小数，不猜测逗号是千分位还是小数点。
5. 含数字的规格修正要求中文字符序列不变；空白可选图号不生成内容。空格、数字之间的 `x/X/*/×` 差异仅用于等价比较。
6. 数量乘单价与金额相差超过 0.01 时提示复核，绝不用算式补写数值，保留原件可能存在的折扣或错误。

原 Windows OCR 空白数字格重试仍保留，新复核在其后运行，覆盖“非空但错误”的文本。

## 印章和页脚

`scan_tables.py` 整页分组红色印章，圆环、星形共享一个透明 PNG；表格与页脚使用同一组对象，避免重复导出。印章删除后，公司名和日期仍是可编辑文字。

`verify_stamp_prose` 复核表格外、与印章相交的已有文字行。要求两次读数一致、置信度至少 0.98、与原读数相似度至少 0.75、数字片段完全一致。不凭常见公司名补字，也不重绘印章。日期保留原日期专用放大重试。

## 报告字段

| 字段 | 含义 |
| --- | --- |
| `cell_ocr_checked` | 参与复核的单元格数，不代表正确率 |
| `cell_ocr_changes` | 修改前后文本和页、表、行、列；页脚另带 `region` |
| `cell_ocr_issues` | 识别分歧或数量单价金额关系异常 |
| `cell_ocr_errors` | 复核失败原因，保留第一轮结果 |
| `warnings` | 界面可读提示，包含需要人工复核的位置 |

扫描页和内嵌图片共用 `collect_cell_ocr_report`，合并时累加而非覆盖。扫描 PDF 的 `numbers_checked` 可能为零，因为没有原生文字层，不能据此声称 OCR 数字无误。

## 构建和测试

编译路径见 `build.ps1`，依赖约束见 `requirements-build.txt`。本工作区补充包在 `.build-deps`；OCR 模型随 EXE 打包，转换时不下载模型、不上传文件。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Clean
$env:PYTHONPATH = "$PWD\.build-deps;$PWD\editable_engine"
$env:PYTHONIOENCODING = 'utf-8'
$taskPython = 'C:\Users\ZZ\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
& $taskPython tests\test_numeric_check.py
& $taskPython tests\test_scan_tables.py
& $taskPython tests\test_cell_ocr.py
& $taskPython tests\verify_final_exe.py
```

干净构建重建 `build`、`dist`；先关闭测试启动的旧 EXE，避免文件被占用。最终程序为 `dist/PDFToWord-Single.exe`。

`verify_final_exe.py` 只复制最终 EXE 到 `tmp/final-exe-run`，清除子进程 PYTHONPATH，要求新生成输出再审计；通过后才复制到 `output`。`output/final-exe-audit.json` 记录运行状态、UTC 时间、程序/源文件/DOCX 的 SHA256、耗时、表数和节数。失败会更新状态，不沿用旧的通过报告。

| 测试 | 覆盖内容 |
| --- | --- |
| `test_cell_ocr.py` | 错误金额、缺失价格、多位小数、分歧拒绝、规格约束、空白保留、算式只提示 |
| `audit_xindi.py` | 三页 108 条，432 个数字格、总额、大写金额、规格及空白图号 |
| `audit_quotation.py` | 原两页 50 条的图号、数量、单价、金额及合并格 |
| `audit_stamps.py` | 印章数量、透明背景、红色像素守恒、无重复、公司和日期可编辑、绘图 ID 唯一 |
| `audit_photo_one.py` | 拍照样本 41 条、5 列、图号、数量、金额和独立灰色印章 |
| 混合图片与原生字体夹具 | 图片表格替换和私有区数字回归 |

预期数据从原扫描页核对记录，不能直接用待测 OCR 输出更新预期。

## 渲染验收和已知限制

最终 EXE 生成的 DOCX 必须重新渲染，逐页检查文字、边界、分页和印章。本机 QA 使用 `tmp/libreoffice-runtime/program` 的已验证便携渲染器及文档技能 `render_docx.py`，临时 PNG/PDF 仅供验收。

GUI 启动、控件检查、截图是独立项目，构建及 DOCX 渲染通过不能替代 GUI 截图。截图工具超时必须如实记入日志。

两次模型读数仍可能一致误读。当前规则依赖表格线及中文表头；无线表格、严重模糊、遮挡、未覆盖版式仍需人工核对。普通中文品名不会无条件替换。拍照路径仍先调用 Windows OCR，不能声称已验证无 Windows OCR 语言包的环境。
