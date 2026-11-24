# PDF 转 Word

Windows 10/11 x64 本地转换器。双击 `PDFtoWord.exe`，选择或拖入 PDF，选好转换方式后点击「导出 Word」。文件在本机处理，输出为新的 `.docx`。

| 转换方式 | 适用场景 | Word 中的结果 |
| --- | --- | --- |
| 保留原版排版 | 打印、归档、复杂表格、公式、图章 | 以 PDF 实际显示尺寸逐页嵌入 PNG。每页图像可整体选中，文字、单元格和公式随图保存。 |
| 可编辑文字和表格 | 能选中文字的报告、合同、简单表格 | 本机桌面版 Microsoft Word 2013 或更新版本推断段落和表格。分栏、跨页表格、脚注、缺失字体可能引起重排。 |
| 扫描文字识别 | 需要从扫描件提取文字 | 调用 Windows OCR，逐行生成可编辑段落。图形不参与文字重建，识别不到文字的页面以图片保留。 |

## 清晰度与兼容性

保版默认 300 DPI，图纸或需要放大查阅的小字可选择 450 或 600 DPI。每页最多生成 3000 万像素，达到上限会列出受影响页码。原始扫描件本身的清晰度决定可见细节。保版保留打印外观，PDF 链接、可填写表单、批注交互和数字签名不会转换为 Word 的对应功能。可见表单内容的还原取决于 PDF 的页面渲染结果。

Word 的页面最长边为 55.88 厘米。遇到超大页时，保版会提示对应页码并终止，避免静默缩小页面。输入密码仅用于保版和 OCR 的本地 PDF 打开操作。Word 重排模式需要无需打开密码的副本。

扫描识别优先使用已安装的简体中文 OCR，也可回退到系统可用的识别语言。若缺少所需语言，可在 Windows 语言设置中添加对应 OCR 组件。签名、金额、小数点、编号、表格的阅读顺序请对照原件核查。

## 命令行

以下示例供 PowerShell 使用。图形程序应通过管道运行，以便脚本等待转换结束。`$LASTEXITCODE` 为 0 表示成功。输入输出路径支持空格和中文。已有同名文件时会报错。

```powershell
.\PDFtoWord.exe fixed "巡检记录.pdf" "巡检记录-保版.docx" 300 | Out-Host
.\PDFtoWord.exe edit "巡检记录.pdf" "巡检记录-可编辑.docx" | Out-Host
.\PDFtoWord.exe ocr "扫描签收单.pdf" "扫描签收单-文字.docx" | Out-Host
```

## 从源码构建

在 Windows x64 上使用 .NET Framework 4.8 和 Windows 10/11 SDK。于 64 位 PowerShell 中运行 `.\build.ps1`，输出为 `bin\PDFtoWord.exe`。亦可用 `-WindowsMetadata` 指定 SDK 的 `UnionMetadata\<版本>\Windows.winmd`。运行成品不需要 SDK、Python 或额外的下载。

## 样例与验证

`samples` 中是本项目生成的测试 PDF。巡检记录含中英文表格、旋转裁切横页及扫描签收单，学术短文含双栏、上下标和脚注，表单含文字框和已勾选的复选框。运行 `.\tests\smoke.ps1` 可验证三页尺寸、图像数量和同名文件保护。对排版要求较高时，可用桌面 Word 打开保版结果，再导出 PDF，与原稿按相同分辨率逐页比较。

开源项目 [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) 是另一种基于规则重建文档的方法，Artifex 已停止主动维护。本工具的保版模式使用 Windows PDF 渲染，编辑模式依赖本机 Word 的 PDF 重排。[微软对 PDF 转换能力的说明](https://support.microsoft.com/en-us/word/opening-pdfs-in-word)
