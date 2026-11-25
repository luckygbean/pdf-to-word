# PDF 转 Word

Windows 10/11 x64 本地转换器。双击 `PDFtoWord.exe`，选择或拖入 PDF，选好转换方式后点击「导出 Word」。文件在本机处理，输出为新的 `.docx`。

| 转换方式 | 适用场景 | Word 中的结果 |
| --- | --- | --- |
| 保留原版排版 | 打印、归档、复杂表格、公式、图章 | 以 PDF 实际显示尺寸逐页嵌入 PNG。每页图像可整体选中，文字、单元格和公式随图保存。 |
| 可编辑文字和表格 | 招标文件、电子表单，以及能选中文字的报告或合同 | 含可识别线框表格时，按原稿分页重建正文和单元格。其他文档由本机桌面版 Microsoft Word 2013 或更新版本重排。 |
| 扫描文字识别 | 需要从扫描件提取文字 | 调用 Windows OCR，逐行生成可编辑段落。图形不参与文字重建，识别不到文字的页面以图片保留。 |

## 清晰度与兼容性

保版默认 300 DPI，图纸或需要放大查阅的小字可选择 450 或 600 DPI。每页最多生成 3000 万像素，达到上限会列出受影响页码。原始扫描件本身的清晰度决定可见细节。保版保留打印外观，PDF 链接、可填写表单、批注交互和数字签名不会转换为 Word 的对应功能。可见表单内容的还原取决于 PDF 的页面渲染结果。

Word 的页面最长边为 55.88 厘米。遇到超大页时，保版会提示对应页码并终止，避免静默缩小页面。输入密码仅用于保版和 OCR 的本地 PDF 打开操作。Word 重排模式需要无需打开密码的副本。

可编辑模式会在整份 PDF 中查找线框表格。找到后，每一页按原尺寸建立一个 Word 节，表格按边线重建行、列及合并单元格，正文按原稿行的位置排版。支持的目录链接指向对应页书签。扫描页或图章页可保留整页图像；含插图或公式的页面可将可提取的文字与背景分别保存。转换结束时会显示表格数、单元格数以及整页使用图像的页码。PDF 数字签名的验证仍需使用原 PDF。

空白单元格可直接填写。固定行高用于保持打印版式；填写较长内容时，可在 Word 的「表格属性 → 行」中放宽行高。未识别成单元格的线框随图形保留，公式、扫描文字和复杂图形的编辑能力取决于原稿的编码。

扫描识别优先使用已安装的简体中文 OCR，也可回退到系统可用的识别语言。若缺少所需语言，可在 Windows 语言设置中添加对应 OCR 组件。签名、金额、小数点、编号、表格的阅读顺序请对照原件核查。

## 命令行

以下示例供 PowerShell 使用。图形程序应通过管道运行，以便脚本等待转换结束。`$LASTEXITCODE` 为 0 表示成功。输入输出路径支持空格和中文。已有同名文件时会报错。

```powershell
.\PDFtoWord.exe fixed "巡检记录.pdf" "巡检记录-保版.docx" 300 | Out-Host
.\PDFtoWord.exe edit "巡检记录.pdf" "巡检记录-可编辑.docx" | Out-Host
.\PDFtoWord.exe ocr "扫描签收单.pdf" "扫描签收单-文字.docx" | Out-Host
```

## 从源码构建

在 Windows x64 上使用 .NET Framework 4.8 和 Windows 10/11 SDK。于 64 位 PowerShell 中运行 `.\build.ps1`，输出为 `bin\PDFtoWord.exe`。亦可用 `-WindowsMetadata` 指定 SDK 的 `UnionMetadata\<版本>\Windows.winmd`。首次构建从已锁定版本的 pypdfium2 Windows wheel 提取 PDFium，并校验 wheel 与 DLL 的 SHA-256。依赖、版权声明均嵌入 EXE，成品离线使用。首次使用可编辑模式会把 PDFium 解包到当前用户的本地应用数据目录。

## 样例与验证

`samples` 中是本项目生成的测试 PDF。巡检记录含中英文表格、旋转裁切横页及扫描签收单，学术短文含双栏、上下标和脚注，表单含文字框和已勾选的复选框。运行 `.\tests\smoke.ps1` 可验证三页尺寸、图像数量和同名文件保护。对排版要求较高时，可用桌面 Word 打开保版结果，再导出 PDF，与原稿按相同分辨率逐页比较。

`tests/editable_form.py` 将 `samples/qualification.pdf` 交给实际 EXE，逐格比较源 PDF 与 DOCX 的文字和合并范围，并试填所有空白格后重新读取。运行此开发测试需要 Python、pdfplumber 与 python-docx 1.2 或更新版本。另可在装有 Word 的桌面运行 `tests/word_pagination.ps1`，验证备注、注释和开户银行字段所在页。运行 `.\PDFtoWord.exe --licenses | Out-Host` 可查看内置 PDF 引擎的第三方声明。

`python tests/mixed_document.py --word` 在临时目录生成含表单控件、正文、合并单元格和横向扫描页的 PDF，运行实际 EXE，并用桌面 Word 导出 PDF 检查分页、表格填写与图像位置。开发依赖包括 reportlab、Pillow、pypdfium2、numpy 和 python-docx。

保版使用 Windows PDF 渲染，电子表单使用 PDFium 提取文字与绘图路径。普通文档使用 Word 的 PDF 导入功能。[PDFium 项目](https://pdfium.googlesource.com/pdfium/) · [微软对 PDF 转换能力的说明](https://support.microsoft.com/en-us/word/opening-pdfs-in-word)
