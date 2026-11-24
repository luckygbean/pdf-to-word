using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Windows.Data.Pdf;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;

public static class OcrText
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly Regex CjkSpace = new Regex(@"(?<=[\u3000-\u303f\u3400-\u9fff\uf900-\ufaff\uff01-\uff60]) +(?=[\u3000-\u303f\u3400-\u9fff\uf900-\ufaff\uff01-\uff60])");

    // The caller owns the unique output path and discards the partial file on error.
    public static async Task<string> ConvertAsync(string input, string output, string password, Action<int, int> progress)
    {
        OcrEngine engine = OcrEngine.TryCreateFromLanguage(new Language("zh-Hans"))
            ?? OcrEngine.TryCreateFromUserProfileLanguages();
        if (engine == null && OcrEngine.AvailableRecognizerLanguages.Count > 0)
            engine = OcrEngine.TryCreateFromLanguage(OcrEngine.AvailableRecognizerLanguages[0]);
        if (engine == null)
            throw new InvalidOperationException("Windows 缺少 OCR 识别语言。请在 Windows 设置的“语言和区域”中安装所需语言的“光学字符识别”组件。");

        var file = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(input));
        using (var stream = await file.OpenAsync(FileAccessMode.Read))
        {
            var pdf = string.IsNullOrEmpty(password)
                ? await PdfDocument.LoadFromStreamAsync(stream)
                : await PdfDocument.LoadFromStreamAsync(stream, password);
            if (pdf.PageCount == 0)
                throw new InvalidDataException("PDF 没有页面。");

            var body = new XElement(W + "body");
            var imageRels = new XElement(Rel + "Relationships");
            var imagePages = new List<uint>();
            using (var outputFile = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (var zip = new ZipArchive(outputFile, ZipArchiveMode.Create))
            {
                for (uint i = 0; i < pdf.PageCount; i++)
                {
                    using (var page = pdf.GetPage(i))
                    using (var raster = new InMemoryRandomAccessStream())
                    {
                        // Cap both resolution and the OS model's maximum input dimension.
                        var scale = Math.Min(300.0 / 96, Math.Min(3508.0, OcrEngine.MaxImageDimension)
                            / Math.Max(page.Size.Width, page.Size.Height));
                        var options = new PdfPageRenderOptions {
                            DestinationWidth = (uint)Math.Max(1, Math.Round(page.Size.Width * scale)),
                            DestinationHeight = (uint)Math.Max(1, Math.Round(page.Size.Height * scale)),
                            BitmapEncoderId = BitmapEncoder.PngEncoderId,
                            IsIgnoringHighContrast = true
                        };
                        await page.RenderToStreamAsync(raster, options);
                        raster.Seek(0);
                        var decoder = await BitmapDecoder.CreateAsync(raster);
                        OcrResult result;
                        using (var bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied))
                            result = await engine.RecognizeAsync(bitmap);

                        var lines = result.Lines.Select(line => CleanLine(line.Text))
                            .Where(line => !string.IsNullOrWhiteSpace(line)).ToList();
                        bool startNewPage = i > 0;
                        if (lines.Count == 0)
                        {
                            uint number = i + 1;
                            string imageId = "img" + number.ToString(CultureInfo.InvariantCulture);
                            string target = "media/" + imageId + ".png";
                            using (var source = raster.GetInputStreamAt(0).AsStreamForRead())
                            using (var destination = zip.CreateEntry("word/" + target, CompressionLevel.NoCompression).Open())
                                source.CopyTo(destination);
                            imagePages.Add(number);
                            imageRels.Add(Relationship(imageId, R.NamespaceName + "/image", target));
                            body.Add(Paragraph(PageImage(number, imageId, page.Size.Width, page.Size.Height), startNewPage));
                        }
                        else
                        {
                            foreach (string line in lines)
                            {
                                var run = new XElement(W + "r",
                                    new XElement(W + "rPr",
                                        new XElement(W + "rFonts", new XAttribute(W + "ascii", "Calibri"),
                                            new XAttribute(W + "hAnsi", "Calibri"), new XAttribute(W + "eastAsia", "宋体")),
                                        new XElement(W + "sz", new XAttribute(W + "val", "22"))),
                                    new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), line));
                                body.Add(Paragraph(run, startNewPage));
                                startNewPage = false;
                            }
                        }
                    }
                    if (progress != null)
                        progress((int)i + 1, (int)pdf.PageCount);
                }

                body.Add(new XElement(W + "sectPr",
                    new XElement(W + "pgSz", new XAttribute(W + "w", 11906), new XAttribute(W + "h", 16838)),
                    new XElement(W + "pgMar", new XAttribute(W + "top", 1080), new XAttribute(W + "right", 1080),
                        new XAttribute(W + "bottom", 1080), new XAttribute(W + "left", 1080),
                        new XAttribute(W + "header", 0), new XAttribute(W + "footer", 0), new XAttribute(W + "gutter", 0))));
                WriteXml(zip, "word/document.xml", new XElement(W + "document",
                    new XAttribute(XNamespace.Xmlns + "w", W), body));
                WriteXml(zip, "_rels/.rels", new XElement(Rel + "Relationships",
                    Relationship("doc", R.NamespaceName + "/officeDocument", "word/document.xml")));
                if (imagePages.Count > 0)
                    WriteXml(zip, "word/_rels/document.xml.rels", imageRels);
                WriteXml(zip, "[Content_Types].xml", new XElement(Ct + "Types",
                    new XElement(Ct + "Default", new XAttribute("Extension", "rels"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                    new XElement(Ct + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", "image/png")),
                    new XElement(Ct + "Override", new XAttribute("PartName", "/word/document.xml"),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"))));
            }
            string note = "OCR 语言：" + engine.RecognizerLanguage.LanguageTag + "。请校对姓名、金额和表格列序。";
            if (imagePages.Count > 0)
                note += " 以下原稿页未识别到文字，已保留整页图像：" + string.Join("、", imagePages) + "。";
            return note;
        }
    }

    private static string CleanLine(string text)
    {
        var filtered = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++)
        {
            if (char.IsSurrogatePair(text, i))
                filtered.Append(text[i++]).Append(text[i]);
            else if (XmlConvert.IsXmlChar(text[i]))
                filtered.Append(text[i]);
        }
        return CjkSpace.Replace(filtered.ToString(), "");
    }

    private static XElement Paragraph(XElement run, bool startNewPage)
    {
        return new XElement(W + "p", new XElement(W + "pPr",
            startNewPage ? new XElement(W + "pageBreakBefore") : null,
            new XElement(W + "spacing", new XAttribute(W + "before", 0), new XAttribute(W + "after", 0))), run);
    }

    private static XElement PageImage(uint number, string id, double width, double height)
    {
        // Fit within the A4 text area, allowing room for Word's inline-image baseline.
        double scale = Math.Min((11906 - 2160) * 635.0 / width, (16838 - 2160 - 300) * 635.0 / height);
        long x = Math.Max(1L, (long)(width * scale));
        long y = Math.Max(1L, (long)(height * scale));
        var name = "原稿第 " + number + " 页";
        var picture = new XElement(Pic + "pic",
            new XElement(Pic + "nvPicPr", new XElement(Pic + "cNvPr", new XAttribute("id", number), new XAttribute("name", name)),
                new XElement(Pic + "cNvPicPr")),
            new XElement(Pic + "blipFill", new XElement(A + "blip", new XAttribute(R + "embed", id)),
                new XElement(A + "stretch", new XElement(A + "fillRect"))),
            new XElement(Pic + "spPr",
                new XElement(A + "xfrm", new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(A + "ext", new XAttribute("cx", x), new XAttribute("cy", y))),
                new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))));
        return new XElement(W + "r", new XElement(W + "drawing", new XElement(Wp + "inline",
            new XElement(Wp + "extent", new XAttribute("cx", x), new XAttribute("cy", y)),
            new XElement(Wp + "docPr", new XAttribute("id", number), new XAttribute("name", name)),
            new XElement(A + "graphic", new XElement(A + "graphicData", new XAttribute("uri", Pic.NamespaceName), picture)))));
    }

    private static XElement Relationship(string id, string type, string target)
    {
        return new XElement(Rel + "Relationship", new XAttribute("Id", id), new XAttribute("Type", type), new XAttribute("Target", target));
    }

    private static void WriteXml(ZipArchive zip, string name, XElement root)
    {
        using (var stream = zip.CreateEntry(name).Open())
            new XDocument(root).Save(stream, SaveOptions.DisableFormatting);
    }
}
