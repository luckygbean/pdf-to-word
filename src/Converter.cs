using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Windows.Data.Pdf;
using Windows.Storage;

internal static class Converter
{
    internal static async Task<string> RunAsync(string mode, string input, string target, int dpi, string password, Action<int,int> progress)
    {
        input = Path.GetFullPath(input);
        target = Path.GetFullPath(target);
        if (!File.Exists(input) || !String.Equals(Path.GetExtension(input),".pdf",StringComparison.OrdinalIgnoreCase))
            throw new IOException("请选择一个可以读取的 PDF 文件。");
        if (!String.Equals(Path.GetExtension(target),".docx",StringComparison.OrdinalIgnoreCase))
            throw new IOException("输出文件的扩展名须为 .docx。");
        if (File.Exists(target)) throw new IOException("已有同名文件，请选择新的 Word 文件名。");
        string temp = Path.Combine(Path.GetDirectoryName(target),".pdfword-"+Guid.NewGuid().ToString("N")+".docx");
        try
        {
            string report;
            switch (mode)
            {
                case "fixed": report = await FixedPages.ConvertAsync(input,temp,password,dpi,progress); break;
                case "ocr": report = await OcrText.ConvertAsync(input,temp,password,progress); break;
                case "edit":
                    report = await MixedDocument.TryConvertAsync(input,temp,progress);
                    if (report==null) report = await WordAsync(input,temp);
                    break;
                default: throw new ArgumentException("模式须为 fixed、edit 或 ocr。");
            }
            using (var zip = ZipFile.OpenRead(temp))
            {
                if (zip.GetEntry("[Content_Types].xml")==null || zip.GetEntry("word/document.xml")==null)
                    throw new InvalidDataException("生成的 Word 包不完整。");
            }
            File.Move(temp,target);
            return report;
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }

    private static string Quote(string path)
    {
        // A double quote cannot occur in a Windows file name. All arguments here are files.
        return "\"" + path + "\"";
    }

    private static async Task<string> WordAsync(string input, string temp)
    {
        // Detect a password requirement before starting unattended Word, which could display a hidden prompt.
        try
        {
            var file = await StorageFile.GetFileFromPathAsync(input);
            using (var stream = await file.OpenAsync(FileAccessMode.Read)) await PdfDocument.LoadFromStreamAsync(stream);
        }
        catch (Exception e)
        {
            if (e.HResult==unchecked((int)0x8007052B) || e.HResult==unchecked((int)0x80070056))
                throw new IOException("Word 重排需要无需打开密码的副本。也可在保留原版排版或扫描文字识别模式中输入密码。",e);
            throw;
        }
        string script = Path.Combine(Path.GetTempPath(),"pdfword-"+Guid.NewGuid().ToString("N")+".ps1");
        try
        {
            using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream("editable.ps1"))
            using (var dest = new FileStream(script, FileMode.CreateNew)) source.CopyTo(dest);
            var start = new ProcessStartInfo(Path.Combine(Environment.SystemDirectory,@"WindowsPowerShell\v1.0\powershell.exe"),
                "-NoLogo -NoProfile -NonInteractive -STA -ExecutionPolicy Bypass -File "+Quote(script)+" -InputPath "+Quote(input)+" -OutputPath "+Quote(temp));
            start.UseShellExecute = false;
            start.CreateNoWindow = true;
            start.RedirectStandardOutput = start.RedirectStandardError = true;
            start.StandardOutputEncoding = start.StandardErrorEncoding = Encoding.UTF8;
            using (var proc = Process.Start(start))
            {
                var normal = proc.StandardOutput.ReadToEndAsync();
                var errors = proc.StandardError.ReadToEndAsync();
                await Task.WhenAll(normal,errors);
                await Task.Run((Action)proc.WaitForExit);
                if (proc.ExitCode!=0)
                    throw new IOException("Word 转换失败。"+errors.Result.Trim());
            }
            using (var zip = ZipFile.OpenRead(temp))
            using (var stream = zip.GetEntry("word/document.xml").Open())
            {
                var doc = XDocument.Load(stream);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var chars = doc.Descendants(w+"t").Sum(t=>t.Value.Trim().Length);
                var tables = doc.Descendants(w+"tbl").Count();
                return chars==0 ? "Word 未提取到可编辑文字。扫描页如需改字，请选择「扫描文字识别」。" :
                    "已通过桌面版 Word 重排，正文有 "+chars+" 个文字字符、"+tables+" 张表格。请对照原稿检查分页与合并单元格。";
            }
        }
        finally { if (File.Exists(script)) File.Delete(script); }
    }
}
