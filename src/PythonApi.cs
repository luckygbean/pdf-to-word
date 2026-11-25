using System;
using System.Threading.Tasks;

namespace PdfToWord.Engine
{
    /// <summary>In-process entry point for Python.NET on Windows x64 and .NET Framework.</summary>
    public static class PdfConverter
    {
        public static string Convert(string inputPath, string outputPath, string mode = "edit",
            int dpi = 300, string password = null)
        {
            if (String.IsNullOrWhiteSpace(inputPath)) throw new ArgumentException("A PDF input path is required.", "inputPath");
            if (String.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("A DOCX output path is required.", "outputPath");
            if (dpi != 300 && dpi != 450 && dpi != 600) throw new ArgumentOutOfRangeException("dpi", "Use 300, 450 or 600 DPI.");
            if ((mode == "edit" || mode == "form") && !String.IsNullOrEmpty(password))
                throw new ArgumentException("Editable conversion requires a PDF without an opening password.", "password");
            // Run WinRT work on a pool thread, independently of a Python GUI's
            // apartment and synchronization context. Propagate the original exception.
            return Task.Run(() => Converter.RunAsync(mode,inputPath,outputPath,dpi,password,(i,n) => {}))
                .GetAwaiter().GetResult();
        }
    }
}
