using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length==0 || args.Length==1 && args[0].EndsWith(".pdf",StringComparison.OrdinalIgnoreCase))
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(args.Length==1 ? args[0] : ""));
            return 0;
        }
        try
        {
            Console.SetOut(new StreamWriter(Console.OpenStandardOutput(),new UTF8Encoding(false)) { AutoFlush=true });
            Console.SetError(new StreamWriter(Console.OpenStandardError(),new UTF8Encoding(false)) { AutoFlush=true });
            if (args.Length==1 && args[0]=="--licenses")
            {
                using (var stream=Assembly.GetExecutingAssembly().GetManifestResourceStream("pdfium.NOTICES.txt"))
                using (var notices=new StreamReader(stream,Encoding.UTF8)) Console.Write(notices.ReadToEnd());
                return 0;
            }
            if (args.Length<3 || args.Length>4) throw new ArgumentException("PDFtoWord.exe fixed|edit|form|ocr input.pdf output.docx [300|450|600]");
            int dpi = args.Length==4 ? Int32.Parse(args[3]) : 300;
            if (dpi!=300 && dpi!=450 && dpi!=600) throw new ArgumentException("支持 300、450、600 DPI。");
            var message = Task.Run(()=>Converter.RunAsync(args[0],args[1],args[2],dpi,null,(i,n)=>{})).GetAwaiter().GetResult();
            Console.WriteLine(message);
            return 0;
        }
        catch (Exception e) { Console.Error.WriteLine(e.Message); return 1; }
    }
}
