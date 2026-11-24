using System;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;

internal static class NativePdfium
{
    private const string Hash = "42D87F68957D745CE9377DFB1D8D63691E19082A7EC29FF421D22779D3142B29";
    private static readonly object Gate = new object();
    private static IntPtr module;

    [DllImport("kernel32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    private static extern IntPtr LoadLibraryEx(string path, IntPtr reserved, uint flags);

    private static bool Matches(string file)
    {
        if (!File.Exists(file)) return false;
        using (var stream = File.OpenRead(file))
        using (var sha = SHA256.Create())
            return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-","")==Hash;
    }

    internal static bool TryLoad()
    {
        // Keep Word import available if desktop policy blocks user-local native binaries.
        try { EnsureLoaded(); return true; }
        catch (UnauthorizedAccessException) { return false; }
        catch (IOException) { return false; }
        catch (Win32Exception) { return false; }
    }

    private static void EnsureLoaded()
    {
        lock (Gate)
        {
            if (module!=IntPtr.Zero) return;
            var directory=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PDFtoWord","runtime",Hash);
            Directory.CreateDirectory(directory);
            string dll=Path.Combine(directory,"pdfium.dll");
            using (var mutex=new Mutex(false,@"Local\PDFtoWord-pdfium-"+Hash))
            {
                try { mutex.WaitOne(); } catch(AbandonedMutexException) { }
                try
                {
                    if (!Matches(dll))
                    {
                        string staged=Path.Combine(directory,Guid.NewGuid().ToString("N")+".tmp");
                        try
                        {
                            using (var source=Assembly.GetExecutingAssembly().GetManifestResourceStream("pdfium.dll"))
                            using (var destination=new FileStream(staged,FileMode.CreateNew,FileAccess.Write,FileShare.None))
                                source.CopyTo(destination);
                            if (!Matches(staged)) throw new InvalidDataException("内置 PDF 引擎校验失败。");
                            if(File.Exists(dll)) File.Replace(staged,dll,null);
                            else File.Move(staged,dll);
                        }
                        finally { if (File.Exists(staged)) File.Delete(staged); }
                    }
                    // Resolve dependencies only from the native binary's folder and System32.
                    module=LoadLibraryEx(dll,IntPtr.Zero,0x100 | 0x800);
                    if(module==IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
                }
                finally { mutex.ReleaseMutex(); }
            }
        }
    }
}
