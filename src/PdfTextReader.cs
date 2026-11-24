using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

// Read the visible crop in PDF points. A null result asks the caller to use its
// general-purpose converter instead of silently dropping unsupported artwork.
internal static class PdfTextReader
{
    private static readonly object LibraryLock = new object();

    internal static List<LayoutPage> Read(string pdf)
    {
        lock (LibraryLock)
        {
            if (!NativePdfium.TryLoad()) return null;
            Api.FPDF_InitLibrary();
            try
            {
                using (var file = new FileStream(pdf, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if (file.Length == 0) return null;
                    using (var mapping = MemoryMappedFile.CreateFromFile(file, null, 0, MemoryMappedFileAccess.Read,
                                                                         HandleInheritability.None, true))
                    using (var view = mapping.CreateViewAccessor(0, 0, MemoryMappedFileAccess.Read))
                    {
                        var doc = Api.FPDF_LoadMemDocument64(view.SafeMemoryMappedViewHandle.DangerousGetHandle(),
                                                            (ulong)file.Length, IntPtr.Zero);
                        if (doc == IntPtr.Zero) return null;
                        try { return ReadPages(doc); }
                        finally { Api.FPDF_CloseDocument(doc); }
                    }
                }
            }
            finally { Api.FPDF_DestroyLibrary(); }
        }
    }

    private static List<LayoutPage> ReadPages(IntPtr document)
    {
        int count = Api.FPDF_GetPageCount(document);
        if (count <= 0 || Api.FPDF_GetFormType(document) != 0) return null;
        var pages = new List<LayoutPage>();
        for (int i = 0; i < count; ++i)
        {
            var page = Api.FPDF_LoadPage(document, i);
            if (page == IntPtr.Zero) return null;
            try
            {
                var layout = ReadPage(page);
                if (layout == null) return null;
                pages.Add(layout);
            }
            finally { Api.FPDF_ClosePage(page); }
        }
        return pages;
    }

    private static LayoutPage ReadPage(IntPtr page)
    {
        Rect box;
        if (Api.FPDF_GetPageBoundingBox(page, out box) == 0 || Api.FPDFPage_GetRotation(page) != 0 ||
            Api.FPDFPage_GetAnnotCount(page) != 0) return null;
        var result = new LayoutPage { Width = box.Right-box.Left, Height = box.Top-box.Bottom };
        if (!(result.Width > 0 && result.Width <= 1584 && result.Height > 0 && result.Height <= 1584)) return null;
        int objects = Api.FPDFPage_CountObjects(page);
        bool painted = false;
        var textObjects = new HashSet<IntPtr>();
        for (int i = 0; i < objects; ++i)
        {
            var obj = Api.FPDFPage_GetObject(page, i);
            int active;
            if (obj == IntPtr.Zero || Api.FPDFPageObj_GetIsActive(obj, out active) == 0) return null;
            if (active == 0) return null; // The text stream may still expose hidden glyphs.
            if (Api.FPDFPageObj_HasTransparency(obj) != 0 || !ClipContainsObject(obj)) return null;
            switch (Api.FPDFPageObj_GetType(obj))
            {
                case 1:
                    int mode = Api.FPDFTextObj_GetTextRenderMode(obj);
                    uint f;
                    if ((mode != 0 && mode != 2) || !Gray(obj,true,out f) || f > 40) return null;
                    if (mode == 2 && (!Gray(obj,false,out f) || f > 40)) return null;
                    float l,b,r,t;
                    if (Api.FPDFPageObj_GetBounds(obj,out l,out b,out r,out t)==0) return null;
                    if (r-l>.01 && t-b>.01) textObjects.Add(obj);
                    painted = true;
                    break;
                case 2:
                    if (!ReadPath(obj, box, result.Rules, ref painted)) return null;
                    break;
                default: return null;
            }
        }
        var text = Api.FPDFText_LoadPage(page);
        if (text == IntPtr.Zero) return null;
        try { return ReadCharacters(text, box, result, textObjects) ? result : null; }
        finally { Api.FPDFText_ClosePage(text); }
    }

    private struct Point
    {
        internal double X, Y;
        internal Point(double x, double y) { X=x; Y=y; }
    }

    private static bool AxisAlignedRectangle(List<Point> points)
    {
        if (points.Count != 4 && points.Count != 5) return false;
        if (points.Count == 5 && (Math.Abs(points[4].X-points[0].X) > .01 ||
                                  Math.Abs(points[4].Y-points[0].Y) > .01)) return false;
        for (int i = 0; i < 4; ++i)
        {
            var a = points[i]; var b = points[(i+1) % 4]; var c = points[(i+2) % 4];
            bool horizontal = Math.Abs(a.Y-b.Y) < .01;
            if ((horizontal == (Math.Abs(a.X-b.X)<.01)) ||
                (horizontal == (Math.Abs(b.Y-c.Y)<.01))) return false;
        }
        return true;
    }

    // Office PDFs often clip each label to its cell. Such a clip is redundant if
    // its rectangle encloses the full object. Other clips require general reflow.
    private static bool ClipContainsObject(IntPtr obj)
    {
        var clip=Api.FPDFPageObj_GetClipPath(obj);
        int paths = clip == IntPtr.Zero ? 0 : Api.FPDFClipPath_CountPaths(clip);
        if (paths <= 0) return true;
        float left,bottom,right,top;
        if (Api.FPDFPageObj_GetBounds(obj,out left,out bottom,out right,out top)==0) return false;
        for (int c=0; c<paths; ++c)
        {
            int segments = Api.FPDFClipPath_CountPathSegments(clip,c);
            if (segments!=4 && segments!=5) return false;
            var points = new List<Point>();
            for (int i=0; i<segments; ++i)
            {
                var seg=Api.FPDFClipPath_GetPathSegment(clip,c,i);
                float x,y;
                if (Api.FPDFPathSegment_GetType(seg)!=(i==0?2:0) ||
                    Api.FPDFPathSegment_GetPoint(seg,out x,out y)==0) return false;
                points.Add(new Point(x,y));
            }
            if (!AxisAlignedRectangle(points) || left < points.Min(p=>p.X)-.1 || right > points.Max(p=>p.X)+.1 ||
                bottom < points.Min(p=>p.Y)-.1 || top > points.Max(p=>p.Y)+.1) return false;
        }
        return true;
    }

    private static bool Gray(IntPtr obj, bool fill, out uint value)
    {
        uint g,b,a;
        int ok = fill ? Api.FPDFPageObj_GetFillColor(obj,out value,out g,out b,out a)
                      : Api.FPDFPageObj_GetStrokeColor(obj,out value,out g,out b,out a);
        return ok!=0 && a==255 && value==g && g==b;
    }

    private static bool AddRule(List<LayoutRule> rules, Point a, Point b, double width, uint gray, Rect box)
    {
        if (Math.Abs(a.X-b.X)>.04 && Math.Abs(a.Y-b.Y)>.04) return false;
        if (!Inside(a,box) || !Inside(b,box) || !Finite(width) || width<0 || width>6) return false;
        if (Math.Abs(a.X-b.X)<.02 && Math.Abs(a.Y-b.Y)<.02) return true;
        rules.Add(new LayoutRule { Left=Math.Min(a.X,b.X)-box.Left, Right=Math.Max(a.X,b.X)-box.Left,
            Top=box.Top-Math.Max(a.Y,b.Y), Bottom=box.Top-Math.Min(a.Y,b.Y), Width=Math.Max(.2,width), Gray=(int)gray });
        return true;
    }

    private static bool Inside(Point p, Rect box)
    {
        return p.X>=box.Left-.25 && p.X<=box.Right+.25 && p.Y>=box.Bottom-.25 && p.Y<=box.Top+.25;
    }
    private static bool Finite(double value) { return !double.IsInfinity(value) && !double.IsNaN(value); }

    private static bool ReadPath(IntPtr path, Rect box, List<LayoutRule> rules, ref bool painted)
    {
        Matrix m;
        int fill,stroke;
        if (Api.FPDFPageObj_GetMatrix(path,out m)==0 || Api.FPDFPath_GetDrawMode(path,out fill,out stroke)==0) return false;
        if (fill==0 && stroke==0) return true;
        uint fillGray=0, strokeGray=0;
        float pen=0;
        if (fill!=0 && !Gray(path,true,out fillGray)) return false;
        if (stroke!=0 && (!Gray(path,false,out strokeGray) || Api.FPDFPageObj_GetStrokeWidth(path,out pen)==0 ||
                          Api.FPDFPageObj_GetDashCount(path)!=0)) return false;
        var polygons=new List<List<Point>>();
        List<Point> vertices=null;
        int count=Api.FPDFPath_CountSegments(path);
        for (int i=0; i<count; ++i)
        {
            var seg=Api.FPDFPath_GetPathSegment(path,i);
            int type=Api.FPDFPathSegment_GetType(seg);
            float x,y;
            if ((type!=2 && type!=0) || Api.FPDFPathSegment_GetPoint(seg,out x,out y)==0) return false;
            if (type==2) { vertices=new List<Point>(); polygons.Add(vertices); }
            if (vertices==null) return false;
            vertices.Add(new Point(x*m.A+y*m.C+m.E,x*m.B+y*m.D+m.F));
            if (Api.FPDFPathSegment_GetClose(seg)!=0)
            {
                Point last=vertices[vertices.Count-1], first=vertices[0];
                if (Math.Abs(last.X-first.X)>.001 || Math.Abs(last.Y-first.Y)>.001) vertices.Add(first);
            }
        }
        foreach (var points in polygons)
        {
            if (fill!=0)
            {
                // Word often paints a border as a very thin filled rectangle.
                if (!AxisAlignedRectangle(points)) return false;
                double left=points.Min(p=>p.X), right=points.Max(p=>p.X);
                double bottom=points.Min(p=>p.Y), top=points.Max(p=>p.Y);
                double width=right-left, height=top-bottom;
                if (fillGray==255 && !painted && stroke==0) continue;
                bool junction = Math.Max(width,height)<=1.5;
                if (Math.Min(width,height)>4 || (!junction && Math.Max(width,height)<4*Math.Min(width,height)) || stroke!=0) return false;
                Point a,b;
                if (width>=height) { a=new Point(left,(top+bottom)/2); b=new Point(right,a.Y); }
                else { a=new Point((left+right)/2,bottom); b=new Point(a.X,top); }
                if (!AddRule(rules,a,b,Math.Min(width,height),fillGray,box)) return false;
            }
            if (stroke!=0)
            {
                double scaling=Math.Sqrt(Math.Abs(m.A*m.D-m.B*m.C));
                for (int i=1; i<points.Count; ++i)
                    if (!AddRule(rules,points[i-1],points[i],pen*scaling,strokeGray,box)) return false;
            }
            painted=true;
        }
        return true;
    }

    private static bool ReadCharacters(IntPtr text, Rect box, LayoutPage result, HashSet<IntPtr> objects)
    {
        LayoutLine run = null;
        var buffer = new StringBuilder();
        var fontNames = new Dictionary<IntPtr,string>();
        var seen = new HashSet<IntPtr>();
        IntPtr current = IntPtr.Zero;
        int count = Api.FPDFText_CountChars(text);
        if (count < 0) return false;
        for (int i=0; i<count; ++i)
        {
            if (Api.FPDFText_IsGenerated(text,i)==1) continue;
            uint code=Api.FPDFText_GetUnicode(text,i);
            if (code==0 || code>0x10ffff || code==0xfffd || code==0xfffe || code==0xffff ||
                (code>=0xd800 && code<=0xdfff) || code<32 ||
                Api.FPDFText_HasUnicodeMapError(text,i)!=0) return false;
            var obj=Api.FPDFText_GetTextObject(text,i);
            if (obj==IntPtr.Zero || !objects.Contains(obj))
            {
                if (char.IsWhiteSpace(char.ConvertFromUtf32((int)code),0)) continue;
                return false;
            }
            seen.Add(obj);
            Matrix matrix;
            if (Api.FPDFText_GetMatrix(text,i,out matrix)==0 || Math.Abs(matrix.B)>.001 ||
                Math.Abs(matrix.C)>.001 || matrix.A<=0 || matrix.D<=0) return false;
            double size=Api.FPDFText_GetFontSize(text,i)*matrix.D;
            double ox,oy,left,right,bot,top;
            if (!(size>=1 && size<=300) || Api.FPDFText_GetCharOrigin(text,i,out ox,out oy)==0 ||
                Api.FPDFText_GetCharBox(text,i,out left,out right,out bot,out top)==0) return false;
            if (!Inside(new Point(left,bot),box) || !Inside(new Point(right,top),box)) return false;
            // Word lays out glyphs by advances. Ink bounds would compress
            // leading/trailing full-width punctuation and offset centered labels.
            Rect loose;
            if (Api.FPDFText_GetLooseCharBox(text,i,out loose)==0 ||
                !Finite(loose.Left) || !Finite(loose.Right) || loose.Right<loose.Left) return false;
            left=loose.Left;
            right=loose.Right;
            double baseline=box.Top-oy;
            // Distinct PDF text objects frequently live in adjacent cells. Keeping
            // them independent lets table reconstruction use the actual geometry.
            if (run==null || current!=obj || Math.Abs(run.Baseline-baseline)>.35 ||
                left-box.Left > run.Right+size*1.25 || right-box.Left < run.Left-.5)
            {
                FinishRun(run,buffer,result);
                var font=Api.FPDFTextObj_GetFont(obj);
                string family;
                if (!fontNames.TryGetValue(font,out family))
                {
                    var bytes=new byte[512];
                    ulong length=Api.FPDFFont_GetFamilyName(font,bytes,(UIntPtr)bytes.Length).ToUInt64();
                    if (length<=1 || length>(ulong)bytes.Length) return false;
                    family=Encoding.UTF8.GetString(bytes,0,(int)length-1);
                    fontNames.Add(font,family);
                }
                run=new LayoutLine { Left=left-box.Left, Top=box.Top-top, Right=right-box.Left, Bottom=box.Top-bot,
                    FontSize=size, FontName=family, Bold=Api.FPDFText_GetFontWeight(text,i)>=600 ||
                    Api.FPDFTextObj_GetTextRenderMode(obj)==2, Baseline=baseline };
                current=obj;
            }
            run.Left=Math.Min(run.Left,left-box.Left); run.Right=Math.Max(run.Right,right-box.Left);
            run.Top=Math.Min(run.Top,box.Top-top); run.Bottom=Math.Max(run.Bottom,box.Top-bot);
            buffer.Append(char.ConvertFromUtf32((int)code));
        }
        FinishRun(run,buffer,result);
        return seen.SetEquals(objects);
    }

    private static void FinishRun(LayoutLine run, StringBuilder buffer, LayoutPage page)
    {
        if (run!=null && buffer.Length>0)
        {
            run.Text=buffer.ToString();
            if (!String.IsNullOrWhiteSpace(run.Text)) page.Lines.Add(run);
        }
        buffer.Clear();
    }

    [StructLayout(LayoutKind.Sequential)] private struct Matrix { internal float A,B,C,D,E,F; }
    [StructLayout(LayoutKind.Sequential)] private struct Rect { internal float Left,Top,Right,Bottom; }

    private static class Api
    {
        private const string Dll="pdfium.dll";
        [DllImport(Dll)] internal static extern void FPDF_InitLibrary();
        [DllImport(Dll)] internal static extern void FPDF_DestroyLibrary();
        [DllImport(Dll)] internal static extern IntPtr FPDF_LoadMemDocument64(IntPtr data, ulong size, IntPtr password);
        [DllImport(Dll)] internal static extern void FPDF_CloseDocument(IntPtr doc);
        [DllImport(Dll)] internal static extern int FPDF_GetFormType(IntPtr doc);
        [DllImport(Dll)] internal static extern int FPDF_GetPageCount(IntPtr doc);
        [DllImport(Dll)] internal static extern IntPtr FPDF_LoadPage(IntPtr doc,int i);
        [DllImport(Dll)] internal static extern void FPDF_ClosePage(IntPtr page);
        [DllImport(Dll)] internal static extern int FPDF_GetPageBoundingBox(IntPtr page,out Rect rect);
        [DllImport(Dll)] internal static extern int FPDFPage_GetRotation(IntPtr page);
        [DllImport(Dll)] internal static extern int FPDFPage_GetAnnotCount(IntPtr page);
        [DllImport(Dll)] internal static extern int FPDFPage_CountObjects(IntPtr page);
        [DllImport(Dll)] internal static extern IntPtr FPDFPage_GetObject(IntPtr page,int i);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetType(IntPtr obj);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetIsActive(IntPtr obj,out int active);
        [DllImport(Dll)] internal static extern int FPDFPageObj_HasTransparency(IntPtr obj);
        [DllImport(Dll)] internal static extern IntPtr FPDFPageObj_GetClipPath(IntPtr obj);
        [DllImport(Dll)] internal static extern int FPDFClipPath_CountPaths(IntPtr path);
        [DllImport(Dll)] internal static extern int FPDFClipPath_CountPathSegments(IntPtr path,int i);
        [DllImport(Dll)] internal static extern IntPtr FPDFClipPath_GetPathSegment(IntPtr path,int i,int segment);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetBounds(IntPtr obj,out float l,out float b,out float r,out float t);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetMatrix(IntPtr obj,out Matrix m);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetFillColor(IntPtr obj,out uint r,out uint g,out uint b,out uint a);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetStrokeColor(IntPtr obj,out uint r,out uint g,out uint b,out uint a);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetStrokeWidth(IntPtr obj,out float w);
        [DllImport(Dll)] internal static extern int FPDFPageObj_GetDashCount(IntPtr obj);
        [DllImport(Dll)] internal static extern int FPDFPath_GetDrawMode(IntPtr obj,out int fill,out int stroke);
        [DllImport(Dll)] internal static extern int FPDFPath_CountSegments(IntPtr path);
        [DllImport(Dll)] internal static extern IntPtr FPDFPath_GetPathSegment(IntPtr path,int i);
        [DllImport(Dll)] internal static extern int FPDFPathSegment_GetType(IntPtr segment);
        [DllImport(Dll)] internal static extern int FPDFPathSegment_GetPoint(IntPtr segment,out float x,out float y);
        [DllImport(Dll)] internal static extern int FPDFPathSegment_GetClose(IntPtr segment);
        [DllImport(Dll)] internal static extern IntPtr FPDFText_LoadPage(IntPtr page);
        [DllImport(Dll)] internal static extern void FPDFText_ClosePage(IntPtr page);
        [DllImport(Dll)] internal static extern int FPDFText_CountChars(IntPtr page);
        [DllImport(Dll)] internal static extern int FPDFText_IsGenerated(IntPtr page,int i);
        [DllImport(Dll)] internal static extern int FPDFText_HasUnicodeMapError(IntPtr page,int i);
        [DllImport(Dll)] internal static extern uint FPDFText_GetUnicode(IntPtr page,int i);
        [DllImport(Dll)] internal static extern int FPDFText_GetMatrix(IntPtr page,int i,out Matrix m);
        [DllImport(Dll)] internal static extern IntPtr FPDFText_GetTextObject(IntPtr page,int i);
        [DllImport(Dll)] internal static extern int FPDFText_GetCharOrigin(IntPtr page,int i,out double x,out double y);
        [DllImport(Dll)] internal static extern int FPDFText_GetCharBox(IntPtr page,int i,out double l,out double r,out double b,out double t);
        [DllImport(Dll)] internal static extern int FPDFText_GetLooseCharBox(IntPtr page,int i,out Rect rect);
        [DllImport(Dll)] internal static extern double FPDFText_GetFontSize(IntPtr page,int i);
        [DllImport(Dll)] internal static extern int FPDFText_GetFontWeight(IntPtr page,int i);
        [DllImport(Dll)] internal static extern int FPDFTextObj_GetTextRenderMode(IntPtr obj);
        [DllImport(Dll)] internal static extern IntPtr FPDFTextObj_GetFont(IntPtr obj);
        [DllImport(Dll)] internal static extern UIntPtr FPDFFont_GetFamilyName(IntPtr font,[Out] byte[] buffer,UIntPtr size);
    }
}
