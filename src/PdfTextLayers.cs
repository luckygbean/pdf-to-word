using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

internal static partial class PdfTextReader
{
    private static class Paint
    {
        private const string Dll="pdfium.dll";
        [DllImport(Dll)] internal static extern int FPDFPageObj_SetIsActive(IntPtr obj,int active);
        [DllImport(Dll)] internal static extern IntPtr FPDFBitmap_CreateEx(int w,int h,int format,IntPtr buffer,int stride);
        [DllImport(Dll)] internal static extern int FPDFBitmap_FillRect(IntPtr bitmap,int left,int top,int w,int h,uint color);
        [DllImport(Dll)] internal static extern void FPDFBitmap_Destroy(IntPtr bitmap);
        [DllImport(Dll)] internal static extern void FPDF_RenderPageBitmap(IntPtr bitmap,IntPtr page,int x,int y,int w,int h,int rotation,int flags);
        [DllImport(Dll)] internal static extern IntPtr FPDFPage_GetAnnot(IntPtr page,int index);
        [DllImport(Dll)] internal static extern int FPDFAnnot_GetSubtype(IntPtr annot);
        [DllImport(Dll)] internal static extern void FPDFPage_CloseAnnot(IntPtr annot);
    }

    private static bool OnlyLinks(IntPtr page)
    {
        int count=Api.FPDFPage_GetAnnotCount(page);
        for(int i=0;i<count;i++)
        {
            var annot=Paint.FPDFPage_GetAnnot(page,i);
            if(annot==IntPtr.Zero) return false;
            try { if(Paint.FPDFAnnot_GetSubtype(annot)!=2) return false; }
            finally { Paint.FPDFPage_CloseAnnot(annot); }
        }
        return true;
    }

    private static bool Intersects(Rect a,Rect b)
    {
        return Math.Min(a.Right,b.Right)>Math.Max(a.Left,b.Left)+.01 &&
               Math.Min(a.Top,b.Top)>Math.Max(a.Bottom,b.Bottom)+.01;
    }

    // Visit in reverse paint order. Moving an occluded label above an image or
    // solid masking rectangle can disclose text which was invisible in the PDF.
    private static HashSet<IntPtr> VisibleText(IntPtr page,Rect box)
    {
        var result=new HashSet<IntPtr>();
        var blockers=new List<Rect>();
        for(int i=Api.FPDFPage_CountObjects(page)-1;i>=0;i--)
        {
            var obj=Api.FPDFPage_GetObject(page,i);
            Rect bounds;
            int active;
            if(obj==IntPtr.Zero || Api.FPDFPageObj_GetIsActive(obj,out active)==0) return null;
            if(active==0) continue;
            if(Api.FPDFPageObj_GetBounds(obj,out bounds.Left,out bounds.Bottom,out bounds.Right,out bounds.Top)==0) return null;
            if(bounds.Right-bounds.Left<.01 || bounds.Top-bounds.Bottom<.01) continue;
            if(Api.FPDFPageObj_GetType(obj)==1)
            {
                int mode=Api.FPDFTextObj_GetTextRenderMode(obj);
                Matrix matrix;
                uint color;
                if((mode==0 || mode==2) && Api.FPDFPageObj_HasTransparency(obj)==0 && ClipContainsObject(obj) &&
                   Api.FPDFPageObj_GetMatrix(obj,out matrix)!=0 && Math.Abs(matrix.B)<.001 && Math.Abs(matrix.C)<.001 &&
                   matrix.A>0 && matrix.D>0 && Gray(obj,true,out color) && color<=40 &&
                   (mode!=2 || Gray(obj,false,out color) && color<=40) &&
                   Inside(new Point(bounds.Left,bounds.Bottom),box) && Inside(new Point(bounds.Right,bounds.Top),box) &&
                   !blockers.Any(b=>Intersects(b,bounds)))
                {
                    result.Add(obj);
                    continue;
                }
            }
            blockers.Add(bounds);
        }
        return result;
    }

    // The bitmap stores only graphics and text we could not reconstruct. Edits
    // to native text therefore never reveal a second copy underneath.
    internal static LayoutPage ReadLayers(IntPtr page,Stream cleanBackground)
    {
        Rect box;
        if(Api.FPDFPage_GetRotation(page)!=0 || !OnlyLinks(page) || Api.FPDF_GetPageBoundingBox(page,out box)==0)
            return null;
        var layout=new LayoutPage{Width=box.Right-box.Left, Height=box.Top-box.Bottom};
        if(!(layout.Width>0 && layout.Height>0 && layout.Width<=1584 && layout.Height<=1584)) return null;
        var objects=VisibleText(page,box);
        if(objects==null || objects.Count==0) return null;
        var textPage=Api.FPDFText_LoadPage(page);
        if(textPage==IntPtr.Zero) return null;
        try { if(!ReadCharacters(textPage,box,layout,objects,true) || layout.Lines.Count==0) return null; }
        finally { Api.FPDFText_ClosePage(textPage); }
        foreach(var obj in objects) if(Paint.FPDFPageObj_SetIsActive(obj,0)==0) return null;
        // Changes exist only on the current in-memory page. Session.WithPage
        // closes the handle immediately after this callback without saving.
        double scale=Math.Min(3,Math.Sqrt(16000000/(layout.Width*layout.Height)));
        int width=Math.Max(1,(int)Math.Round(layout.Width*scale)),height=Math.Max(1,(int)Math.Round(layout.Height*scale));
        using(var bitmap=new Bitmap(width,height,PixelFormat.Format32bppRgb))
        {
            var data=bitmap.LockBits(new Rectangle(0,0,width,height),ImageLockMode.WriteOnly,PixelFormat.Format32bppRgb);
            try
            {
                var native=Paint.FPDFBitmap_CreateEx(width,height,3,data.Scan0,data.Stride);
                if(native==IntPtr.Zero) throw new IOException("无法分配 PDF 页面图像。");
                try
                {
                    if(Paint.FPDFBitmap_FillRect(native,0,0,width,height,0xffffffff)==0)
                        throw new IOException("无法初始化 PDF 页面图像。");
                    Paint.FPDF_RenderPageBitmap(native,page,0,0,width,height,0,1);
                }
                finally { Paint.FPDFBitmap_Destroy(native); }
            }
            finally { bitmap.UnlockBits(data); }
            bitmap.Save(cleanBackground,ImageFormat.Png);
        }
        cleanBackground.Position=0;
        return layout;
    }
}
