using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

internal static partial class PdfTextReader
{
    private static class StampApi
    {
        private const string Dll=NativePdfium.LibraryName;
        [DllImport(Dll)] internal static extern IntPtr FPDF_CreateNewDocument();
        [DllImport(Dll)] internal static extern int FPDF_ImportPagesByIndex(IntPtr dest,IntPtr source,ref int index,uint count,int at);
        [DllImport(Dll)] internal static extern int FPDFPage_RemoveObject(IntPtr page,IntPtr obj);
        [DllImport(Dll)] internal static extern void FPDFPageObj_Destroy(IntPtr obj);
        [DllImport(Dll)] internal static extern int FPDFPage_GenerateContent(IntPtr page);
        [DllImport(Dll)] internal static extern int FPDFPage_RemoveAnnot(IntPtr page,int index);
        [DllImport(Dll)] internal static extern int FPDFPage_Flatten(IntPtr page,int usage);
        [DllImport(Dll)] internal static extern int FPDFAnnot_GetFlags(IntPtr annot);
        [DllImport(Dll,CharSet=CharSet.Ansi,ExactSpelling=true)]
        internal static extern uint FPDFAnnot_GetStringValue(IntPtr annot,string key,[Out] byte[] value,uint bytes);
        [DllImport(Dll)] internal static extern IntPtr FPDFDOC_InitFormFillEnvironment(IntPtr doc,IntPtr info);
        [DllImport(Dll)] internal static extern int FPDFAnnot_GetFormFieldType(IntPtr form,IntPtr annot);
        [DllImport(Dll)] internal static extern void FPDFDOC_ExitFormFillEnvironment(IntPtr form);
    }

    // /FT is inheritable. A parent field is not necessarily an annotation;
    // let PDFium's form lookup resolve its type when /FT is not on the widget.
    private static bool IsSignatureField(IntPtr document,IntPtr annot)
    {
        if(Api.FPDFAnnot_HasKey(annot,"FT")!=0)
        {
            var value=new byte[8];
            return StampApi.FPDFAnnot_GetStringValue(annot,"FT",value,8)==8 &&
                Encoding.Unicode.GetString(value)=="Sig\0";
        }
        if(Api.FPDFAnnot_HasKey(annot,"Parent")==0) return false;
        // FPDF_FORMFILLINFO v1 on x64: aligned version followed by 16
        // callback/platform pointers. No UI or script actions are installed.
        var info=new IntPtr[17];
        info[0]=new IntPtr(1);
        var pinned=GCHandle.Alloc(info,GCHandleType.Pinned);
        try
        {
            var form=StampApi.FPDFDOC_InitFormFillEnvironment(document,pinned.AddrOfPinnedObject());
            if(form==IntPtr.Zero) return false;
            try { return StampApi.FPDFAnnot_GetFormFieldType(form,annot)==7; }
            finally { StampApi.FPDFDOC_ExitFormFillEnvironment(form); }
        }
        finally { pinned.Free(); }
    }

    private sealed class StampAnnotation
    {
        internal int Index;
        internal Rect Box;
    }

    // Other annotations could cover confidential text or paint a form value.
    // Keep those pages in the existing general rendering path.
    private static List<StampAnnotation> StampAnnotations(IntPtr document,IntPtr page)
    {
        int count=Api.FPDFPage_GetAnnotCount(page);
        if(count<0 || count>32) return null;
        var found=new List<StampAnnotation>();
        for(int i=0;i<count;i++)
        {
            IntPtr annot=Paint.FPDFPage_GetAnnot(page,i);
            if(annot==IntPtr.Zero) return null;
            try
            {
                int subtype=Paint.FPDFAnnot_GetSubtype(annot);
                if(subtype!=13 && (subtype!=20 || !IsSignatureField(document,annot))) return null;
                if((StampApi.FPDFAnnot_GetFlags(annot)&(1|2|32))!=0) continue;
                Rect box;
                if(Api.FPDFAnnot_HasKey(annot,"AP")==0 || Api.FPDFAnnot_GetRect(annot,out box)==0 ||
                   !Finite(box.Left) || !Finite(box.Right) || !Finite(box.Top) || !Finite(box.Bottom) ||
                   box.Right<=box.Left || box.Top<=box.Bottom) return null;
                found.Add(new StampAnnotation{Index=i,Box=box});
            }
            finally { Paint.FPDFPage_CloseAnnot(annot); }
        }
        return found;
    }

    private static bool HasStampAnnotation(IntPtr document,IntPtr page)
    {
        var stamps=StampAnnotations(document,page);
        return stamps!=null && stamps.Count>0;
    }

    private static bool StampPageBox(IntPtr page,out Rect box)
    {
        if(Api.FPDF_GetPageBoundingBox(page,out box)==0 || Api.FPDFPage_GetRotation(page)!=0) return false;
        double width=box.Right-box.Left,height=box.Top-box.Bottom;
        return width>0 && width<=1584 && height>0 && height<=1584;
    }

    private static bool SingleScannedImage(IntPtr page,Rect box)
    {
        bool image=false;
        for(int i=0;i<Api.FPDFPage_CountObjects(page);i++)
        {
            var obj=Api.FPDFPage_GetObject(page,i);
            int active;
            Rect r;
            if(obj==IntPtr.Zero || Api.FPDFPageObj_GetIsActive(obj,out active)==0) return false;
            if(active==0) continue;
            if(Api.FPDFPageObj_GetBounds(obj,out r.Left,out r.Bottom,out r.Right,out r.Top)==0) return false;
            int type=Api.FPDFPageObj_GetType(obj);
            if(type==1 && r.Right-r.Left<.01 && r.Top-r.Bottom<.01) continue;
            if(type!=3 || image) return false;
            double toleranceX=Math.Max(1,(box.Right-box.Left)*.025);
            double toleranceY=Math.Max(1,(box.Top-box.Bottom)*.025);
            if(r.Left>box.Left+toleranceX || r.Right<box.Right-toleranceX ||
               r.Bottom>box.Bottom+toleranceY || r.Top<box.Top-toleranceY) return false;
            image=true;
        }
        return image;
    }

    private static bool IsScanImagePage(IntPtr document,IntPtr page)
    {
        Rect box;
        return StampPageBox(page,out box) && SingleScannedImage(page,box) && StampAnnotations(document,page)!=null;
    }

    // Decline opaque labels pasted over content. Moving those to the front
    // as a separate object would reveal text intentionally covered by a PDF.
    // Measure the ink box (the Appearance's /Rect may include large margins).
    private static bool OpenStampInk(BitmapData data,int width,int height)
    {
        int left=width,right=-1,top=height,bottom=-1,solid=0;
        var row=new byte[checked(width*4)];
        for(int y=0;y<height;y++)
        {
            Marshal.Copy(IntPtr.Add(data.Scan0,y*data.Stride),row,0,row.Length);
            for(int x=0;x<width;x++)
            {
                byte alpha=row[x*4+3];
                if(alpha<=8) continue;
                left=Math.Min(left,x); right=Math.Max(right,x);
                top=Math.Min(top,y); bottom=y;
                if(alpha>=240) solid++;
            }
        }
        return right>=left && solid<.75*(right-left+1)*(bottom-top+1);
    }

    // Render in page coordinates even for a cropped destination. BGRA from
    // PDFium is straight alpha, as required by Format32bppArgb and PNG.
    private static byte[] RenderStampPixels(IntPtr page,int width,int height,Rectangle crop,bool transparent)
    {
        var format=transparent?PixelFormat.Format32bppArgb:PixelFormat.Format32bppRgb;
        using(var bitmap=new Bitmap(crop.Width,crop.Height,format))
        {
            var data=bitmap.LockBits(new Rectangle(0,0,crop.Width,crop.Height),ImageLockMode.ReadWrite,format);
            bool painted=!transparent;
            try
            {
                var native=Paint.FPDFBitmap_CreateEx(crop.Width,crop.Height,transparent?4:3,data.Scan0,data.Stride);
                if(native==IntPtr.Zero) throw new IOException("无法分配盖章页图像。");
                try
                {
                    if(Paint.FPDFBitmap_FillRect(native,0,0,crop.Width,crop.Height,transparent?0u:0xffffffff)==0)
                        throw new IOException("无法初始化盖章页图像。");
                    Paint.FPDF_RenderPageBitmap(native,page,-crop.X,-crop.Y,width,height,0,0);
                    if(transparent) painted=OpenStampInk(data,crop.Width,crop.Height);
                }
                finally { Paint.FPDFBitmap_Destroy(native); }
            }
            finally { bitmap.UnlockBits(data); }
            if(!painted) return null;
            using(var output=new MemoryStream())
            {
                bitmap.Save(output,ImageFormat.Png);
                return output.ToArray();
            }
        }
    }

    // Flatten just this annotation into a scratch page, keeping its real
    // appearance stream, masks and transforms. Widget appearances are not
    // drawn by FPDF_RenderPageBitmap's ordinary annotation flag.
    private static byte[] StampAppearance(IntPtr source,int pageIndex,int annotation,int width,int height,Rectangle crop)
    {
        var copy=StampApi.FPDF_CreateNewDocument();
        if(copy==IntPtr.Zero) return null;
        IntPtr page=IntPtr.Zero;
        try
        {
            if(StampApi.FPDF_ImportPagesByIndex(copy,source,ref pageIndex,1,0)==0) return null;
            page=Api.FPDF_LoadPage(copy,0);
            if(page==IntPtr.Zero || annotation>=Api.FPDFPage_GetAnnotCount(page)) return null;
            for(int i=Api.FPDFPage_GetAnnotCount(page)-1;i>=0;i--)
                if(i!=annotation && StampApi.FPDFPage_RemoveAnnot(page,i)==0) return null;
            for(int i=Api.FPDFPage_CountObjects(page)-1;i>=0;i--)
            {
                var obj=Api.FPDFPage_GetObject(page,i);
                if(obj==IntPtr.Zero || StampApi.FPDFPage_RemoveObject(page,obj)==0) return null;
                StampApi.FPDFPageObj_Destroy(obj);
            }
            if(StampApi.FPDFPage_GenerateContent(page)==0 || StampApi.FPDFPage_Flatten(page,0)!=1) return null;
            Api.FPDF_ClosePage(page);
            page=Api.FPDF_LoadPage(copy,0);
            return page==IntPtr.Zero?null:RenderStampPixels(page,width,height,crop,true);
        }
        finally
        {
            if(page!=IntPtr.Zero) Api.FPDF_ClosePage(page);
            Api.FPDF_CloseDocument(copy);
        }
    }

    private static LayoutPage ReadStampPage(IntPtr document,int pageIndex,IntPtr page,Stream basePng)
    {
        Rect box;
        if(!StampPageBox(page,out box)) return null;
        var annotations=StampAnnotations(document,page);
        if(annotations==null) return null;
        bool scanned=SingleScannedImage(page,box);
        if(!scanned && annotations.Count==0) return null;
        var layout=new LayoutPage{Width=box.Right-box.Left,Height=box.Top-box.Bottom};
        double scale=Math.Min(4.167,Math.Sqrt(16000000/(layout.Width*layout.Height)));
        int width=Math.Max(1,(int)Math.Round(scale*layout.Width)),height=Math.Max(1,(int)Math.Round(scale*layout.Height));
        double sx=width/layout.Width,sy=height/layout.Height;
        foreach(var annotation in annotations)
        {
            var r=annotation.Box;
            // Clip to the visible page before converting PDF coordinates to ints.
            double l=Math.Max(box.Left,r.Left),b=Math.Max(box.Bottom,r.Bottom);
            double t=Math.Min(box.Top,r.Top),rt=Math.Min(box.Right,r.Right);
            if(rt<=l || t<=b) return null;
            var crop=Rectangle.FromLTRB(Math.Max(0,(int)Math.Floor((l-box.Left)*sx)-1),
                Math.Max(0,(int)Math.Floor((box.Top-t)*sy)-1),Math.Min(width,(int)Math.Ceiling((rt-box.Left)*sx)+1),
                Math.Min(height,(int)Math.Ceiling((box.Top-b)*sy)+1));
            var png=StampAppearance(document,pageIndex,annotation.Index,width,height,crop);
            if(png==null) return null;
            layout.Stamps.Add(new LayoutStamp{Left=crop.Left/sx,Top=crop.Top/sy,Width=crop.Width/sx,Height=crop.Height/sy,Png=png});
        }
        if(!scanned)
        {
            var objects=VisibleText(page,box);
            if(objects==null) return null;
            if(objects.Count>0)
            {
                var text=Api.FPDFText_LoadPage(page);
                if(text==IntPtr.Zero) return null;
                try { if(!ReadCharacters(text,box,layout,objects,true)) return null; }
                finally { Api.FPDFText_ClosePage(text); }
                foreach(var obj in objects) if(Paint.FPDFPageObj_SetIsActive(obj,0)==0) return null;
            }
        }
        byte[] image=RenderStampPixels(page,width,height,new Rectangle(0,0,width,height),false);
        basePng.Write(image,0,image.Length);
        basePng.Position=0;
        return layout;
    }
}
