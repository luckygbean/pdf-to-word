using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Storage.Streams;

internal static class FixedPages
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";

    private static XElement Section(int width, int height)
    {
        var size = new XElement(W+"pgSz", new XAttribute(W+"w", width), new XAttribute(W+"h", height));
        if (width > height) size.SetAttributeValue(W+"orient", "landscape");
        var margins = new XElement(W+"pgMar");
        foreach (var name in new[] {"top","left","right","bottom","header","footer","gutter"})
            margins.SetAttributeValue(W+name, 0);
        return new XElement(W+"sectPr", new XElement(W+"type", new XAttribute(W+"val","nextPage")), size, margins);
    }

    internal static XElement Drawing(int n, long cx, long cy, bool behind=false)
    {
        var name = "PDF 第 " + n + " 页";
        var picture = new XElement(Pic+"pic",
            new XElement(Pic+"nvPicPr", new XElement(Pic+"cNvPr", new XAttribute("id",n), new XAttribute("name",name)), new XElement(Pic+"cNvPicPr")),
            new XElement(Pic+"blipFill", new XElement(A+"blip", new XAttribute(R+"embed","page"+n)), new XElement(A+"stretch",new XElement(A+"fillRect"))),
            new XElement(Pic+"spPr",
                new XElement(A+"xfrm", new XElement(A+"off",new XAttribute("x",0),new XAttribute("y",0)),new XElement(A+"ext",new XAttribute("cx",cx),new XAttribute("cy",cy))),
                new XElement(A+"prstGeom",new XAttribute("prst","rect"),new XElement(A+"avLst"))));
        var anchor = new XElement(Wp+"anchor",new XAttribute("simplePos",0),new XAttribute("relativeHeight",0),new XAttribute("behindDoc",behind?1:0),
            new XAttribute("locked",1),new XAttribute("layoutInCell",0),new XAttribute("allowOverlap",1),
            new XElement(Wp+"simplePos",new XAttribute("x",0),new XAttribute("y",0)),
            new XElement(Wp+"positionH",new XAttribute("relativeFrom","page"),new XElement(Wp+"posOffset",0)),
            new XElement(Wp+"positionV",new XAttribute("relativeFrom","page"),new XElement(Wp+"posOffset",0)),
            new XElement(Wp+"extent",new XAttribute("cx",cx),new XAttribute("cy",cy)),
            new XElement(Wp+"wrapNone"),new XElement(Wp+"docPr",new XAttribute("id",n),new XAttribute("name",name)),
            new XElement(Wp+"cNvGraphicFramePr",new XElement(A+"graphicFrameLocks",new XAttribute("noChangeAspect",1))),
            new XElement(A+"graphic", new XElement(A+"graphicData",new XAttribute("uri", Pic.NamespaceName),picture)));
        return new XElement(W+"r",new XElement(W+"drawing",anchor));
    }

    private static XElement Link(string id, string kind, string target)
    {
        return new XElement(Rel+"Relationship", new XAttribute("Id",id),new XAttribute("Type",R.NamespaceName+"/"+kind),new XAttribute("Target",target));
    }

    private static void Xml(ZipArchive zip, string path, XElement root)
    {
        using (var stream = zip.CreateEntry(path).Open()) root.Save(stream, SaveOptions.DisableFormatting);
    }

    public static async Task<string> ConvertAsync(string input, string output, string password, int dpi, Action<int,int> progress)
    {
        var file = await StorageFile.GetFileFromPathAsync(input);
        using (var stream = await file.OpenAsync(FileAccessMode.Read))
        {
            var pdf = await PdfDocument.LoadFromStreamAsync(stream, password ?? "");
            return await WritePagesAsync(pdf, output, dpi, progress);
        }
    }

    internal static async Task<bool> RenderAsync(PdfPage page,InMemoryRandomAccessStream pixels,int dpi)
    {
        var size=page.Size;
        double scale = Math.Min(dpi/96.0, Math.Sqrt(30000000.0/(size.Width*size.Height)));

        var options = new PdfPageRenderOptions { DestinationWidth = (uint)Math.Max(1,Math.Round(size.Width*scale)),
            DestinationHeight = (uint)Math.Max(1,Math.Round(size.Height*scale)), BitmapEncoderId = BitmapEncoder.PngEncoderId,
            IsIgnoringHighContrast = true, BackgroundColor = Windows.UI.Colors.White };
        await page.RenderToStreamAsync(pixels,options);
        pixels.Seek(0);
        return scale+0.001 < dpi/96.0;
    }

    private static async Task<string> WritePagesAsync(PdfDocument pdf, string output, int dpi, Action<int,int> progress)
    {
        int count = checked((int)pdf.PageCount);
        if (count == 0) throw new InvalidDataException("PDF 没有页面。");
        var body = new XElement(W+"body");
        var rels = new XElement(Rel+"Relationships",Link("settings","settings","settings.xml"));
        var reduced = new List<int>();
        using (var zip = ZipFile.Open(output, ZipArchiveMode.Create))
        {
            for (int i = 0; i < count; i++)
            using (var page = pdf.GetPage((uint)i))
            using (var pixels = new InMemoryRandomAccessStream())
            {
                var size = page.Size;
                int width = (int)Math.Round(size.Width*15), height = (int)Math.Round(size.Height*15);
                if (Math.Min(width,height)<1 || Math.Max(width,height)>31680)
                    throw new InvalidDataException("第 "+(i+1)+" 页尺寸超出 Word 支持范围（最长边 55.88 厘米）。请先裁切或缩小该页。");
                if(await RenderAsync(page,pixels,dpi)) reduced.Add(i+1);
                string media = "media/page"+(i+1)+".png";
                using (var dest = zip.CreateEntry("word/"+media, CompressionLevel.NoCompression).Open())
                using (var source = pixels.AsStreamForRead()) await source.CopyToAsync(dest);
                rels.Add(Link("page"+(i+1),"image",media));
                var props = new XElement(W+"pPr",new XElement(W+"spacing",new XAttribute(W+"before",0),new XAttribute(W+"after",0),new XAttribute(W+"line",20),new XAttribute(W+"lineRule","exact")));
                var section = Section(width,height);
                if (i+1 < count) props.Add(section);
                body.Add(new XElement(W+"p",props,Drawing(i+1,width*635L,height*635L)));
                if (i+1 == count) body.Add(section);
                progress(i+1,count);
            }
            Xml(zip,"word/document.xml",new XElement(W+"document",new XAttribute(XNamespace.Xmlns+"w",W),new XAttribute(XNamespace.Xmlns+"wp",Wp),
                new XAttribute(XNamespace.Xmlns+"a",A),new XAttribute(XNamespace.Xmlns+"pic",Pic),new XAttribute(XNamespace.Xmlns+"r",R),body));
            Xml(zip,"word/settings.xml",new XElement(W+"settings",
                new XElement(W+"compat",new XElement(W+"compatSetting",new XAttribute(W+"name","compatibilityMode"),new XAttribute(W+"uri","http://schemas.microsoft.com/office/word"),new XAttribute(W+"val",15))),
                new XElement(W+"doNotAutoCompressPictures")));
            Xml(zip,"word/_rels/document.xml.rels",rels);
            Xml(zip,"_rels/.rels",new XElement(Rel+"Relationships",Link("main","officeDocument","word/document.xml")));
            Xml(zip,"[Content_Types].xml",new XElement(Ct+"Types",
                new XElement(Ct+"Default",new XAttribute("Extension","rels"),new XAttribute("ContentType","application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(Ct+"Default",new XAttribute("Extension","png"),new XAttribute("ContentType","image/png")),
                new XElement(Ct+"Override",new XAttribute("PartName","/word/document.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")),
                new XElement(Ct+"Override",new XAttribute("PartName","/word/settings.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"))));
        }
        return reduced.Count==0 ? "" : "以下页受单页 3000 万像素上限影响，已按比例降低分辨率："+String.Join("、",reduced);
    }
}
