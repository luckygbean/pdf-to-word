using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Storage;
using Windows.Storage.Streams;

internal static class MixedDocument
{
    // The general Word importer remains available for ordinary PDFs.
    private static bool CanReconstruct(PdfTextReader.Session reader,bool formOnly)
    {
        for(int i=0;i<reader.PageCount;i++)
        {
            var page=reader.ReadPage(i);
            if(page!=null && page.Rules.Count>0 && RuledForms.PreparePage(page) && page.Tables.Count>0) return true;
            if(formOnly) continue;
            if(reader.HasStampAnnotation(i)) return true;
            if(!reader.IsScanImagePage(i)) continue;
            using(var png=new MemoryStream())
            {
                var scan=reader.ReadStampPage(i,png);
                if(scan!=null && new ScannedSeals(png,scan).Locate().Count>0) return true;
            }
        }
        return false;
    }

    private static async Task WriteImageAsync(PdfDocument pdf,int index,RuledWord writer,int dpi,LayoutPage form)
    {
        using(var source=pdf.GetPage((uint)index))
        using(var pixels=new InMemoryRandomAccessStream())
        {
            double width=source.Size.Width*.75, height=source.Size.Height*.75;
            if(Math.Min(width,height)<.05 || Math.Max(width,height)>1584)
                throw new InvalidDataException("第 "+(index+1)+" 页尺寸超出 Word 支持范围。");
            await FixedPages.RenderAsync(source,pixels,dpi);
            using(var data=pixels.AsStreamForRead())
            {
                if(form==null) writer.AddImage(width,height,data);
                else
                {
                    // The background already paints the labels and rules. Only
                    // the transparent native tables participate in text entry.
                    form.Lines.Clear();
                    form.Rules.Clear();
                    writer.AddNative(form,data);
                }
            }
        }
    }

    private static bool WriteLayers(PdfTextReader.Session reader,int index,RuledWord writer)
    {
        using(var background=new MemoryStream())
        {
            var page=reader.WithPage(index,p=>PdfTextReader.ReadLayers(p,background));
            if(page==null) return false;
            writer.AddNative(page,background);
            return true;
        }
    }

    private static string[] NearbyText(PdfTextReader.Session reader,int index)
    {
        var lines=new List<string>();
        for(int i=Math.Max(0,index-5);i<Math.Min(reader.PageCount,index+6);i++)
        {
            if(i==index) continue;
            var page=reader.ReadPage(i);
            if(page!=null) lines.AddRange(page.Lines.Select(l=>l.Text));
        }
        return lines.ToArray();
    }

    private static async Task<LayoutPage> WriteStampsAsync(PdfTextReader.Session reader,int index,RuledWord writer)
    {
        bool scanned=reader.IsScanImagePage(index);
        if(!scanned && !reader.HasStampAnnotation(index)) return null;
        using(var png=new MemoryStream())
        {
            var page=reader.ReadStampPage(index,png);
            if(page==null) return null;
            if(scanned)
            {
                var scan=new ScannedSeals(png,page);
                var outlines=scan.Locate();
                if(outlines.Count==0 && page.Stamps.Count==0) return null;
                // Ink printed on paper precedes the PDF annotation appearances.
                page.Stamps.InsertRange(0,outlines.Select(scan.Lift).ToList());
                if(await StampedScanText.ReadAsync(scan,page)>0)
                    StampedScanText.CheckIdentifiers(page,()=>NearbyText(reader,index));
                png.SetLength(0);
                ScannedSeals.SavePng(png,scan.Pixels,scan.Width,scan.Height);
            }
            writer.AddNative(page,png);
            return page;
        }
    }

    internal static async Task<string> TryConvertAsync(string input,string output,Action<int,int> progress,bool backgroundForm=false,int dpi=300)
    {
        using(var reader=PdfTextReader.Open(input))
        {
            if(reader==null || !CanReconstruct(reader,backgroundForm)) return null;
            int tables=0, cells=0, layered=0, stamps=0;
            var stampPages=new List<int>();
            var stampPagesWithoutText=new List<int>();
            var images=new List<int>();
            IRandomAccessStream stream=null;
            PdfDocument render=null;
            try
            {
                using(var writer=new RuledWord(output,reader.PageCount,backgroundForm))
                {
                    for(int i=0;i<reader.PageCount;i++)
                    {
                        var page=reader.ReadPage(i);
                        bool native=RuledForms.PreparePage(page);
                        if(native)
                        {
                            tables+=page.TableCount;
                            cells+=page.CellCount;
                        }
                        LayoutPage stamped=null;
                        if(!backgroundForm && !native) stamped=await WriteStampsAsync(reader,i,writer);
                        if(stamped!=null)
                        {
                            stampPages.Add(i+1);
                            stamps+=stamped.Stamps.Count;
                            if(stamped.Lines.Count==0) stampPagesWithoutText.Add(i+1);
                        }
                        else if(!backgroundForm && native) writer.AddNative(page);
                        else if(!backgroundForm && WriteLayers(reader,i,writer)) layered++;
                        else
                        {
                            if(render==null)
                            {
                                var file=await StorageFile.GetFileFromPathAsync(input);
                                stream=await file.OpenAsync(FileAccessMode.Read);
                                render=await PdfDocument.LoadFromStreamAsync(stream);
                            }
                            await WriteImageAsync(render,i,writer,backgroundForm?dpi:300,
                                backgroundForm && native && page.TableCount>0?page:null);
                            if(!backgroundForm) images.Add(i+1);
                        }
                        progress(i+1,reader.PageCount);
                    }
                    writer.Complete();
                }
            }
            finally { if(stream!=null) stream.Dispose(); }
            if(backgroundForm) return String.Format("已生成 {0} 页填写版，保留原稿底图，并叠放 {1} 张可输入文字的 Word 表格。",reader.PageCount,tables);
            string result=String.Format("已按原稿生成 {0} 页，重建 {1} 张可编辑表格、{2} 个独立单元格。",reader.PageCount,tables,cells);
            if(layered>0) result+="其中 "+layered+" 页将可编辑文字与背景图形分别保存。";
            if(stamps>0) result+="第 "+String.Join("、",stampPages)+" 页共分离 "+stamps+" 枚印章，可在 Word 中单独选中。";
            if(stampPagesWithoutText.Count>0) result+="第 "+String.Join("、",stampPagesWithoutText)+" 页未提取到可编辑正文，原文以图像保留。";
            if(images.Count>0) result+="以下页整页保留为图像："+String.Join("、",images)+"。";
            return result;
        }
    }
}
