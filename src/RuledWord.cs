using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

// Rebuild the ruled form as ordinary Word tables and paragraphs. Page and cell
// positions are measured from the PDF; Word never has to infer a merged boundary.
internal sealed class RuledWord
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private int runId, lineId;

    private static int Tw(double pt) { return (int)Math.Round(pt*20); }
    private static string Pt(double pt) { return pt.ToString("0.###",CultureInfo.InvariantCulture)+"pt"; }
    private static XAttribute At(string name,object value) { return new XAttribute(W+name,value); }
    private static XElement Val(string name,object value) { return new XElement(W+name,At("val",value)); }
    private static XElement Size(string name, int twips) { return new XElement(W+name,At("w",twips),At("type","dxa")); }
    private static string Gray(int gray)
    {
        string g = Math.Max(0,Math.Min(gray,255)).ToString("X2",CultureInfo.InvariantCulture);
        return g+g+g;
    }

    private sealed class TextRow
    {
        internal List<LayoutLine> Runs = new List<LayoutLine>();
        internal double Baseline { get { return Runs.Max(run=>Base(run)); } }
        internal double FontSize { get { return Runs.Max(x=>x.FontSize); } }
        internal double Height { get { return FontSize*1.15; } }
        internal double Top { get { return Baseline - Height + FontSize*.16; } }
    }

    private static double Base(LayoutLine run)
    {
        return run.Baseline>0 ? run.Baseline : run.Bottom-run.FontSize*.08;
    }

    private static List<TextRow> Rows(IEnumerable<LayoutLine> runs)
    {
        var result = new List<TextRow>();
        TextRow last = null;
        foreach (var run in runs.OrderBy(Base))
        {
            if(last==null || Base(run)-Base(last.Runs[0])>Math.Max(.7,Math.Min(run.FontSize,last.Runs[0].FontSize)*.18))
            {
                last = new TextRow();
                result.Add(last);
            }
            last.Runs.Add(run);
        }
        foreach (var row in result) row.Runs.Sort((a,b)=>a.Left.CompareTo(b.Left));
        return result;
    }

    private static XElement Spacing(double height)
    {
        return new XElement(W+"spacing",At("before",0),At("after",0),At("line",Math.Max(1,Tw(height))),At("lineRule","exact"));
    }

    private static XElement Spacer(double height)
    {
        return new XElement(W+"p",new XElement(W+"pPr",Spacing(height),new XElement(W+"rPr",Val("sz",2))));
    }

    private XElement Run(LayoutLine line)
    {
        string font=String.IsNullOrWhiteSpace(line.FontName)?"SimSun":line.FontName;
        int size=Math.Max(2,(int)Math.Round(2*line.FontSize));
        var style=new XElement(W+"rPr",
            new XElement(W+"rFonts",At("ascii",font),At("hAnsi",font),At("eastAsia",font),At("cs",font)),
            line.Bold?Val("b",1):null, Val("sz",size),Val("szCs",size));
        // Match the PDF's advance width, including scaled labels and the white
        // space carried by full-width punctuation.
        if(line.Text.Trim().Length>=3 && line.Right>line.Left+1)
            style.Add(new XElement(W+"fitText",At("val",Math.Max(1,Tw(line.Right-line.Left))),At("id",++runId)));
        return new XElement(W+"r",style,new XElement(W+"t",new XAttribute(XNamespace.Xml+"space","preserve"),line.Text));
    }

    private XElement Paragraph(TextRow row,double left,double height)
    {
        var first=row.Runs[0];
        var tabs=new XElement(W+"tabs");
        var pr=new XElement(W+"pPr", Val("widowControl",0),Val("autoSpaceDE",0),Val("autoSpaceDN",0),
            Val("snapToGrid",0), Spacing(height),
            new XElement(W+"ind",At("left",Math.Max(0,Tw(first.Left-left)))),
            new XElement(W+"rPr",Val("sz",Math.Max(2,(int)Math.Round(2*row.FontSize)))));
        var p=new XElement(W+"p",pr);
        double prevRight=first.Left;
        foreach (var run in row.Runs)
        {
            if(run!=first && run.Left-prevRight>Math.Max(1.4,run.FontSize*.2))
            {
                tabs.Add(new XElement(W+"tab",At("val","left"),At("pos",Math.Max(0,Tw(run.Left-left)))));
                p.Add(new XElement(W+"r",new XElement(W+"tab")));
            }
            p.Add(Run(run));
            prevRight=run.Right;
        }
        if(tabs.HasElements) pr.Element(W+"autoSpaceDE").AddBeforeSelf(tabs);
        return p;
    }

    private XElement Cell(RuledTable table,RuledCell cell,int row,List<LayoutLine> labels)
    {
        double left=table.X[cell.Col], right=table.X[cell.Col+cell.ColSpan];
        double top=table.Y[cell.Row], bottom=table.Y[cell.Row+cell.RowSpan];
        var props=new XElement(W+"tcPr",Size("tcW",Tw(right)-Tw(left)));
        if(cell.ColSpan>1) props.Add(Val("gridSpan",cell.ColSpan));
        if(cell.RowSpan>1) props.Add(Val("vMerge",cell.Row==row?"restart":"continue"));
        var tc=new XElement(W+"tc",props);
        if(row!=cell.Row || labels.Count==0)
        {
            // Preserve usable paragraph formatting for user-entered form data.
            props.Add(Val("vAlign","center"));
            tc.Add(new XElement(W+"p",new XElement(W+"pPr",new XElement(W+"spacing",At("before",0),At("after",0)),
                new XElement(W+"ind",At("left",20)),new XElement(W+"rPr",Val("sz",20)))));
            return tc;
        }
        var lines=Rows(labels);
        double firstTop=lines[0].Top;
        double lastBottom=lines.Last().Top+lines.Last().Height;
        bool centered=Math.Abs((firstTop-top)-(bottom-lastBottom))<=2.0;
        if(!centered)
            props.Add(new XElement(W+"tcMar",Size("top",Math.Max(0,Tw(firstTop-top)))));
        props.Add(Val("vAlign",centered?"center":"top"));
        for(int i=0;i<lines.Count;i++)
        {
            var line=lines[i];
            if(i>0)
            {
                double gap=line.Top-(lines[i-1].Top+lines[i-1].Height);
                if(gap>.05) tc.Add(Spacer(gap));
            }
            double height = Math.Min(line.Height,Math.Max(1,(i+1<lines.Count?lines[i+1].Top:bottom)-line.Top));
            tc.Add(Paragraph(line,left,height));
        }
        return tc;
    }

    private XElement Table(RuledTable table,HashSet<LayoutLine> remaining)
    {
        // Resolve even an interior slot of a merge to its one editable owner.
        var slots=new RuledCell[table.Y.Length-1,table.X.Length-1];
        foreach(var cell in table.Cells)
            for(int r=cell.Row;r<cell.Row+cell.RowSpan;r++)
                for(int c=cell.Col;c<cell.Col+cell.ColSpan;c++) slots[r,c]=cell;
        var labels=table.Cells.ToDictionary(cell=>cell,cell=>new List<LayoutLine>());
        foreach(var line in remaining.ToList())
        {
            double x=(line.Left+line.Right)/2, y=(line.Top+line.Bottom)/2;
            if(x<table.X[0] || x>=table.X.Last() || y<table.Y[0] || y>=table.Y.Last()) continue;
            int col=Array.BinarySearch(table.X,x), row=Array.BinarySearch(table.Y,y);
            var owner=slots[row<0?~row-1:row,col<0?~col-1:col];
            labels[owner].Add(line);
            remaining.Remove(line);
        }
        var borders=new XElement(W+"tblBorders");
        foreach(string side in new[]{"top","left","bottom","right","insideH","insideV"})
            borders.Add(new XElement(W+side,At("val","single"),At("sz",Math.Max(2,(int)Math.Round(table.BorderWidth*8))),At("color",Gray(table.BorderGray))));
        var margins=new XElement(W+"tblCellMar");
        foreach(string side in new[]{"top","left","bottom","right"}) margins.Add(Size(side,0));
        var props=new XElement(W+"tblPr",Size("tblW",Tw(table.X.Last())-Tw(table.X[0])),
            Size("tblInd",Tw(table.X[0])),borders,new XElement(W+"tblLayout",At("type","fixed")),margins);
        var grid=new XElement(W+"tblGrid");
        for(int c=1;c<table.X.Length;c++)
            grid.Add(new XElement(W+"gridCol",At("w",Tw(table.X[c])-Tw(table.X[c-1]))));
        var tbl=new XElement(W+"tbl",props,grid);
        for(int r=0;r<table.Y.Length-1;r++)
        {
            var tr=new XElement(W+"tr",new XElement(W+"trPr",new XElement(W+"cantSplit"),
                new XElement(W+"trHeight",At("val",Tw(table.Y[r+1])-Tw(table.Y[r])),At("hRule","exact"))));
            for(int c=0;c<slots.GetLength(1);)
            {
                var cell=slots[r,c];
                tr.Add(Cell(table,cell,r,labels[cell]));
                c+=cell.ColSpan;
            }
            tbl.Add(tr);
        }
        return tbl;
    }

    private XElement LineShape(LayoutRule rule)
    {
        // In VML the line endpoints determine the extent; the page-relative
        // margins set its origin. This remains an editable Word line.
        string position="position:absolute;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;"
            +"margin-left:"+Pt(rule.Left)+";margin-top:"+Pt(rule.Top)+";z-index:1";
        return new XElement(W+"r",new XElement(W+"pict",new XElement(V+"line",
            new XAttribute("id","PDFRule"+(++lineId)),new XAttribute("style",position),new XAttribute("from","0,0"),
            new XAttribute("to",Pt(rule.Right-rule.Left)+","+Pt(rule.Bottom-rule.Top)),
            new XAttribute("strokecolor","#"+Gray(rule.Gray)),new XAttribute("strokeweight",Pt(rule.Width)))));
    }

    private void Page(XElement body,LayoutPage page,bool last)
    {
        var remaining=new HashSet<LayoutLine>(page.Lines);
        var blocks=new List<Tuple<double,double,XElement>>();
        foreach(var table in page.Tables)
            blocks.Add(Tuple.Create(table.Y[0],table.Y.Last(),Table(table,remaining)));
        foreach(var row in Rows(remaining))
            blocks.Add(Tuple.Create(row.Top,row.Top+row.Height,Paragraph(row,0,row.Height)));
        double cursor=0;
        foreach(var block in blocks.OrderBy(b=>b.Item1))
        {
            if(block.Item1>cursor+.05) body.Add(Spacer(block.Item1-cursor));
            body.Add(block.Item3);
            cursor=Math.Max(cursor,block.Item1)+(block.Item2-block.Item1);
        }
        var end=Spacer(.05);
        foreach(var rule in page.Rules) end.Add(LineShape(rule));
        body.Add(end);
        var size=new XElement(W+"pgSz",At("w",Tw(page.Width)),At("h",Tw(page.Height)));
        if(page.Width>page.Height) size.SetAttributeValue(W+"orient","landscape");
        var margin=new XElement(W+"pgMar");
        foreach(string side in new[]{"top","right","bottom","left","header","footer","gutter"})
            margin.SetAttributeValue(W+side,0);
        var section=new XElement(W+"sectPr",Val("type","nextPage"),size,margin);
        if(last) body.Add(section);
        else end.Element(W+"pPr").Add(section);
    }

    private static void Part(ZipArchive archive,string path,XElement xml)
    {
        using(var stream=archive.CreateEntry(path).Open()) xml.Save(stream,SaveOptions.DisableFormatting);
    }

    internal static void Write(string output,IList<LayoutPage> pages)
    {
        var body=new XElement(W+"body");
        var writer=new RuledWord();
        for(int i=0;i<pages.Count;i++) writer.Page(body,pages[i],i+1==pages.Count);
        using(var zip=ZipFile.Open(output,ZipArchiveMode.Create))
        {
            Part(zip,"word/document.xml",new XElement(W+"document",new XAttribute(XNamespace.Xmlns+"w",W),
                new XAttribute(XNamespace.Xmlns+"v",V),body));
            Part(zip,"word/settings.xml",new XElement(W+"settings",new XElement(W+"compat",
                new XElement(W+"compatSetting",At("name","compatibilityMode"),At("uri","http://schemas.microsoft.com/office/word"),At("val",15)))));
            Part(zip,"word/_rels/document.xml.rels",new XElement(Rel+"Relationships",new XElement(Rel+"Relationship",
                new XAttribute("Id","settings"),new XAttribute("Type",R.NamespaceName+"/settings"),new XAttribute("Target","settings.xml"))));
            Part(zip,"_rels/.rels",new XElement(Rel+"Relationships",new XElement(Rel+"Relationship",
                new XAttribute("Id","main"),new XAttribute("Type",R.NamespaceName+"/officeDocument"),new XAttribute("Target","word/document.xml"))));
            Part(zip,"[Content_Types].xml",new XElement(Ct+"Types",
                new XElement(Ct+"Default",new XAttribute("Extension","rels"),new XAttribute("ContentType","application/vnd.openxmlformats-package.relationships+xml")),
                new XElement(Ct+"Override",new XAttribute("PartName","/word/document.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")),
                new XElement(Ct+"Override",new XAttribute("PartName","/word/settings.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"))));
        }
    }
}
