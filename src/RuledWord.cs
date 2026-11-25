using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

// Rebuild the ruled form as ordinary Word tables and paragraphs. Page and cell
// positions are measured from the PDF; Word never has to infer a merged boundary.
internal sealed partial class RuledWord : IDisposable
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private int runId, lineId;
    private readonly int pageCount;
    private readonly ZipArchive archive;
    private readonly FileStream documentPart;
    private readonly XmlWriter xml;
    private readonly XElement relationships = new XElement(Rel+"Relationships");
    private int pagesWritten;
    private bool complete;

    // The XML is streamed a page at a time. ZipArchive can have only one entry
    // open for writing; a temporary file allows image parts to stream directly.
    internal RuledWord(string output,int totalPages)
    {
        if(totalPages<1) throw new ArgumentOutOfRangeException("totalPages");
        pageCount=totalPages;
        string spool=Path.Combine(Path.GetDirectoryName(Path.GetFullPath(output)),Path.GetRandomFileName());
        documentPart=new FileStream(spool,FileMode.CreateNew,FileAccess.ReadWrite,FileShare.None,
            65536,FileOptions.SequentialScan|FileOptions.DeleteOnClose);
        try
        {
            archive=ZipFile.Open(output,ZipArchiveMode.Create);
            xml=XmlWriter.Create(documentPart,new XmlWriterSettings{CloseOutput=false});
            xml.WriteStartDocument();
            xml.WriteStartElement("w","document",W.NamespaceName);
            xml.WriteAttributeString("xmlns","v",null,V.NamespaceName);
            xml.WriteAttributeString("xmlns","r",null,R.NamespaceName);
            xml.WriteStartElement("w","body",W.NamespaceName);
            relationships.Add(Relationship("settings","settings","settings.xml"));
            relationships.Add(Relationship("styles","styles","styles.xml"));
        }
        catch { Dispose(); throw; }
    }

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
        RememberFont(font,line.Text,line.Bold);
        int size=Math.Max(2,(int)Math.Round(2*line.FontSize));
        var style=new XElement(W+"rPr",
            new XElement(W+"rFonts",At("ascii",font),At("hAnsi",font),At("eastAsia",font),At("cs",font)),
            line.Bold?Val("b",1):null, Val("sz",size),Val("szCs",size));
        // Match the PDF's advance width, including scaled labels and the white
        // space carried by full-width punctuation.
        if(line.Text.Trim().Length>=3 && line.Right>line.Left+1)
            style.Add(new XElement(W+"fitText",At("val",Math.Max(1,Tw(line.Right-line.Left))),At("id",++runId)));
        var run=new XElement(W+"r",style,new XElement(W+"t",new XAttribute(XNamespace.Xml+"space","preserve"),line.Text));
        return line.LinkPage>0 && line.LinkPage<=pageCount
            ? new XElement(W+"hyperlink",At("anchor","PDFPage"+line.LinkPage),run)
            : run;
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
        if(row!=cell.Row || (labels.Count==0 && cell.NestedTables.Count==0))
        {
            // Preserve usable paragraph formatting for user-entered form data.
            props.Add(Val("vAlign","center"));
            tc.Add(new XElement(W+"p",new XElement(W+"pPr",new XElement(W+"spacing",At("before",0),At("after",0)),
                new XElement(W+"ind",At("left",20)),new XElement(W+"rPr",Val("sz",20)))));
            return tc;
        }
        // Word applies the largest top cell margin in a row to its neighbors
        // too. In a tall vertical merge that can hide adjacent cells entirely.
        // Padding inside this cell's own text flow has no effect on its peers.
        props.Add(Val("vAlign","top"));
        var remaining=new HashSet<LayoutLine>(labels);
        var blocks=new List<Tuple<double,double,XElement>>();
        foreach(var child in cell.NestedTables)
            blocks.Add(Tuple.Create(child.Y[0],child.Y.Last(),Table(child,remaining,left)));
        var lines=Rows(remaining);
        for(int i=0;i<lines.Count;i++)
        {
            var line=lines[i];
            double height = Math.Min(line.Height,Math.Max(1,(i+1<lines.Count?lines[i+1].Top:bottom)-line.Top));
            var paragraph=Paragraph(line,left,height);
            // Word can wrap the final fitted run as a unit when neighboring
            // short punctuation uses a different advance than the PDF. Allow
            // a small end overhang on lines already close to the cell edge.
            if(right-line.Runs.Last().Right<line.FontSize)
                paragraph.Element(W+"pPr").Element(W+"ind").SetAttributeValue(W+"right",-Tw(Math.Min(5,line.FontSize*.5)));
            blocks.Add(Tuple.Create(line.Top,line.Top+height,paragraph));
        }
        double cursor=top;
        foreach(var block in blocks.OrderBy(b=>b.Item1))
        {
            if(block.Item1>cursor+.05) tc.Add(Spacer(block.Item1-cursor));
            tc.Add(block.Item3);
            cursor=Math.Max(cursor,block.Item1)+(block.Item2-block.Item1);
        }
        // Word requires a trailing paragraph after a table inside a cell.
        if(tc.Elements().Last().Name==W+"tbl") tc.Add(Spacer(.05));
        return tc;
    }

    private XElement Table(RuledTable table,HashSet<LayoutLine> remaining,double parentLeft=0)
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
        {
            LayoutRule edge;
            bool hasEdge=table.OuterBorders.TryGetValue(side,out edge);
            double width=hasEdge?edge.Width:table.BorderWidth;
            int gray=hasEdge?edge.Gray:table.BorderGray;
            borders.Add(new XElement(W+side,At("val","single"),
                At("sz",Math.Max(2,(int)Math.Round(width*8))),At("color",Gray(gray))));
        }
        var margins=new XElement(W+"tblCellMar");
        foreach(string side in new[]{"top","left","bottom","right"}) margins.Add(Size(side,0));
        var props=new XElement(W+"tblPr",Size("tblW",Tw(table.X.Last())-Tw(table.X[0])),
            Size("tblInd",Tw(table.X[0]-parentLeft)),borders,new XElement(W+"tblLayout",At("type","fixed")),margins);
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

    private void Page(XElement body,LayoutPage page,bool last,XElement picture)
    {
        var remaining=new HashSet<LayoutLine>(page.Lines);
        var blocks=new List<Tuple<double,double,XElement>>();
        foreach(var table in page.Tables)
            blocks.Add(Tuple.Create(table.Y[0],table.Y.Last(),Table(table,remaining)));
        var rows=Rows(remaining);
        foreach(var row in rows)
            blocks.Add(Tuple.Create(row.Top,row.Top+row.Height,Paragraph(row,0,row.Height)));
        blocks.Sort((a,b)=>a.Item1.CompareTo(b.Item1));
        // Labels in neighboring columns, superscripts and diagram text can have
        // overlapping line boxes. Sequential paragraphs would push all later
        // lines down, sometimes creating an extra physical page. Page-relative
        // Word text frames preserve editable text at these independent positions.
        double predicted=0, drift=0;
        foreach(var block in blocks)
        {
            drift=Math.Max(drift,predicted-block.Item1);
            predicted=Math.Max(predicted,block.Item1)+block.Item2-block.Item1;
        }
        bool anchorText=drift>2 || predicted>page.Height-2;
        double cursor=0;
        foreach(var block in blocks)
        {
            if(anchorText && block.Item3.Name==W+"p")
            {
                block.Item3.Element(W+"pPr").AddFirst(new XElement(W+"framePr",
                    At("w",Tw(page.Width)),At("wrap","none"),At("hAnchor","page"),At("vAnchor","page"),
                    At("x",0),At("y",Math.Max(0,Tw(block.Item1))),At("hSpace",0),At("vSpace",0)));
                body.Add(block.Item3);
                continue;
            }
            if(block.Item1>cursor+.05) body.Add(Spacer(block.Item1-cursor));
            body.Add(block.Item3);
            cursor=Math.Max(cursor,block.Item1)+(block.Item2-block.Item1);
        }
        var end=Spacer(.05);
        if(picture!=null) end.Add(picture);
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

    private static XElement Relationship(string id,string type,string target)
    {
        return new XElement(Rel+"Relationship",new XAttribute("Id",id),new XAttribute("Type",R.NamespaceName+"/"+type),
            new XAttribute("Target",target));
    }

    // Empty cells and paragraph marks have no text run from which Word can
    // infer a font. Define their style in the package, including the formatting
    // for text typed later, independent of the user's Normal template.
    private static XElement Styles()
    {
        var runDefaults=new XElement(W+"rPr",
            new XElement(W+"rFonts",At("ascii","Times New Roman"),At("hAnsi","Times New Roman"),
                At("eastAsia","SimSun"),At("cs","Times New Roman")),
            Val("sz",20),Val("szCs",20),
            new XElement(W+"lang",At("val","en-US"),At("eastAsia","zh-CN")));
        var paragraphDefaults=new XElement(W+"pPr",
            Val("widowControl",0),Val("autoSpaceDE",0),Val("autoSpaceDN",0),Val("snapToGrid",0),
            new XElement(W+"spacing",At("before",0),At("after",0),At("line",240),At("lineRule","auto")));
        return new XElement(W+"styles",
            new XElement(W+"docDefaults",new XElement(W+"rPrDefault",runDefaults),
                new XElement(W+"pPrDefault",paragraphDefaults)),
            new XElement(W+"style",At("type","paragraph"),At("default",1),At("styleId","Normal"),Val("name","Normal")));
    }

    internal void AddNative(LayoutPage page,Stream cleanBackgroundPng=null)
    {
        AddPage(page,cleanBackgroundPng,true);
    }

    internal void AddImage(double widthPt,double heightPt,Stream png)
    {
        if(png==null) throw new ArgumentNullException("png");
        AddPage(new LayoutPage{Width=widthPt,Height=heightPt},png,false);
    }

    private void AddPage(LayoutPage page,Stream png,bool behind)
    {
        if(complete || pagesWritten>=pageCount) throw new InvalidOperationException("Too many pages.");
        int n=pagesWritten+1;
        XElement picture=null;
        if(png!=null)
        {
            string target="media/page"+n+".png";
            using(var destination=archive.CreateEntry("word/"+target,CompressionLevel.NoCompression).Open())
                png.CopyTo(destination);
            relationships.Add(Relationship("page"+n,"image",target));
            picture=FixedPages.Drawing(n,Tw(page.Width)*635L,Tw(page.Height)*635L,behind);
        }
        var body=new XElement(W+"body");
        Page(body,page,n==pageCount,picture);
        var first=body.Descendants(W+"p").First();
        first.Element(W+"pPr").AddAfterSelf(new XElement(W+"bookmarkStart",At("id",n),At("name","PDFPage"+n)),
            new XElement(W+"bookmarkEnd",At("id",n)));
        foreach(var element in body.Elements()) element.WriteTo(xml);
        pagesWritten++;
    }

    internal void Complete()
    {
        if(complete || pagesWritten!=pageCount) throw new InvalidOperationException("Incomplete document.");
        xml.WriteEndElement();
        xml.WriteEndElement();
        xml.WriteEndDocument();
        xml.Flush();
        documentPart.Position=0;
        using(var entry=archive.CreateEntry("word/document.xml").Open()) documentPart.CopyTo(entry);
        Part(archive,"word/styles.xml",Styles());
        bool embedded=WriteFontTable();
        Part(archive,"word/settings.xml",new XElement(W+"settings",
            Val("view","print"),
            embedded?new XElement(W+"embedTrueTypeFonts"):null,embedded?new XElement(W+"saveSubsetFonts"):null,new XElement(W+"compat",
            new XElement(W+"compatSetting",At("name","compatibilityMode"),At("uri","http://schemas.microsoft.com/office/word"),At("val",15))),
            new XElement(W+"doNotAutoCompressPictures")));
        Part(archive,"word/_rels/document.xml.rels",relationships);
        Part(archive,"_rels/.rels",new XElement(Rel+"Relationships",Relationship("main","officeDocument","word/document.xml")));
        Part(archive,"[Content_Types].xml",new XElement(Ct+"Types",
            embedded?new XElement(Ct+"Default",new XAttribute("Extension","odttf"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.obfuscatedFont")):null,
            new XElement(Ct+"Override",new XAttribute("PartName","/word/fontTable.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml")),
            new XElement(Ct+"Default",new XAttribute("Extension","rels"),new XAttribute("ContentType","application/vnd.openxmlformats-package.relationships+xml")),
            new XElement(Ct+"Default",new XAttribute("Extension","png"),new XAttribute("ContentType","image/png")),
            new XElement(Ct+"Override",new XAttribute("PartName","/word/document.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")),
            new XElement(Ct+"Override",new XAttribute("PartName","/word/styles.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")),
            new XElement(Ct+"Override",new XAttribute("PartName","/word/settings.xml"),new XAttribute("ContentType","application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"))));
        complete=true;
    }

    public void Dispose()
    {
        try { if(xml!=null) xml.Dispose(); }
        finally
        {
            try { if(archive!=null) archive.Dispose(); }
            finally { documentPart.Dispose(); }
        }
    }

}
