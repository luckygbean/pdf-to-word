using System;
using System.Collections.Generic;
using System.Linq;

internal static class RuledForms
{
    // A PDF text object can span columns even when individual glyphs are drawn
    // in different cells. Do not assign such an entire object to a single cell.
    private static bool LabelsFit(LayoutPage page, RuledTable table)
    {
        foreach(var line in page.Lines)
        {
            double y=(line.Top+line.Bottom)/2, x=(line.Left+line.Right)/2;
            if(y<=table.Y[0] || y>=table.Y.Last()) continue;
            if(x<table.X[0] || x>=table.X.Last()) return false;
            var cell=table.Cells.FirstOrDefault(c=>x>=table.X[c.Col] && x<table.X[c.Col+c.ColSpan] &&
                y>=table.Y[c.Row] && y<table.Y[c.Row+c.RowSpan]);
            if(cell==null || line.Left<table.X[cell.Col]-.75 ||
               line.Right>table.X[cell.Col+cell.ColSpan]+.75) return false;
        }
        return true;
    }

    // Only consume a PDF rule if its entire length is accounted for by real cell edges.
    // A short underline crossing a merged cell at the same Y as a distant table edge
    // must survive as a separate Word line.
    private static bool IsBorder(LayoutRule rule,RuledTable table)
    {
        bool horizontal=rule.IsHorizontal;
        var along=horizontal?table.X:table.Y;
        var across=horizontal?table.Y:table.X;
        double start=horizontal?rule.Left:rule.Top, end=horizontal?rule.Right:rule.Bottom;
        double at=horizontal?rule.Top:rule.Left;
        const double tolerance=0.65;
        if(start<along[0]-tolerance || end>along[along.Length-1]+tolerance ||
            at<across[0]-tolerance || at>across[across.Length-1]+tolerance) return false;
        var intervals=new List<Tuple<double,double>>();
        foreach(var cell in table.Cells)
        {
            int first=horizontal?cell.Row:cell.Col, span=horizontal?cell.RowSpan:cell.ColSpan;
            if(Math.Abs(at-across[first])>tolerance && Math.Abs(at-across[first+span])>tolerance) continue;
            int begin=horizontal?cell.Col:cell.Row, length=horizontal?cell.ColSpan:cell.RowSpan;
            intervals.Add(Tuple.Create(along[begin],along[begin+length]));
        }
        double covered=start;
        foreach(var interval in intervals.OrderBy(x=>x.Item1))
        {
            if(interval.Item1>covered+tolerance) break;
            covered=Math.Max(covered,interval.Item2);
            if(covered>=end-tolerance) return true;
        }
        return false;
    }

    internal static string TryConvert(string input,string output,Action<int,int> progress)
    {
        var pages=PdfTextReader.Read(input);
        if(pages==null || pages.Count==0) return null;
        int processed=0;
        foreach(var page in pages)
        {
            // Whole-document fallback avoids dropping unsupported pages in mixed PDFs.
            if(page.Width<=0 || page.Height<=0 || Math.Max(page.Width,page.Height)>1584 ||
                page.Lines.Count==0 || page.Rules.Count==0) return null;
            page.Tables=GridDetection.Find(page.Rules);
            if(page.Tables.Count==0) return null;
            var remaining=new List<LayoutRule>(page.Rules);
            double previousTableBottom=-1;
            foreach(var table in page.Tables)
            {
                if(table.Y[0]<previousTableBottom || !LabelsFit(page,table)) return null;
                previousTableBottom=table.Y.Last();
                var parsed=remaining.Where(r=>IsBorder(r,table)).ToList();
                if(parsed.Count==0) return null;
                var style=parsed.GroupBy(r=>new {Width=Math.Round(r.Width,2), r.Gray})
                    .OrderByDescending(group=>group.Sum(r=>r.Right-r.Left+r.Bottom-r.Top)).First().Key;
                table.BorderWidth=style.Width;
                table.BorderGray=style.Gray;
                var borders=new HashSet<LayoutRule>(parsed);
                remaining.RemoveAll(borders.Contains);
            }
            if(remaining.Any(r=>!r.IsHorizontal)) return null;
            page.Rules=remaining;
            progress(++processed,pages.Count);
        }
        RuledWord.Write(output,pages);
        return String.Format("已按 PDF 的表格线和文字位置重建 {0} 页、{1} 张可编辑表格、{2} 个独立单元格。",
            pages.Count,pages.Sum(p=>p.Tables.Count),pages.Sum(p=>p.Tables.Sum(t=>t.Cells.Count)));
    }
}
