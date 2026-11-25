using System;
using System.Collections.Generic;
using System.Linq;

internal static class RuledForms
{
    // A PDF text object can span columns even when individual glyphs are drawn
    // in different cells. Do not assign such an entire object to a single cell.
    private static bool LabelsFit(LayoutPage page, RuledTable table,bool nested)
    {
        foreach(var line in page.Lines)
        {
            double y=(line.Top+line.Bottom)/2, x=(line.Left+line.Right)/2;
            if(y<=table.Y[0] || y>=table.Y.Last()) continue;
            if(x<table.X[0] || x>=table.X.Last())
            {
                if(nested && (line.Right<=table.X[0]+.75 || line.Left>=table.X.Last()-.75)) continue;
                return false;
            }
            var cell=table.Cells.FirstOrDefault(c=>x>=table.X[c.Col] && x<table.X[c.Col+c.ColSpan] &&
                y>=table.Y[c.Row] && y<table.Y[c.Row+c.RowSpan]);
            // Full-width punctuation can have an advance past the edge even
            // though all of its ink fits the source cell.
            bool ink=line.InkRight>line.InkLeft;
            if(cell==null || (ink?line.InkLeft:line.Left)<table.X[cell.Col]-.75 ||
               (ink?line.InkRight:line.Right)>table.X[cell.Col+cell.ColSpan]+.75) return false;
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

    private static RuledCell EnclosingCell(RuledTable outer,RuledTable inner)
    {
        const double tolerance=.45;
        return outer.Cells.FirstOrDefault(cell=>inner.X[0]>=outer.X[cell.Col]-tolerance &&
            inner.X.Last()<=outer.X[cell.Col+cell.ColSpan]+tolerance &&
            inner.Y[0]>=outer.Y[cell.Row]-tolerance && inner.Y.Last()<=outer.Y[cell.Row+cell.RowSpan]+tolerance &&
            cell.NestedTables.All(sibling=>sibling.Y.Last()<=inner.Y[0]+tolerance));
    }

    private static LayoutRule DominantRule(IEnumerable<LayoutRule> rules)
    {
        return rules.GroupBy(r=>new { Width=Math.Round(r.Width,2), r.Gray })
            .OrderByDescending(g=>g.Sum(r=>r.Right-r.Left+r.Bottom-r.Top))
            .Select(g=>g.First()).FirstOrDefault();
    }

    private static void SetBorderStyles(RuledTable table,List<LayoutRule> borders)
    {
        var usual=DominantRule(borders);
        table.BorderWidth=usual.Width;
        table.BorderGray=usual.Gray;
        foreach(string side in new[]{"top","bottom","left","right"})
        {
            bool horizontal=side=="top" || side=="bottom";
            var axis=horizontal?table.Y:table.X;
            double edge=(side=="top" || side=="left")?axis[0]:axis.Last();
            var style=DominantRule(borders.Where(r=>r.IsHorizontal==horizontal &&
                Math.Abs((horizontal?r.Top:r.Left)-edge)<=.65));
            if(style!=null) table.OuterBorders.Add(side,style);
        }
    }

    internal static bool PreparePage(LayoutPage page)
    {
        if(page==null || page.Width<=0 || page.Height<=0 || Math.Max(page.Width,page.Height)>1584)
            return false;
        var tables=GridDetection.Find(page.Rules);
        var remaining=new List<LayoutRule>(page.Rules);
        RuledTable previous=null;
        foreach(var table in tables)
        {
            RuledCell owner=null;
            if(previous!=null && table.Y[0]<previous.Y.Last())
            {
                owner=EnclosingCell(previous,table);
                if(owner==null) continue;
            }
            if(!LabelsFit(page,table,owner!=null)) continue;
            var parsed=remaining.Where(r=>IsBorder(r,table)).ToList();
            if(parsed.Count==0) continue;
            SetBorderStyles(table,parsed);
            var borders=new HashSet<LayoutRule>(parsed);
            remaining.RemoveAll(borders.Contains);
            if(owner==null)
            {
                page.Tables.Add(table);
                previous=table;
            }
            else owner.NestedTables.Add(table);
            page.TableCount++;
            page.CellCount+=table.Cells.Count;
        }
        // Rules outside recognized cells, including incomplete table edges, stay visible.
        page.Rules=remaining;
        return true;
    }
}
