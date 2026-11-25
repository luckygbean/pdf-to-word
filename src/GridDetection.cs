using System;
using System.Collections.Generic;
using System.Linq;

internal sealed class RuledTable
{
    internal double[] X, Y;
    internal List<RuledCell> Cells = new List<RuledCell>();
    internal double BorderWidth = 0.5;
    internal int BorderGray = 128;
    internal readonly Dictionary<string,LayoutRule> OuterBorders = new Dictionary<string,LayoutRule>();
}

internal sealed class RuledCell
{
    internal int Row, Col, RowSpan, ColSpan;
    internal readonly List<RuledTable> NestedTables = new List<RuledTable>();
}

internal static class GridDetection
{
    // Tolerate sub-point discrepancies where separate Word/PDF paths meet.
    private const double Tolerance = 0.45;

    private sealed class Stroke
    {
        internal bool Horizontal;
        internal double At, Start, End;
        internal readonly List<int> Crossings = new List<int>();
    }

    private sealed class Sets
    {
        private readonly int[] parent;
        internal Sets(int size) { parent = Enumerable.Range(0, size).ToArray(); }
        internal int Root(int item)
        {
            while (parent[item] != item)
            {
                parent[item] = parent[parent[item]];
                item = parent[item];
            }
            return item;
        }
        internal void Join(int a, int b) { parent[Root(a)] = Root(b); }
    }

    // One PDF writer may stroke an entire edge; another strokes it cell by cell.
    // Work with the same intervals for both encodings. Keep a genuine gap open.
    private static List<Stroke> Normalize(IList<LayoutRule> rules, bool horizontal,double tolerance,double minimumLength)
    {
        var raw = rules.Where(r => r.IsHorizontal == horizontal).Select(r => new Stroke {
            Horizontal = horizontal,
            At = horizontal ? (r.Top + r.Bottom) / 2 : (r.Left + r.Right) / 2,
            Start = horizontal ? r.Left : r.Top,
            End = horizontal ? r.Right : r.Bottom
        }).Where(s => s.End-s.Start > minimumLength).OrderBy(s => s.At).ToList();
        var result = new List<Stroke>();
        for (int i = 0; i < raw.Count; )
        {
            int j = i + 1;
            while (j < raw.Count && raw[j].At - raw[i].At <= tolerance) j++;
            var band = raw.GetRange(i, j-i);
            // Choose an observed axis, weighted by the source stroke length.
            double axis = band.OrderByDescending(s => s.End-s.Start).First().At;
            Stroke current = null;
            foreach (var s in band.OrderBy(s => s.Start))
            {
                if (current == null || s.Start > current.End + tolerance)
                {
                    s.At = axis;
                    current = s;
                    result.Add(s);
                }
                else current.End = Math.Max(current.End, s.End);
            }
            i = j;
        }
        return result;
    }

    internal static List<RuledTable> Find(IList<LayoutRule> rules)
    {
        if (rules==null || rules.Count>16000) return new List<RuledTable>();
        var tables=FindGrids(rules,Tolerance,0);
        // Office emits double horizontal dividers as two long, equal, parallel
        // pieces per column. Their join blocks are wider than a single rule.
        // Use this wider joining distance only for grids containing such a pair.
        var doubles=new List<LayoutRule>();
        foreach(var band in rules.Where(r=>r.IsHorizontal && r.Right-r.Left>=10 && r.Width<=2)
            .GroupBy(r=>new { Start=Math.Round(r.Left), End=Math.Round(r.Right) }))
        {
            LayoutRule previous=null;
            foreach(var rule in band.OrderBy(r=>r.Top))
            {
                if(previous!=null && rule.Top-previous.Top>=.8 && rule.Top-previous.Top<=2 &&
                   Math.Abs(rule.Width-previous.Width)<.1) doubles.Add(rule);
                previous=rule;
            }
        }
        if(doubles.Count==0) return tables;
        const double doubleRuleTolerance=2.4;
        foreach(var candidate in FindGrids(rules,doubleRuleTolerance,2.5))
        {
            double l=candidate.X[0],r=candidate.X.Last(),t=candidate.Y[0],b=candidate.Y.Last();
            if(!doubles.Any(d=>d.Left>=l-doubleRuleTolerance && d.Right<=r+doubleRuleTolerance &&
                candidate.Y.Any(y=>Math.Abs(y-d.Top)<=doubleRuleTolerance))) continue;
            var overlaps=tables.Where(o=>Math.Min(o.X.Last(),r)>Math.Max(o.X[0],l)+Tolerance &&
                Math.Min(o.Y.Last(),b)>Math.Max(o.Y[0],t)+Tolerance).ToList();
            if(overlaps.Any(o=>o.X[0]<l-doubleRuleTolerance || o.X.Last()>r+doubleRuleTolerance ||
                o.Y[0]<t-doubleRuleTolerance || o.Y.Last()>b+doubleRuleTolerance) ||
                overlaps.Sum(o=>o.Cells.Count)>candidate.Cells.Count) continue;
            foreach(var replaced in overlaps) tables.Remove(replaced);
            tables.Add(candidate);
        }
        return tables.OrderBy(table=>table.Y[0]).ThenBy(table=>table.X[0]).ToList();
    }

    private static List<RuledTable> FindGrids(IList<LayoutRule> rules,double tolerance,double minimumLength)
    {
        var tables = new List<RuledTable>();
        var h = Normalize(rules,true,tolerance,minimumLength);
        var v = Normalize(rules,false,tolerance,minimumLength);
        if ((long)h.Count * v.Count > 4000000) return tables;
        var strokes = h.Concat(v).ToList();
        int totalCrossings = 0;
        for (int i = 0; i < h.Count; i++)
        {
            var a = h[i];
            for (int j = 0; j < v.Count; j++)
            {
                var b = v[j];
                if (b.At < a.Start-tolerance || a.End+tolerance < b.At ||
                    a.At < b.Start-tolerance || b.End+tolerance < a.At) continue;
                a.Crossings.Add(h.Count+j);
                b.Crossings.Add(i);
                if (++totalCrossings > 120000) return tables;
            }
        }

        // Repeatedly discard dangling strokes (underlines and other non-cell marks).
        var degrees = strokes.Select(s => s.Crossings.Count).ToArray();
        var leaves = new Queue<int>(Enumerable.Range(0, strokes.Count).Where(i => degrees[i] < 2));
        while (leaves.Count > 0)
        {
            int i = leaves.Dequeue();
            foreach (int j in strokes[i].Crossings)
                if (--degrees[j] == 1) leaves.Enqueue(j);
        }

        var sets = new Sets(strokes.Count);
        for (int i = 0; i < h.Count; i++)
            if (degrees[i] >= 2)
                foreach (int j in strokes[i].Crossings)
                    if (degrees[j] >= 2) sets.Join(i, j);
        var groups = Enumerable.Range(0, strokes.Count).Where(i => degrees[i] >= 2)
            .GroupBy(i => sets.Root(i), i => strokes[i]);
        foreach (var group in groups)
        {
            var table = Reconstruct(group.ToList(),tolerance);
            if (table != null) tables.Add(table);
        }
        return tables.OrderBy(t => t.Y[0]).ThenBy(t => t.X[0]).ToList();
    }

    // A present edge must cover the whole elementary interval. A partial edge means
    // the source cannot be represented faithfully as a rectangular Word cell grid.
    private static bool[,] EdgeCoverage(List<Stroke> lines,double[] axes,double[] along,double tolerance)
    {
        var result = new bool[axes.Length, along.Length-1];
        foreach (var line in lines)
        {
            int row = Array.BinarySearch(axes, line.At);
            for (int col = 0; col < along.Length-1; col++)
            {
                double left = along[col], right = along[col+1];
                if (line.End <= left + tolerance || line.Start >= right - tolerance) continue;
                if (line.Start > left + tolerance || line.End < right - tolerance) return null;
                result[row,col] = true;
            }
        }
        return result;
    }

    private static RuledTable Reconstruct(List<Stroke> strokes,double tolerance)
    {
        var h = strokes.Where(s => s.Horizontal).ToList();
        var v = strokes.Where(s => !s.Horizontal).ToList();
        var x = v.Select(s => s.At).Distinct().OrderBy(n => n).ToArray();
        var y = h.Select(s => s.At).Distinct().OrderBy(n => n).ToArray();
        int cols = x.Length-1, rows = y.Length-1;
        if (cols < 1 || rows < 1 || cols > 63 || rows > 2000 || rows*cols > 16000) return null;
        var horizontal = EdgeCoverage(h,y,x,tolerance);
        var vertical = EdgeCoverage(v,x,y,tolerance);
        if (horizontal == null || vertical == null) return null;
        var cells = new Sets(rows*cols);
        for (int r = 0; r < rows; r++)
        {
            if (!vertical[0,r] || !vertical[cols,r]) return null;
            for (int c = 0; c < cols; c++)
            {
                if (!horizontal[0,c] || !horizontal[rows,c]) return null;
                int index = r*cols+c;
                if (c > 0 && !vertical[c,r]) cells.Join(index, index-1);
                if (r > 0 && !horizontal[r,c]) cells.Join(index, index-cols);
            }
        }

        var table = new RuledTable { X = x, Y = y };
        foreach (var members in Enumerable.Range(0, rows*cols).GroupBy(i => cells.Root(i)))
        {
            int top = members.Min(i => i/cols), bottom = members.Max(i => i/cols);
            int left = members.Min(i => i%cols), right = members.Max(i => i%cols);
            if (members.Count() != (bottom-top+1)*(right-left+1)) return null;
            foreach (int i in members)
            {
                int r = i/cols, c = i%cols;
                if ((r > top && horizontal[r,c]) || (c > left && vertical[c,r])) return null;
            }
            table.Cells.Add(new RuledCell { Row = top, Col = left, RowSpan = bottom-top+1, ColSpan = right-left+1 });
        }
        // A single outline is also commonly a photo frame or a decorative page border.
        return table.Cells.Count > 1 ? table : null;
    }
}
