using System;
using System.Collections.Generic;
using System.Linq;

internal sealed class RuledTable
{
    internal double[] X, Y;
    internal List<RuledCell> Cells = new List<RuledCell>();
    internal double BorderWidth = 0.5;
    internal int BorderGray = 128;
}

internal sealed class RuledCell
{
    internal int Row, Col, RowSpan, ColSpan;
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
    private static List<Stroke> Normalize(IList<LayoutRule> rules, bool horizontal)
    {
        var raw = rules.Where(r => r.IsHorizontal == horizontal).Select(r => new Stroke {
            Horizontal = horizontal,
            At = horizontal ? (r.Top + r.Bottom) / 2 : (r.Left + r.Right) / 2,
            Start = horizontal ? r.Left : r.Top,
            End = horizontal ? r.Right : r.Bottom
        }).Where(s => s.End > s.Start).OrderBy(s => s.At).ToList();
        var result = new List<Stroke>();
        for (int i = 0; i < raw.Count; )
        {
            int j = i + 1;
            while (j < raw.Count && raw[j].At - raw[i].At <= Tolerance) j++;
            var band = raw.GetRange(i, j-i);
            // Choose an observed axis, weighted by the source stroke length.
            double axis = band.OrderByDescending(s => s.End-s.Start).First().At;
            Stroke current = null;
            foreach (var s in band.OrderBy(s => s.Start))
            {
                if (current == null || s.Start > current.End + Tolerance)
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
        var tables = new List<RuledTable>();
        if (rules == null || rules.Count > 16000) return tables;
        var h = Normalize(rules, true);
        var v = Normalize(rules, false);
        if ((long)h.Count * v.Count > 4000000) return tables;
        var strokes = h.Concat(v).ToList();
        int totalCrossings = 0;
        for (int i = 0; i < h.Count; i++)
        {
            var a = h[i];
            for (int j = 0; j < v.Count; j++)
            {
                var b = v[j];
                if (b.At < a.Start-Tolerance || a.End+Tolerance < b.At ||
                    a.At < b.Start-Tolerance || b.End+Tolerance < a.At) continue;
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
            var table = Reconstruct(group.ToList());
            if (table != null) tables.Add(table);
        }
        return tables.OrderBy(t => t.Y[0]).ThenBy(t => t.X[0]).ToList();
    }

    // A present edge must cover the whole elementary interval. A partial edge means
    // the source cannot be represented faithfully as a rectangular Word cell grid.
    private static bool[,] EdgeCoverage(List<Stroke> lines, double[] axes, double[] along)
    {
        var result = new bool[axes.Length, along.Length-1];
        foreach (var line in lines)
        {
            int row = Array.BinarySearch(axes, line.At);
            for (int col = 0; col < along.Length-1; col++)
            {
                double left = along[col], right = along[col+1];
                if (line.End <= left + Tolerance || line.Start >= right - Tolerance) continue;
                if (line.Start > left + Tolerance || line.End < right - Tolerance) return null;
                result[row,col] = true;
            }
        }
        return result;
    }

    private static RuledTable Reconstruct(List<Stroke> strokes)
    {
        var h = strokes.Where(s => s.Horizontal).ToList();
        var v = strokes.Where(s => !s.Horizontal).ToList();
        var x = v.Select(s => s.At).Distinct().OrderBy(n => n).ToArray();
        var y = h.Select(s => s.At).Distinct().OrderBy(n => n).ToArray();
        int cols = x.Length-1, rows = y.Length-1;
        if (cols < 1 || rows < 1 || cols > 63 || rows > 2000 || rows*cols > 16000) return null;
        var horizontal = EdgeCoverage(h, y, x);
        var vertical = EdgeCoverage(v, x, y);
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
