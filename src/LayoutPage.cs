using System.Collections.Generic;

// All coordinates use points from the top left of the visible PDF page.
internal sealed class LayoutPage
{
    internal double Width, Height;
    internal List<LayoutLine> Lines = new List<LayoutLine>();
    internal List<LayoutRule> Rules = new List<LayoutRule>();
    internal List<RuledTable> Tables = new List<RuledTable>();
    internal int TableCount, CellCount;
}

internal sealed class LayoutLine
{
    internal double Left, Top, Right, Bottom, FontSize, Baseline;
    internal double InkLeft, InkRight;
    internal string Text, FontName;
    internal bool Bold;
    internal int LinkPage; // One-based internal destination, zero for ordinary text.
}

internal sealed class LayoutRule
{
    internal double Left, Top, Right, Bottom;
    internal double Width = 0.5;
    internal int Gray = 128;
    internal bool IsHorizontal { get { return Right - Left > Bottom - Top; } }
}
