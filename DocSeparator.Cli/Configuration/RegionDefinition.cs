using UglyToad.PdfPig.Geometry;

namespace DocSeparator.Cli.Configuration;

internal sealed class RegionDefinition
{
    public RegionDefinition(string name, RegionRectangle rectangle, PageFilter filter)
    {
        Name = name;
        Rectangle = rectangle;
        Filter = filter;
    }

    public string Name { get; }

    public RegionRectangle Rectangle { get; }

    public PageFilter Filter { get; }
}

internal readonly struct RegionRectangle
{
    public RegionRectangle(double left, double top, double right, double bottom)
    {
        if (right <= left || bottom <= top)
        {
            throw new ArgumentException("Invalid rectangle coordinates (right must be > left and bottom > top).");
        }

        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
    }

    public double Left { get; }

    public double Top { get; }

    public double Right { get; }

    public double Bottom { get; }

    public PdfRectangle ToPdfRectangle(double pageHeight)
    {
        // Legacy coordinates assume origin at top-left; PdfPig uses bottom-left.
        var lower = Math.Max(0, pageHeight - Bottom);
        var upper = Math.Min(pageHeight, pageHeight - Top);
        return new PdfRectangle(Left, lower, Right, upper);
    }
}

internal sealed class PageFilter
{
    private readonly Func<int, bool> _predicate;

    private PageFilter(Func<int, bool> predicate, string description)
    {
        _predicate = predicate;
        Description = description;
    }

    public static PageFilter All { get; } = new(page => true, "all pages");

    public string Description { get; }

    public bool Includes(int pageNumber) => _predicate(pageNumber);

    public static PageFilter Even { get; } = new(page => page % 2 == 0, "even pages");

    public static PageFilter Odd { get; } = new(page => page % 2 != 0, "odd pages");

    public static PageFilter FromInclusiveRanges(IReadOnlyList<(int Start, int End)> ranges)
    {
        return new PageFilter(page =>
        {
            foreach (var range in ranges)
            {
                if (page >= range.Start && page <= range.End)
                {
                    return true;
                }
            }

            return false;
        }, string.Join(",", ranges.Select(r => r.Start == r.End ? r.Start.ToString() : $"{r.Start}-{r.End}")));
    }
}
