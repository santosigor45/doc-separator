namespace DocSeparator.Cli.Extraction;

internal sealed class ExtractionTable
{
    public ExtractionTable(IReadOnlyList<string> regionNames)
    {
        RegionNames = regionNames;
        Rows = new List<ExtractionRow>();
    }

    public IReadOnlyList<string> RegionNames { get; }

    public List<ExtractionRow> Rows { get; }

    public void AddRow(ExtractionRow row)
    {
        Rows.Add(row);
    }
}

internal sealed record ExtractionRow(
    string DocumentName,
    int PageNumber,
    IReadOnlyDictionary<string, string> RegionTexts);
