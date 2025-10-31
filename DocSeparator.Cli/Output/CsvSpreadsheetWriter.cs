using System.Text;

namespace DocSeparator.Cli.Output;

internal sealed class CsvSpreadsheetWriter : ISpreadsheetWriter
{
    public void Write(ExtractionTable table, string outputPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");

        using var writer = new StreamWriter(outputPath, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        var headers = new List<string> { "Document", "Page" };
        headers.AddRange(table.RegionNames);
        writer.WriteLine(string.Join(",", headers.Select(Quote)));

        foreach (var row in table.Rows)
        {
            var values = new List<string>(headers.Count)
            {
                Quote(row.DocumentName),
                Quote(row.PageNumber.ToString())
            };

            foreach (var regionName in table.RegionNames)
            {
                row.RegionTexts.TryGetValue(regionName, out var text);
                values.Add(Quote(text ?? string.Empty));
            }

            writer.WriteLine(string.Join(",", values));
        }
    }

    private static string Quote(string value)
    {
        if (value.IndexOfAny(new[] { ',', '"', '\n', '\r' }) >= 0)
        {
            return $"\"{value.Replace("\"", "\"\"")}\"";
        }

        return value;
    }
}
