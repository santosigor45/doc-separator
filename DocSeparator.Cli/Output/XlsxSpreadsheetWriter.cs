using ClosedXML.Excel;

namespace DocSeparator.Cli.Output;

internal sealed class XlsxSpreadsheetWriter : ISpreadsheetWriter
{
    public void Write(ExtractionTable table, string outputPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Extraction");

        worksheet.Cell(1, 1).Value = "Document";
        worksheet.Cell(1, 2).Value = "Page";

        for (var i = 0; i < table.RegionNames.Count; i++)
        {
            worksheet.Cell(1, i + 3).Value = table.RegionNames[i];
        }

        var currentRow = 2;
        foreach (var row in table.Rows)
        {
            worksheet.Cell(currentRow, 1).Value = row.DocumentName;
            worksheet.Cell(currentRow, 2).Value = row.PageNumber;

            for (var i = 0; i < table.RegionNames.Count; i++)
            {
                row.RegionTexts.TryGetValue(table.RegionNames[i], out var value);
                worksheet.Cell(currentRow, i + 3).Value = value ?? string.Empty;
            }

            currentRow++;
        }

        worksheet.Columns().AdjustToContents();
        workbook.SaveAs(outputPath);
    }
}
