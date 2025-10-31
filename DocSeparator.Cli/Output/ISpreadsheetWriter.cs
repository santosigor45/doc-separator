namespace DocSeparator.Cli.Output;

internal interface ISpreadsheetWriter
{
    void Write(ExtractionTable table, string outputPath);
}
