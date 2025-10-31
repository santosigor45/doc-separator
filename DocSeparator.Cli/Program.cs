using DocSeparator.Cli.Configuration;
using DocSeparator.Cli.Extraction;
using DocSeparator.Cli.Infrastructure;
using DocSeparator.Cli.Output;
using Microsoft.Extensions.Logging;

namespace DocSeparator.Cli;

internal static class Program
{
    private static int Main(string[] args)
    {
        if (!CommandLineOptions.TryParse(args, out var options, out var error, out var showHelp))
        {
            Console.Error.WriteLine(error);
            CommandLineOptions.PrintUsage();
            return 1;
        }

        if (showHelp)
        {
            CommandLineOptions.PrintUsage();
            return 0;
        }

        using var loggerFactory = LoggerFactory.Create(builder =>
        {
            builder
                .SetMinimumLevel(LogLevel.Information)
                .AddSimpleConsole(c =>
                {
                    c.IncludeScopes = false;
                    c.SingleLine = true;
                    c.TimestampFormat = "HH:mm:ss ";
                });
        });

        var logger = loggerFactory.CreateLogger("DocSeparator");

        try
        {
            var loader = new ConfigurationLoader(logger);
            var configuration = loader.Load(options.RegionConfigPath);

            var processor = new DocumentProcessor(logger);
            var table = processor.Process(options, configuration);

            ISpreadsheetWriter writer = options.Format switch
            {
                OutputFormat.Csv => new CsvSpreadsheetWriter(),
                OutputFormat.Xlsx => new XlsxSpreadsheetWriter(),
                _ => throw new ArgumentOutOfRangeException(nameof(options.Format), options.Format, null)
            };

            writer.Write(table, options.OutputPath);

            logger.LogInformation("Extraction completed. Rows written: {Count}. Output: {OutputPath}", table.Rows.Count, options.OutputPath);
            return 0;
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Processing failed: {Message}", ex.Message);
            return 1;
        }
    }
}
