namespace DocSeparator.Cli.Infrastructure;

internal enum OutputFormat
{
    Xlsx,
    Csv
}

internal sealed record CommandLineOptions(
    string InputPath,
    string RegionConfigPath,
    string OutputPath,
    OutputFormat Format,
    int? MaxDegreeOfParallelism);

internal static class CommandLineOptionsExtensions
{
    public static bool TryParse(
        string[] args,
        out CommandLineOptions options,
        out string error,
        out bool showHelp)
    {
        options = default!;
        error = string.Empty;
        showHelp = false;

        if (args.Any(arg => arg is "-h" or "--help" or "-?"))
        {
            showHelp = true;
            return true;
        }

        var parsed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        for (var index = 0; index < args.Length; index++)
        {
            var arg = args[index];
            if (!arg.StartsWith("--", StringComparison.Ordinal))
            {
                error = $"Unexpected argument '{arg}'. Expected --key value pairs.";
                return false;
            }

            var key = arg[2..];
            if (parsed.ContainsKey(key))
            {
                error = $"Duplicate argument '--{key}'.";
                return false;
            }

            if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                error = $"Missing value for '--{key}'.";
                return false;
            }

            parsed[key] = args[++index];
        }

        if (!parsed.TryGetValue("input", out var inputPath))
        {
            error = "Missing required argument '--input <path>' (PDF file or directory).";
            return false;
        }

        OutputFormat format = OutputFormat.Xlsx;
        if (parsed.TryGetValue("format", out var formatValue))
        {
            switch (formatValue.Trim().ToLowerInvariant())
            {
                case "xlsx":
                    format = OutputFormat.Xlsx;
                    break;
                case "csv":
                    format = OutputFormat.Csv;
                    break;
                default:
                    error = "Unknown format. Valid values: xlsx, csv.";
                    return false;
            }
        }

        string outputPath;
        if (parsed.TryGetValue("output", out var explicitOutput))
        {
            outputPath = explicitOutput;
        }
        else
        {
            var extension = format == OutputFormat.Csv ? ".csv" : ".xlsx";
            outputPath = Path.Combine(Environment.CurrentDirectory, $"extracted{extension}");
        }

        string regionConfigPath = parsed.TryGetValue("regions", out var regionsPath)
            ? regionsPath
            : "config/regions.json";

        int? maxDegreeOfParallelism = null;
        if (parsed.TryGetValue("parallelism", out var parallelismValue))
        {
            if (!int.TryParse(parallelismValue, out var parsedParallelism) || parsedParallelism <= 0)
            {
                error = "Invalid '--parallelism' value. Provide a positive integer.";
                return false;
            }

            maxDegreeOfParallelism = parsedParallelism;
        }

        options = new CommandLineOptions(
            inputPath,
            regionConfigPath,
            outputPath,
            format,
            maxDegreeOfParallelism);

        return true;
    }

    public static void PrintUsage()
    {
        const string usage = """
Usage:
  doc-separator --input <path> [--regions <path>] [--output <file>] [--format xlsx|csv] [--parallelism <n>]

Arguments:
  --input        PDF file or directory to process (required).
  --regions      Path to region configuration (default: config/regions.json, falls back to pdfregion.txt).
  --output       Output spreadsheet path (default: ./extracted.xlsx or .csv).
  --format       Output format: xlsx (default) or csv.
  --parallelism  Maximum parallel workers (defaults to configuration or CPU count).
  --help         Show this message.
""";
        Console.WriteLine(usage);
    }
}
