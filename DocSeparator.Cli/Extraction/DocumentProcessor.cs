using System.Collections.Concurrent;
using System.Collections.Immutable;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocSeparator.Cli.Configuration;
using DocSeparator.Cli.Infrastructure;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Geometry;

namespace DocSeparator.Cli.Extraction;

internal sealed class DocumentProcessor
{
    private readonly ILogger _logger;

    public DocumentProcessor(ILogger logger)
    {
        _logger = logger;
    }

    public ExtractionTable Process(CommandLineOptions options, AppConfiguration configuration)
    {
        var files = ResolveInputFiles(options.InputPath);
        if (files.Count == 0)
        {
            throw new FileNotFoundException($"No PDF files found at '{options.InputPath}'.");
        }

        var regionNames = configuration.Regions.Select(r => r.Name).ToImmutableArray();
        var table = new ExtractionTable(regionNames);

        var maxParallelism = DetermineParallelism(options, configuration);
        _logger.LogInformation("Using up to {Parallelism} parallel workers.", maxParallelism);

        foreach (var file in files)
        {
            var stopwatch = Stopwatch.StartNew();
            var documentRows = ProcessDocument(file, configuration.Regions, maxParallelism);
            stopwatch.Stop();

            foreach (var row in documentRows)
            {
                table.AddRow(row);
            }

            _logger.LogInformation(
                "Processed {File} ({PageCount} pages) in {ElapsedMs} ms ({Throughput:F1} pages/sec).",
                file,
                documentRows.Count,
                stopwatch.ElapsedMilliseconds,
                documentRows.Count == 0 ? 0 : documentRows.Count / Math.Max(0.001, stopwatch.Elapsed.TotalSeconds));
        }

        return table;
    }

    private static IReadOnlyList<string> ResolveInputFiles(string inputPath)
    {
        if (File.Exists(inputPath))
        {
            return string.Equals(Path.GetExtension(inputPath), ".pdf", StringComparison.OrdinalIgnoreCase)
                ? new[] { Path.GetFullPath(inputPath) }
                : Array.Empty<string>();
        }

        if (Directory.Exists(inputPath))
        {
            return Directory.EnumerateFiles(inputPath, "*.pdf", SearchOption.AllDirectories)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }

        return Array.Empty<string>();
    }

    private static int DetermineParallelism(CommandLineOptions options, AppConfiguration configuration)
    {
        if (options.MaxDegreeOfParallelism.HasValue)
        {
            return options.MaxDegreeOfParallelism.Value;
        }

        if (configuration.MaxDegreeOfParallelism.HasValue)
        {
            return configuration.MaxDegreeOfParallelism.Value;
        }

        var logical = Environment.ProcessorCount;
        return Math.Max(1, logical - 1);
    }

    private List<ExtractionRow> ProcessDocument(string filePath, IReadOnlyList<RegionDefinition> regions, int maxParallelism)
    {
        int pageCount;
        using (var pdf = PdfDocument.Open(filePath))
        {
            pageCount = pdf.NumberOfPages;
        }

        var rows = new ConcurrentDictionary<int, ExtractionRow>();
        var regionArray = regions.ToArray();
        var documentName = Path.GetFileName(filePath);
        var pdfBytes = File.ReadAllBytes(filePath);

        if (pageCount == 0)
        {
            return new List<ExtractionRow>();
        }

        if (maxParallelism <= 1 || pageCount == 1)
        {
            using var document = PdfDocument.Open(new ReadOnlyMemoryStream(pdfBytes));
            for (var pageNumber = 1; pageNumber <= pageCount; pageNumber++)
            {
                var row = ExtractPage(document, pageNumber, regionArray, documentName);
                rows[pageNumber] = row;
            }
        }
        else
        {
            var ranges = Partitioner.Create(1, pageCount + 1);
            Parallel.ForEach(
                ranges,
                new ParallelOptions { MaxDegreeOfParallelism = maxParallelism },
                range =>
                {
                    using var document = PdfDocument.Open(new ReadOnlyMemoryStream(pdfBytes));
                    for (var pageNumber = range.Item1; pageNumber < range.Item2; pageNumber++)
                    {
                        var row = ExtractPage(document, pageNumber, regionArray, documentName);
                        rows[pageNumber] = row;
                    }
                });
        }

        return rows
            .OrderBy(kvp => kvp.Key)
            .Select(kvp => kvp.Value)
            .ToList();
    }

    private static ExtractionRow ExtractPage(PdfDocument document, int pageNumber, IReadOnlyList<RegionDefinition> regions, string documentName)
    {
        var page = document.GetPage(pageNumber);
        var pageHeight = page.Height;
        var words = page.GetWords().ToList();

        var regionTexts = new Dictionary<string, string>(regions.Count, StringComparer.Ordinal);
        foreach (var region in regions)
        {
            if (!region.Filter.Includes(pageNumber))
            {
                regionTexts[region.Name] = string.Empty;
                continue;
            }

            var rect = region.Rectangle.ToPdfRectangle(pageHeight);
            var text = ExtractText(words, rect);
            regionTexts[region.Name] = text;
        }

        return new ExtractionRow(documentName, pageNumber, regionTexts);
    }

    private static string ExtractText(IReadOnlyList<Word> words, PdfRectangle region)
    {
        const double lineThreshold = 3.0;
        var filtered = new List<Word>();

        foreach (var word in words)
        {
            if (Intersects(region, word.BoundingBox))
            {
                filtered.Add(word);
            }
        }

        if (filtered.Count == 0)
        {
            return string.Empty;
        }

        filtered.Sort((a, b) =>
        {
            var vertical = b.BoundingBox.Top.CompareTo(a.BoundingBox.Top);
            return vertical != 0 ? vertical : a.BoundingBox.Left.CompareTo(b.BoundingBox.Left);
        });

        var builder = new System.Text.StringBuilder();
        double? currentLine = null;

        foreach (var word in filtered)
        {
            if (currentLine is null)
            {
                currentLine = word.BoundingBox.Top;
            }
            else if (Math.Abs(currentLine.Value - word.BoundingBox.Top) > lineThreshold)
            {
                builder.AppendLine();
                currentLine = word.BoundingBox.Top;
            }
            else if (builder.Length > 0)
            {
                builder.Append(' ');
            }

            builder.Append(word.Text);
        }

        return builder.ToString().Trim();
    }

    private static bool Intersects(PdfRectangle region, PdfRectangle other)
    {
        return !(other.Left > region.Right ||
                 other.Right < region.Left ||
                 other.Bottom > region.Top ||
                 other.Top < region.Bottom);
    }

    private sealed class ReadOnlyMemoryStream : MemoryStream
    {
        public ReadOnlyMemoryStream(byte[] buffer)
            : base(buffer, writable: false)
        {
        }
    }
}
