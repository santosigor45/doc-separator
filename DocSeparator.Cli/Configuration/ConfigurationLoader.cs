using Microsoft.Extensions.Logging;
using System.Text.Json;

namespace DocSeparator.Cli.Configuration;

internal sealed class ConfigurationLoader
{
    private readonly ILogger _logger;

    public ConfigurationLoader(ILogger logger)
    {
        _logger = logger;
    }

    public AppConfiguration Load(string requestedPath)
    {
        var resolvedPath = ResolveConfigurationPath(requestedPath);
        _logger.LogInformation("Using region configuration: {Path}", resolvedPath);

        return Path.GetExtension(resolvedPath).Equals(".json", StringComparison.OrdinalIgnoreCase)
            ? LoadFromJson(resolvedPath)
            : LoadLegacy(resolvedPath);
    }

    private string ResolveConfigurationPath(string requestedPath)
    {
        var candidates = new List<string>();

        if (!string.IsNullOrWhiteSpace(requestedPath))
        {
            candidates.Add(requestedPath);
            if (!Path.IsPathRooted(requestedPath))
            {
                candidates.Add(Path.GetFullPath(requestedPath));
            }
        }

        var legacyPath = Path.GetFullPath("pdfregion.txt");
        if (!candidates.Contains(legacyPath, StringComparer.OrdinalIgnoreCase))
        {
            candidates.Add(legacyPath);
        }

        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new FileNotFoundException($"Region configuration not found. Looked in: {string.Join(", ", candidates)}", requestedPath);
    }

    private AppConfiguration LoadFromJson(string path)
    {
        var json = File.ReadAllText(path);
        var options = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true
        };

        var dto = JsonSerializer.Deserialize<AppConfigurationDto>(json, options)
                  ?? throw new InvalidOperationException("Configuration file is empty or unreadable.");

        if (dto.Regions is null || dto.Regions.Count == 0)
        {
            throw new InvalidOperationException("Configuration must declare at least one region.");
        }

        var regions = new List<RegionDefinition>(dto.Regions.Count);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var regionDto in dto.Regions)
        {
            if (string.IsNullOrWhiteSpace(regionDto.Name))
            {
                throw new InvalidOperationException("Each region must have a non-empty name.");
            }

            if (!names.Add(regionDto.Name))
            {
                throw new InvalidOperationException($"Duplicate region name detected: '{regionDto.Name}'.");
            }

            if (regionDto.Rectangle is null)
            {
                throw new InvalidOperationException($"Region '{regionDto.Name}' is missing rectangle coordinates.");
            }

            var filter = ParsePageScope(regionDto.PageScope);

            regions.Add(new RegionDefinition(
                regionDto.Name,
                new RegionRectangle(
                    regionDto.Rectangle.Left,
                    regionDto.Rectangle.Top,
                    regionDto.Rectangle.Right,
                    regionDto.Rectangle.Bottom),
                filter));
        }

        var parallelism = dto.Parallelism > 0 ? dto.Parallelism : (int?)null;
        return new AppConfiguration(regions, parallelism);
    }

    private AppConfiguration LoadLegacy(string path)
    {
        var lines = File.ReadAllLines(path);
        var regions = new List<RegionDefinition>();
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var index = 1;

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (string.IsNullOrEmpty(line))
            {
                continue;
            }

            string name;
            string coordinatePart;

            var separatorIndex = line.IndexOfAny(new[] { ':', '=' });
            if (separatorIndex >= 0)
            {
                name = line[..separatorIndex].Trim();
                coordinatePart = line[(separatorIndex + 1)..].Trim();
            }
            else
            {
                name = $"Region{index}";
                coordinatePart = line;
            }

            if (!names.Add(name))
            {
                name = $"Region{index}";
                names.Add(name);
            }

            var components = coordinatePart.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (components.Length != 4)
            {
                throw new InvalidOperationException($"Invalid coordinate specification '{coordinatePart}' in line '{line}'. Expected format: left,top,right,bottom.");
            }

            if (!double.TryParse(components[0], out var left) ||
                !double.TryParse(components[1], out var top) ||
                !double.TryParse(components[2], out var right) ||
                !double.TryParse(components[3], out var bottom))
            {
                throw new InvalidOperationException($"Could not parse rectangle values in line '{line}'.");
            }

            regions.Add(new RegionDefinition(
                name,
                new RegionRectangle(left, top, right, bottom),
                PageFilter.All));

            index++;
        }

        if (regions.Count == 0)
        {
            throw new InvalidOperationException("Legacy configuration contained no region definitions.");
        }

        _logger.LogWarning("Loaded legacy region configuration from {Path}. Consider migrating to config/regions.json for richer options.", path);
        return new AppConfiguration(regions, null);
    }

    private static PageFilter ParsePageScope(string? scope)
    {
        if (string.IsNullOrWhiteSpace(scope))
        {
            return PageFilter.All;
        }

        scope = scope.Trim().ToLowerInvariant();
        return scope switch
        {
            "all" => PageFilter.All,
            "even" => PageFilter.Even,
            "odd" => PageFilter.Odd,
            _ => ParseRanges(scope)
        };
    }

    private static PageFilter ParseRanges(string scope)
    {
        var tokens = scope.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (tokens.Length == 0)
        {
            return PageFilter.All;
        }

        var ranges = new List<(int Start, int End)>(tokens.Length);
        foreach (var token in tokens)
        {
            if (token.Contains('-', StringComparison.Ordinal))
            {
                var parts = token.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                if (parts.Length != 2 ||
                    !int.TryParse(parts[0], out var start) ||
                    !int.TryParse(parts[1], out var end))
                {
                    throw new InvalidOperationException($"Invalid page range token '{token}'.");
                }

                if (start <= 0 || end <= 0 || end < start)
                {
                    throw new InvalidOperationException($"Invalid page range '{token}'. Page numbers must be positive and increasing.");
                }

                ranges.Add((start, end));
            }
            else
            {
                if (!int.TryParse(token, out var page) || page <= 0)
                {
                    throw new InvalidOperationException($"Invalid page number '{token}'.");
                }

                ranges.Add((page, page));
            }
        }

        return PageFilter.FromInclusiveRanges(ranges);
    }

    private sealed record AppConfigurationDto(List<RegionDto> Regions, int? Parallelism);

    private sealed record RegionDto(string Name, RectangleDto Rectangle, string? PageScope);

    private sealed record RectangleDto(double Left, double Top, double Right, double Bottom);
}
