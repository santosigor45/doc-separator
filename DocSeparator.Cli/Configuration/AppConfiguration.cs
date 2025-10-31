namespace DocSeparator.Cli.Configuration;

internal sealed class AppConfiguration
{
    public AppConfiguration(IReadOnlyList<RegionDefinition> regions, int? maxDegreeOfParallelism)
    {
        Regions = regions;
        MaxDegreeOfParallelism = maxDegreeOfParallelism;
    }

    public IReadOnlyList<RegionDefinition> Regions { get; }

    public int? MaxDegreeOfParallelism { get; }
}
