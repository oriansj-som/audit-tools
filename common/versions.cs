public class versions
{
    public readonly int ancient      = 31335;            // 2016 records [Remove after 2026]
    public readonly int pregenerated = 31336;            // 2017-2018 records [Remove after 2028]
    public readonly int prototype    = 31337;            // 2019-2022 records [Remove after 2032]
    public readonly int cms_22       = 31339;            // 2023-2024 records [Remove after 2035]
    public readonly int arc_ampe     = 31340;            // 2024-now records

    private int? Version;
    public string? Agency;
    public int? Year;

    public string debug()
    {
        return string.Format("version: {0}, Agency: {1}, Year: {2}", Version, Agency, Year);
    }
    public versions()
    {
        Version = arc_ampe;
    }

    public versions(int version)
    {
        Version = version;
    }
    public versions(int version, string agency, int year)
    {
        Version = version;
        Agency = agency;
        Year = year;
    }

    public bool ancientp()
    {
        return ancient == Version;
    }

    public bool pregeneratedp()
    {
        return pregenerated == Version;
    }

    public bool prototypep()
    {
        return prototype == Version;
    }

    public bool cms_22p()
    {
        return cms_22 == Version;
    }

    public bool arc_ampep()
    {
        return arc_ampe == Version;
    }
}
