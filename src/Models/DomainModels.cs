namespace ElTools.Models;

public sealed record MappingRule(string SourceBlockName, string TargetBlockName, string HeightSourceTag = "ВЫСОТА", string HeightTargetTag = "ВЫСОТА");
public sealed record MappingProfile(string Name, IReadOnlyList<MappingRule> Rules);
public sealed record TraceRequest(ObjectId PolylineId, ObjectId SourceBlockId, ObjectId TargetBlockId, string CableMark);
public sealed record TraceResult(double Length2D, double HeightDelta, double TotalLength, string CableMark);
public sealed record EomDataRecord(string CableMark, double TotalLength, string Group, string Type);
public sealed record SpecificationRow(string CableType, string Group, double TotalLength);

public sealed class SettingsModel
{
    public string ActiveProfile { get; set; } = "Default";
    public List<MappingProfile> MappingProfiles { get; set; } = new();
}

public enum LicenseState
{
    Disabled,
    Valid,
    Invalid,
    Expired
}
