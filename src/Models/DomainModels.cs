// FILE: src/Models/DomainModels.cs
// VERSION: 1.2.0
// START_MODULE_CONTRACT
//   PURPOSE: Define shared domain models used by mapping, tracing, licensing, and specification workflows.
//   SCOPE: Immutable records for operational data, settings root model, and license state enum.
//   DEPENDS: none
//   LINKS: M-MODELS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   MappingRule/MappingProfile - Block mapping profile and rule model.
//   TraceRequest/TraceResult - Cable trace input and calculated output model.
//   EomDataRecord/SpecificationRow - Persisted EOM data and grouped specification row.
//   SettingsModel/LicenseState - Plugin settings root and license validation state.
//   PanelLayoutMapConfig/OlsSelectedDevice - OLS selection parsing and selector/legacy layout mapping models.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY

namespace ElTools.Models;

public sealed record MappingRule(string SourceBlockName, string TargetBlockName, string HeightSourceTag = "Р вЂ™Р В«Р РЋР С›Р СћР С’", string HeightTargetTag = "Р вЂ™Р В«Р РЋР С›Р СћР С’");
public sealed record MappingProfile(string Name, IReadOnlyList<MappingRule> Rules);
public sealed record TraceRequest(ObjectId PolylineId, ObjectId SourceBlockId, ObjectId TargetBlockId, string CableMark);
public sealed record TraceResult(double Length2D, double HeightDelta, double TotalLength, string CableMark);
public sealed record EomDataRecord(string CableMark, double TotalLength, string Group, string Type);
public sealed record SpecificationRow(string CableType, string Group, double TotalLength);
public sealed record LoadBlockData(
    ObjectId BlockId,
    string Group,
    double PowerWatts,
    double Voltage,
    string Shield,
    string? Phase = null,
    string? Room = null,
    string? Note = null);
public sealed record InstallTypeRule(int Priority, string MatchBy, string Value, string Result);
public sealed record InstallTypeRuleSet(string Default, IReadOnlyList<InstallTypeRule> Rules);
public sealed record GroupTraceAggregate(
    string Shield,
    string Group,
    double TotalPowerWatts,
    double TotalLengthMeters,
    IReadOnlyDictionary<string, double> LengthByInstallType);
public sealed record ExcelInputRow(
    string Shield,
    string Group,
    double PowerKw,
    double Voltage,
    double TotalLengthMeters,
    double CeilingLengthMeters,
    double FloorLengthMeters,
    double RiserLengthMeters,
    string? Phase = null,
    string? GroupType = null);
public sealed record ExcelOutputRow(
    string Shield,
    string Group,
    string CircuitBreaker,
    string RcdDiff,
    string Cable,
    int CircuitBreakerModules,
    int RcdModules,
    string? Note = null);
public sealed record ValidationIssue(string Code, ObjectId EntityId, string Message, string Severity = "Warning");
public sealed record OlsSelectedDevice(
    ObjectId EntityId,
    string SourceBlockName,
    OlsSourceSignature SourceSignature,
    string? DeviceKey,
    int Modules,
    string? Group,
    string? Note,
    Point3d InsertionPoint);
public sealed record OlsSourceSignature(
    string SourceBlockName,
    string? VisibilityValue = null);
public sealed record SkippedOlsDeviceIssue(
    ObjectId EntityId,
    string Reason,
    string? DeviceKey = null,
    string? SourceBlockName = null);
public sealed record PanelLayoutSelectorRule(
    int Priority,
    string SourceBlockName,
    string? VisibilityValue,
    string LayoutBlockName,
    int? FallbackModules = null);
public sealed record PanelLayoutMapRule(
    string DeviceKey,
    string LayoutBlockName,
    int? FallbackModules = null);

public sealed class PanelLayoutAttributeTags
{
    public string Device { get; set; } = "Р С’Р СџР СџР С’Р В Р С’Р Сћ";
    public string Modules { get; set; } = "Р СљР С›Р вЂќР Р€Р вЂєР вЂўР в„ў";
    public string Group { get; set; } = "Р вЂњР В Р Р€Р СџР СџР С’";
    public string Note { get; set; } = "Р СџР В Р ВР СљР вЂўР В§Р С’Р СњР ВР вЂў";
}

public sealed class PanelLayoutMapConfig
{
    public string Version { get; set; } = "2.0";
    public int DefaultModulesPerRow { get; set; } = 24;
    public PanelLayoutAttributeTags AttributeTags { get; set; } = new();
    public List<PanelLayoutSelectorRule> SelectorRules { get; set; } = new();
    public List<PanelLayoutMapRule> LayoutMap { get; set; } = new();
}

public sealed class SettingsModel
{
    public string ActiveProfile { get; set; } = "Default";
    public List<MappingProfile> MappingProfiles { get; set; } = new();
    public InstallTypeRuleSet InstallTypeRules { get; set; } = new(
        "Р СњР ВµР С•Р С—РЎР‚Р ВµР Т‘Р ВµР В»Р ВµР Р…Р С•",
        new List<InstallTypeRule>
        {
            new(1, "Linetype", "E_CEIL", "Р СџР С•РЎвЂљР С•Р В»Р С•Р С”"),
            new(2, "Layer", "Р В­Р С›Р Сљ_Р СџР С›Р вЂє", "Р СџР С•Р В»"),
            new(3, "Layer", "Р В­Р С›Р Сљ_Р РЋР СћР С›Р Р‡Р С™", "Р РЋРЎвЂљР С•РЎРЏР С”")
        });
    public string ExcelTemplatePath { get; set; } = "Р В Р В°РЎРѓРЎвЂЎР ВµРЎвЂљ_Р РЃР В°Р В±Р В»Р С•Р Р….xlsx";
    public int PanelModulesPerRow { get; set; } = 24;
    public string? GroupRegex { get; set; }
}

public enum LicenseState
{
    Disabled,
    Valid,
    Invalid,
    Expired
}