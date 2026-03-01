// FILE: src/Shared/PluginConfig.cs
// VERSION: 1.4.0
// START_MODULE_CONTRACT
//   PURPOSE: Keep canonical constants for attributes, commands, metadata keys, and template contracts.
//   SCOPE: Shared static values reused by command handlers and services.
//   DEPENDS: none
//   LINKS: M-CONFIG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   Strings - Canonical Russian labels encoded via Unicode escapes.
//   AttributeTags - Russian block attribute names.
//   Commands - AutoCAD command names exposed by plugin.
//   Metadata - XData and dictionary keys.
//   TemplateBlocks - Optional block names for OLS drawing.
//   PanelLayout - Panel layout mapping config and OLS attribute tag defaults.
// END_MODULE_MAP

namespace ElTools.Shared;

public static class PluginConfig
{
    public static class Strings
    {
        public const string Group = "\u0413\u0420\u0423\u041f\u041f\u0410";
        public const string Power = "\u041c\u041e\u0429\u041d\u041e\u0421\u0422\u042c";
        public const string Voltage = "\u041d\u0410\u041f\u0420\u042f\u0416\u0415\u041d\u0418\u0415";
        public const string Shield = "\u0429\u0418\u0422";
        public const string Phase = "\u0424\u0410\u0417\u0410";
        public const string Room = "\u041f\u041e\u041c\u0415\u0429\u0415\u041d\u0418\u0415";
        public const string Note = "\u041f\u0420\u0418\u041c\u0415\u0427\u0410\u041d\u0418\u0415";
        public const string InstallType = "\u0422\u0418\u041f_\u041f\u0420\u041e\u041a\u041b\u0410\u0414\u041a\u0418";
        public const string Unknown = "\u041d\u0435\u043e\u043f\u0440\u0435\u0434\u0435\u043b\u0435\u043d\u043e";
        public const string Ceiling = "\u041f\u043e\u0442\u043e\u043b\u043e\u043a";
        public const string Floor = "\u041f\u043e\u043b";
        public const string Riser = "\u0421\u0442\u043e\u044f\u043a";
    }

    public static class AttributeTags
    {
        public const string Group = Strings.Group;
        public const string Power = Strings.Power;
        public const string Voltage = Strings.Voltage;
        public const string Shield = Strings.Shield;
        public const string Phase = Strings.Phase;
        public const string Room = Strings.Room;
        public const string Note = Strings.Note;
    }

    public static class Commands
    {
        public const string Map = "EOM_MAP";
        public const string Trace = "EOM_TRACE";
        public const string Spec = "EOM_SPEC";
        public const string MapConfig = "EOM_MAPCFG";
        public const string Update = "EOM_\u041e\u0411\u041d\u041e\u0412\u0418\u0422\u042c";
        public const string ExportExcel = "EOM_\u042d\u041a\u0421\u041f\u041e\u0420\u0422_EXCEL";
        public const string ImportExcel = "EOM_\u0418\u041c\u041f\u041e\u0420\u0422_EXCEL";
        public const string BuildOls = "EOM_\u041f\u041e\u0421\u0422\u0420\u041e\u0418\u0422\u042c_\u041e\u041b\u0421";
        public const string BuildPanelLayout = "EOM_\u041a\u041e\u041c\u041f\u041e\u041d\u041e\u0412\u041a\u0410_\u0429\u0418\u0422\u0410";
        public const string PanelLayoutConfig = "EOM_LAYOUTCFG";
        public const string BindPanelLayoutVisualization = "EOM_\u0421\u0412\u042f\u0417\u0410\u0422\u042c_\u0412\u0418\u0417\u0423\u0410\u041b\u0418\u0417\u0410\u0426\u0418\u042e";
        public const string BindPanelLayoutVisualizationAlias = "EOM_BIND_VIS";
        public const string Validate = "EOM_\u041f\u0420\u041e\u0412\u0415\u0420\u041a\u0410";
        public const string ActiveGroup = "EOM_\u0410\u041a\u0422\u0418\u0412\u041d\u0410\u042f_\u0413\u0420\u0423\u041f\u041f\u0410";
        public const string AssignGroup = "EOM_\u041d\u0410\u0417\u041d\u0410\u0427\u0418\u0422\u042c_\u0413\u0420\u0423\u041f\u041f\u0423";
        public const string InstallTypeSettings = "EOM_\u041d\u0410\u0421\u0422\u0420\u041e\u0419\u041a\u0418_\u041f\u0420\u041e\u041a\u041b\u0410\u0414\u041a\u0418";
    }

    public static class Metadata
    {
        public const string EomDataAppName = "EOM_DATA";
        public const string ElToolAppName = "ELTOOL";
        public const string OlsLayoutAppName = "ELTOOL_OLS_LAYOUT";
        public const string GroupKey = Strings.Group;
        public const string InstallTypeKey = Strings.InstallType;
    }

    public static class TemplateBlocks
    {
        public const string Input = "\u041e\u041b\u0421_\u0412\u0412\u041e\u0414";
        public const string Breaker = "\u041e\u041b\u0421_\u0410\u0412\u0422\u041e\u041c\u0410\u0422";
        public const string Rcd = "\u041e\u041b\u0421_\u0423\u0417\u041e";
    }

    public static class PanelLayout
    {
        public const string MapFileName = "PanelLayoutMap.json";
        public const string DeviceTag = "\u0410\u041f\u041f\u0410\u0420\u0410\u0422";
        public const string ModulesTag = "\u041c\u041e\u0414\u0423\u041b\u0415\u0419";
        public const string GroupTag = "\u0413\u0420\u0423\u041f\u041f\u0410";
        public const string NoteTag = "\u041f\u0420\u0418\u041c\u0415\u0427\u0410\u041d\u0418\u0415";
    }
}
