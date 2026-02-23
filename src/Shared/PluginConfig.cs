// FILE: src/Shared/PluginConfig.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Keep canonical constants for attributes, commands, metadata keys, and template contracts.
//   SCOPE: Shared static values reused by command handlers and services.
//   DEPENDS: none
//   LINKS: M-CONFIG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   AttributeTags - Russian block attribute names.
//   Commands - AutoCAD command names exposed by plugin.
//   Metadata - XData and dictionary keys.
// END_MODULE_MAP

namespace ElTools.Shared;

public static class PluginConfig
{
    public static class AttributeTags
    {
        public const string Group = "ГРУППА";
        public const string Power = "МОЩНОСТЬ";
        public const string Voltage = "НАПРЯЖЕНИЕ";
        public const string Shield = "ЩИТ";
        public const string Phase = "ФАЗА";
        public const string Room = "ПОМЕЩЕНИЕ";
        public const string Note = "ПРИМЕЧАНИЕ";
    }

    public static class Commands
    {
        public const string Map = "EOM_MAP";
        public const string Trace = "EOM_TRACE";
        public const string Spec = "EOM_SPEC";
        public const string Update = "EOM_ОБНОВИТЬ";
        public const string ExportExcel = "EOM_ЭКСПОРТ_EXCEL";
        public const string ImportExcel = "EOM_ИМПОРТ_EXCEL";
        public const string BuildOls = "EOM_ПОСТРОИТЬ_ОЛС";
        public const string BuildPanelLayout = "EOM_КОМПОНОВКА_ЩИТА";
        public const string Validate = "EOM_ПРОВЕРКА";
        public const string ActiveGroup = "EOM_АКТИВНАЯ_ГРУППА";
        public const string AssignGroup = "EOM_НАЗНАЧИТЬ_ГРУППУ";
        public const string InstallTypeSettings = "EOM_НАСТРОЙКИ_ПРОКЛАДКИ";
    }

    public static class Metadata
    {
        public const string EomDataAppName = "EOM_DATA";
        public const string ElToolAppName = "ELTOOL";
        public const string GroupKey = "ГРУППА";
        public const string InstallTypeKey = "ТИП_ПРОКЛАДКИ";
    }
}
