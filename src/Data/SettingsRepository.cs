// FILE: src/Data/SettingsRepository.cs
// VERSION: 1.3.0
// START_MODULE_CONTRACT
//   PURPOSE: Load and save plugin runtime configuration files (settings, install-type rules, panel layout map).
//   SCOPE: Local JSON file persistence near ElTools.dll with defaults and normalization.
//   DEPENDS: M-JSON-ADAPTER, M-LOGGING, M-CONFIG
//   LINKS: M-SETTINGS, M-JSON-ADAPTER, M-LOGGING, M-CONFIG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   LoadSettings - Reads settings from Settings.json.
//   SaveSettings - Persists settings to Settings.json.
//   ResolveTemplatePath - Converts template path to absolute path near plugin DLL.
//   OpenInstallTypeConfig - Returns absolute path to install-type rule file.
//   LoadPanelLayoutMap - Loads selector and legacy OLS-to-layout mapping from plugin-local JSON.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using ElTools.Integrations;
using ElTools.Models;
using ElTools.Services;
using ElTools.Shared;
using System.Linq;

namespace ElTools.Data;

public class SettingsRepository
{
    private const string SettingsFileName = "Settings.json";
    private const string InstallTypeRulesFileName = "InstallTypeRules.json";
    private const string PanelLayoutMapFileName = PluginConfig.PanelLayout.MapFileName;
    private readonly JsonAdapter _json = new();
    private readonly LogService _log = new();

    // START_CONTRACT: LoadSettings
    //   PURPOSE: Load settings.
    //   INPUTS: none
    //   OUTPUTS: { SettingsModel - result of load settings }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: LoadSettings

    public SettingsModel LoadSettings()
    {
        // START_BLOCK_LOAD_SETTINGS
        string settingsPath = GetSettingsPath();
        if (!File.Exists(settingsPath))
        {
            _log.Write($"Р В¤Р В°Р в„–Р В» Р Р…Р В°РЎРѓРЎвЂљРЎР‚Р С•Р ВµР С” Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…: {settingsPath}. Р ВРЎРѓР С—Р С•Р В»РЎРЉР В·РЎС“РЎР‹РЎвЂљРЎРѓРЎРЏ Р Р…Р В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р С‘ Р С—Р С• РЎС“Р СР С•Р В»РЎвЂЎР В°Р Р…Р С‘РЎР‹.");
            return NormalizeSettings(new SettingsModel());
        }

        string raw = File.ReadAllText(settingsPath);
        SettingsModel parsed = _json.Deserialize<SettingsModel>(raw) ?? new SettingsModel();
        return NormalizeSettings(parsed);
        // END_BLOCK_LOAD_SETTINGS
    }

    // START_CONTRACT: SaveSettings
    //   PURPOSE: Persist settings.
    //   INPUTS: { model: SettingsModel - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: SaveSettings

    public void SaveSettings(SettingsModel model)
    {
        // START_BLOCK_SAVE_SETTINGS
        SettingsModel normalizedModel = NormalizeSettings(model);
        string settingsPath = GetSettingsPath();
        EnsureParentDirectory(settingsPath);
        string raw = _json.Serialize(normalizedModel);
        File.WriteAllText(settingsPath, raw);
        _log.Write($"Р СњР В°РЎРѓРЎвЂљРЎР‚Р С•Р в„–Р С”Р С‘ РЎРѓР С•РЎвЂ¦РЎР‚Р В°Р Р…Р ВµР Р…РЎвЂ№: {settingsPath}");
        // END_BLOCK_SAVE_SETTINGS
    }

    // START_CONTRACT: LoadInstallTypeRules
    //   PURPOSE: Load install type rules.
    //   INPUTS: none
    //   OUTPUTS: { InstallTypeRuleSet - result of load install type rules }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: LoadInstallTypeRules

    public InstallTypeRuleSet LoadInstallTypeRules()
    {
        // START_BLOCK_LOAD_INSTALL_TYPE_RULES
        string rulesPath = GetInstallTypeRulesPath();
        if (!File.Exists(rulesPath))
        {
            InstallTypeRuleSet defaults = new SettingsModel().InstallTypeRules;
            SaveInstallTypeRules(defaults);
            return defaults;
        }

        string raw = File.ReadAllText(rulesPath);
        InstallTypeRuleSet? parsed = _json.Deserialize<InstallTypeRuleSet>(raw);
        if (parsed is null || parsed.Rules.Count == 0)
        {
            _log.Write($"Р В¤Р В°Р в„–Р В» Р С—РЎР‚Р В°Р Р†Р С‘Р В» Р С—РЎР‚Р С•Р С”Р В»Р В°Р Т‘Р С”Р С‘ Р Р…Р ВµР С”Р С•РЎР‚РЎР‚Р ВµР С”РЎвЂљР ВµР Р…: {rulesPath}. Р СџРЎР‚Р С‘Р СР ВµР Р…Р ВµР Р…РЎвЂ№ Р В·Р Р…Р В°РЎвЂЎР ВµР Р…Р С‘РЎРЏ Р С—Р С• РЎС“Р СР С•Р В»РЎвЂЎР В°Р Р…Р С‘РЎР‹.");
            return new SettingsModel().InstallTypeRules;
        }

        return new InstallTypeRuleSet(parsed.Default, parsed.Rules.OrderBy(x => x.Priority).ToList());
        // END_BLOCK_LOAD_INSTALL_TYPE_RULES
    }

    // START_CONTRACT: SaveInstallTypeRules
    //   PURPOSE: Persist install type rules.
    //   INPUTS: { rules: InstallTypeRuleSet - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: SaveInstallTypeRules

    public void SaveInstallTypeRules(InstallTypeRuleSet rules)
    {
        // START_BLOCK_SAVE_INSTALL_TYPE_RULES
        string rulesPath = GetInstallTypeRulesPath();
        var normalized = new InstallTypeRuleSet(
            string.IsNullOrWhiteSpace(rules.Default) ? "Р СњР ВµР С•Р С—РЎР‚Р ВµР Т‘Р ВµР В»Р ВµР Р…Р С•" : rules.Default.Trim(),
            rules.Rules
                .OrderBy(x => x.Priority)
                .Select(x => new InstallTypeRule(
                    x.Priority,
                    x.MatchBy.Trim(),
                    x.Value.Trim(),
                    x.Result.Trim()))
                .ToList());
        EnsureParentDirectory(rulesPath);
        string raw = _json.Serialize(normalized);
        File.WriteAllText(rulesPath, raw);
        _log.Write($"Р СџРЎР‚Р В°Р Р†Р С‘Р В»Р В° РЎвЂљР С‘Р С—Р В° Р С—РЎР‚Р С•Р С”Р В»Р В°Р Т‘Р С”Р С‘ РЎРѓР С•РЎвЂ¦РЎР‚Р В°Р Р…Р ВµР Р…РЎвЂ№: {rulesPath}");
        // END_BLOCK_SAVE_INSTALL_TYPE_RULES
    }

    // START_CONTRACT: OpenInstallTypeConfig
    //   PURPOSE: Open install type config.
    //   INPUTS: none
    //   OUTPUTS: { string - textual result for open install type config }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: OpenInstallTypeConfig

    public string OpenInstallTypeConfig()
    {
        // START_BLOCK_OPEN_INSTALL_TYPE_CONFIG
        string rulesPath = GetInstallTypeRulesPath();
        if (!File.Exists(rulesPath))
        {
            SaveInstallTypeRules(new SettingsModel().InstallTypeRules);
        }

        return rulesPath;
        // END_BLOCK_OPEN_INSTALL_TYPE_CONFIG
    }

    // START_CONTRACT: LoadPanelLayoutMap
    //   PURPOSE: Load panel layout map.
    //   INPUTS: none
    //   OUTPUTS: { PanelLayoutMapConfig - result of load panel layout map }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: LoadPanelLayoutMap

    public PanelLayoutMapConfig LoadPanelLayoutMap()
    {
        // START_BLOCK_LOAD_PANEL_LAYOUT_MAP
        string mapPath = GetPanelLayoutMapPath();
        if (!File.Exists(mapPath))
        {
            PanelLayoutMapConfig defaults = CreateDefaultPanelLayoutMap();
            SavePanelLayoutMap(defaults);
            _log.Write($"Р В¤Р В°Р в„–Р В» Р С”Р В°РЎР‚РЎвЂљРЎвЂ№ Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘ Р Р…Р Вµ Р Р…Р В°Р в„–Р Т‘Р ВµР Р…. Р РЋР С•Р В·Р Т‘Р В°Р Р… РЎв‚¬Р В°Р В±Р В»Р С•Р Р…: {mapPath}");
            return defaults;
        }

        string raw = File.ReadAllText(mapPath);
        PanelLayoutMapConfig? parsed = _json.Deserialize<PanelLayoutMapConfig>(raw);
        if (parsed is null)
        {
            _log.Write($"Р В¤Р В°Р в„–Р В» Р С”Р В°РЎР‚РЎвЂљРЎвЂ№ Р С”Р С•Р СР С—Р С•Р Р…Р С•Р Р†Р С”Р С‘ Р Р…Р ВµР С”Р С•РЎР‚РЎР‚Р ВµР С”РЎвЂљР ВµР Р…: {mapPath}. Р СџРЎР‚Р С‘Р СР ВµР Р…Р ВµР Р…РЎвЂ№ Р В·Р Р…Р В°РЎвЂЎР ВµР Р…Р С‘РЎРЏ Р С—Р С• РЎС“Р СР С•Р В»РЎвЂЎР В°Р Р…Р С‘РЎР‹.");
            return CreateDefaultPanelLayoutMap();
        }

        return NormalizePanelLayoutMap(parsed);
        // END_BLOCK_LOAD_PANEL_LAYOUT_MAP
    }

    // START_CONTRACT: SavePanelLayoutMap
    //   PURPOSE: Persist panel layout map.
    //   INPUTS: { model: PanelLayoutMapConfig - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: SavePanelLayoutMap

    public void SavePanelLayoutMap(PanelLayoutMapConfig model)
    {
        // START_BLOCK_SAVE_PANEL_LAYOUT_MAP
        string mapPath = GetPanelLayoutMapPath();
        PanelLayoutMapConfig normalized = NormalizePanelLayoutMap(model);
        EnsureParentDirectory(mapPath);
        string raw = _json.Serialize(normalized);
        File.WriteAllText(mapPath, raw);
        // END_BLOCK_SAVE_PANEL_LAYOUT_MAP
    }

    // START_CONTRACT: ResolveTemplatePath
    //   PURPOSE: Resolve template path.
    //   INPUTS: { templatePath: string - method parameter }
    //   OUTPUTS: { string - textual result for resolve template path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: ResolveTemplatePath

    public string ResolveTemplatePath(string templatePath)
    {
        // START_BLOCK_RESOLVE_TEMPLATE_PATH
        string templateValue = string.IsNullOrWhiteSpace(templatePath)
            ? new SettingsModel().ExcelTemplatePath
            : templatePath.Trim();
        if (Path.IsPathRooted(templateValue))
        {
            return Path.GetFullPath(templateValue);
        }

        return Path.GetFullPath(Path.Combine(GetPluginDirectory(), templateValue));
        // END_BLOCK_RESOLVE_TEMPLATE_PATH
    }

    // START_CONTRACT: NormalizeSettings
    //   PURPOSE: Normalize settings.
    //   INPUTS: { model: SettingsModel - method parameter }
    //   OUTPUTS: { SettingsModel - result of normalize settings }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: NormalizeSettings

    private SettingsModel NormalizeSettings(SettingsModel model)
    {
        // START_BLOCK_NORMALIZE_SETTINGS
        model.ExcelTemplatePath = ResolveTemplatePath(model.ExcelTemplatePath);
        return model;
        // END_BLOCK_NORMALIZE_SETTINGS
    }

    // START_CONTRACT: CreateDefaultPanelLayoutMap
    //   PURPOSE: Create default panel layout map.
    //   INPUTS: none
    //   OUTPUTS: { PanelLayoutMapConfig - result of create default panel layout map }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: CreateDefaultPanelLayoutMap

    private static PanelLayoutMapConfig CreateDefaultPanelLayoutMap()
    {
        // START_BLOCK_CREATE_DEFAULT_PANEL_LAYOUT_MAP
        return new PanelLayoutMapConfig
        {
            Version = "2.0",
            DefaultModulesPerRow = 24,
            AttributeTags = new PanelLayoutAttributeTags
            {
                Device = PluginConfig.PanelLayout.DeviceTag,
                Modules = PluginConfig.PanelLayout.ModulesTag,
                Group = PluginConfig.PanelLayout.GroupTag,
                Note = PluginConfig.PanelLayout.NoteTag
            },
            SelectorRules = new List<PanelLayoutSelectorRule>(),
            LayoutMap = new List<PanelLayoutMapRule>
            {
                new("Р вЂ™Р вЂ™Р С›Р вЂќ", "\u041e\u041b\u0421_\u0412\u0412\u041e\u0414"),
                new("QS", "\u041e\u041b\u0421_\u0412\u0412\u041e\u0414"),
                new("QF", "\u041e\u041b\u0421_\u0410\u0412\u0422\u041e\u041c\u0410\u0422"),
                new("Р С’Р вЂ™Р СћР С›Р СљР С’Р Сћ", "\u041e\u041b\u0421_\u0410\u0412\u0422\u041e\u041c\u0410\u0422"),
                new("Р вЂ™Р С’", "\u041e\u041b\u0421_\u0410\u0412\u0422\u041e\u041c\u0410\u0422"),
                new("Р Р€Р вЂ”Р С›", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("RCD", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("Р С’Р вЂ™Р вЂќР Сћ", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("Р вЂќР ВР В¤", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("Р вЂќР ВР В¤Р С’Р вЂ™Р СћР С›Р СљР С’Р Сћ", "\u041e\u041b\u0421_\u0423\u0417\u041e")
            }
        };
        // END_BLOCK_CREATE_DEFAULT_PANEL_LAYOUT_MAP
    }

    // START_CONTRACT: NormalizePanelLayoutMap
    //   PURPOSE: Normalize panel layout map.
    //   INPUTS: { model: PanelLayoutMapConfig - method parameter }
    //   OUTPUTS: { PanelLayoutMapConfig - result of normalize panel layout map }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: NormalizePanelLayoutMap

    private PanelLayoutMapConfig NormalizePanelLayoutMap(PanelLayoutMapConfig model)
    {
        // START_BLOCK_NORMALIZE_PANEL_LAYOUT_MAP
        PanelLayoutMapConfig normalized = model ?? CreateDefaultPanelLayoutMap();
        if (normalized.DefaultModulesPerRow <= 0)
        {
            normalized.DefaultModulesPerRow = 24;
        }

        normalized.AttributeTags ??= new PanelLayoutAttributeTags();
        normalized.AttributeTags.Device = NormalizeTag(normalized.AttributeTags.Device, PluginConfig.PanelLayout.DeviceTag);
        normalized.AttributeTags.Modules = NormalizeTag(normalized.AttributeTags.Modules, PluginConfig.PanelLayout.ModulesTag);
        normalized.AttributeTags.Group = NormalizeTag(normalized.AttributeTags.Group, PluginConfig.PanelLayout.GroupTag);
        normalized.AttributeTags.Note = NormalizeTag(normalized.AttributeTags.Note, PluginConfig.PanelLayout.NoteTag);

        normalized.SelectorRules = normalized.SelectorRules
            .Where(x => x is not null)
            .Where(x => !string.IsNullOrWhiteSpace(x.SourceBlockName) && !string.IsNullOrWhiteSpace(x.LayoutBlockName))
            .Select(x => new PanelLayoutSelectorRule(
                x.Priority < 0 ? 0 : x.Priority,
                x.SourceBlockName.Trim(),
                NormalizeVisibilityValue(x.VisibilityValue),
                x.LayoutBlockName.Trim(),
                x.FallbackModules is > 0 ? x.FallbackModules : null))
            .OrderBy(x => x.Priority)
            .ThenBy(x => x.SourceBlockName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.VisibilityValue ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ToList();

        normalized.LayoutMap = normalized.LayoutMap
            .Where(x => x is not null)
            .Where(x => !string.IsNullOrWhiteSpace(x.DeviceKey) && !string.IsNullOrWhiteSpace(x.LayoutBlockName))
            .Select(x => new PanelLayoutMapRule(
                x.DeviceKey.Trim(),
                x.LayoutBlockName.Trim(),
                x.FallbackModules is > 0 ? x.FallbackModules : null))
            .ToList();

        normalized.Version = "2.0";
        return normalized;
        // END_BLOCK_NORMALIZE_PANEL_LAYOUT_MAP
    }

    // START_CONTRACT: GetSettingsPath
    //   PURPOSE: Retrieve settings path.
    //   INPUTS: none
    //   OUTPUTS: { string - textual result for retrieve settings path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: GetSettingsPath

    private static string GetSettingsPath()
    {
        // START_BLOCK_GET_SETTINGS_PATH
        return ResolveStoragePath(SettingsFileName);
        // END_BLOCK_GET_SETTINGS_PATH
    }

    // START_CONTRACT: GetInstallTypeRulesPath
    //   PURPOSE: Retrieve install type rules path.
    //   INPUTS: none
    //   OUTPUTS: { string - textual result for retrieve install type rules path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: GetInstallTypeRulesPath

    private static string GetInstallTypeRulesPath()
    {
        // START_BLOCK_GET_INSTALL_RULES_PATH
        return ResolveStoragePath(InstallTypeRulesFileName);
        // END_BLOCK_GET_INSTALL_RULES_PATH
    }

    // START_CONTRACT: GetPanelLayoutMapPath
    //   PURPOSE: Retrieve panel layout map path.
    //   INPUTS: none
    //   OUTPUTS: { string - textual result for retrieve panel layout map path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: GetPanelLayoutMapPath

    private static string GetPanelLayoutMapPath()
    {
        // START_BLOCK_GET_PANEL_LAYOUT_MAP_PATH
        return ResolveStoragePath(PanelLayoutMapFileName);
        // END_BLOCK_GET_PANEL_LAYOUT_MAP_PATH
    }

    // START_CONTRACT: ResolveStoragePath
    //   PURPOSE: Resolve storage path.
    //   INPUTS: { fileName: string - method parameter }
    //   OUTPUTS: { string - textual result for resolve storage path }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: ResolveStoragePath

    private static string ResolveStoragePath(string fileName)
    {
        // START_BLOCK_RESOLVE_STORAGE_PATH
        return Path.Combine(GetPluginDirectory(), fileName);
        // END_BLOCK_RESOLVE_STORAGE_PATH
    }

    // START_CONTRACT: GetPluginDirectory
    //   PURPOSE: Retrieve plugin directory.
    //   INPUTS: none
    //   OUTPUTS: { string - textual result for retrieve plugin directory }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: GetPluginDirectory

    private static string GetPluginDirectory()
    {
        // START_BLOCK_GET_PLUGIN_DIRECTORY
        string assemblyPath = typeof(SettingsRepository).Assembly.Location;
        if (!string.IsNullOrWhiteSpace(assemblyPath))
        {
            string? directory = Path.GetDirectoryName(assemblyPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                return Path.GetFullPath(directory);
            }
        }

        return Path.GetFullPath(AppContext.BaseDirectory);
        // END_BLOCK_GET_PLUGIN_DIRECTORY
    }

    // START_CONTRACT: EnsureParentDirectory
    //   PURPOSE: Ensure parent directory.
    //   INPUTS: { path: string - method parameter }
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: EnsureParentDirectory

    private static void EnsureParentDirectory(string path)
    {
        // START_BLOCK_ENSURE_PARENT_DIRECTORY
        string? directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }
        // END_BLOCK_ENSURE_PARENT_DIRECTORY
    }

    // START_CONTRACT: NormalizeTag
    //   PURPOSE: Normalize tag.
    //   INPUTS: { raw: string? - method parameter; fallback: string - method parameter }
    //   OUTPUTS: { string - textual result for normalize tag }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: NormalizeTag

    private static string NormalizeTag(string? raw, string fallback)
    {
        // START_BLOCK_NORMALIZE_TAG
        return string.IsNullOrWhiteSpace(raw) ? fallback : raw.Trim();
        // END_BLOCK_NORMALIZE_TAG
    }

    // START_CONTRACT: NormalizeVisibilityValue
    //   PURPOSE: Normalize visibility value.
    //   INPUTS: { raw: string? - method parameter }
    //   OUTPUTS: { string? - result of normalize visibility value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-SETTINGS
    // END_CONTRACT: NormalizeVisibilityValue

    private static string? NormalizeVisibilityValue(string? raw)
    {
        // START_BLOCK_NORMALIZE_VISIBILITY_VALUE
        return string.IsNullOrWhiteSpace(raw) ? null : raw.Trim();
        // END_BLOCK_NORMALIZE_VISIBILITY_VALUE
    }
}