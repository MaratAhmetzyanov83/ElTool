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

    public SettingsModel LoadSettings()
    {
        // START_BLOCK_LOAD_SETTINGS
        string settingsPath = GetSettingsPath();
        if (!File.Exists(settingsPath))
        {
            _log.Write($"Файл настроек не найден: {settingsPath}. Используются настройки по умолчанию.");
            return NormalizeSettings(new SettingsModel());
        }

        string raw = File.ReadAllText(settingsPath);
        SettingsModel parsed = _json.Deserialize<SettingsModel>(raw) ?? new SettingsModel();
        return NormalizeSettings(parsed);
        // END_BLOCK_LOAD_SETTINGS
    }

    public void SaveSettings(SettingsModel model)
    {
        // START_BLOCK_SAVE_SETTINGS
        SettingsModel normalizedModel = NormalizeSettings(model);
        string settingsPath = GetSettingsPath();
        EnsureParentDirectory(settingsPath);
        string raw = _json.Serialize(normalizedModel);
        File.WriteAllText(settingsPath, raw);
        _log.Write($"Настройки сохранены: {settingsPath}");
        // END_BLOCK_SAVE_SETTINGS
    }

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
            _log.Write($"Файл правил прокладки некорректен: {rulesPath}. Применены значения по умолчанию.");
            return new SettingsModel().InstallTypeRules;
        }

        return new InstallTypeRuleSet(parsed.Default, parsed.Rules.OrderBy(x => x.Priority).ToList());
        // END_BLOCK_LOAD_INSTALL_TYPE_RULES
    }

    public void SaveInstallTypeRules(InstallTypeRuleSet rules)
    {
        // START_BLOCK_SAVE_INSTALL_TYPE_RULES
        string rulesPath = GetInstallTypeRulesPath();
        var normalized = new InstallTypeRuleSet(
            string.IsNullOrWhiteSpace(rules.Default) ? "Неопределено" : rules.Default.Trim(),
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
        _log.Write($"Правила типа прокладки сохранены: {rulesPath}");
        // END_BLOCK_SAVE_INSTALL_TYPE_RULES
    }

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

    public PanelLayoutMapConfig LoadPanelLayoutMap()
    {
        // START_BLOCK_LOAD_PANEL_LAYOUT_MAP
        string mapPath = GetPanelLayoutMapPath();
        if (!File.Exists(mapPath))
        {
            PanelLayoutMapConfig defaults = CreateDefaultPanelLayoutMap();
            SavePanelLayoutMap(defaults);
            _log.Write($"Файл карты компоновки не найден. Создан шаблон: {mapPath}");
            return defaults;
        }

        string raw = File.ReadAllText(mapPath);
        PanelLayoutMapConfig? parsed = _json.Deserialize<PanelLayoutMapConfig>(raw);
        if (parsed is null)
        {
            _log.Write($"Файл карты компоновки некорректен: {mapPath}. Применены значения по умолчанию.");
            return CreateDefaultPanelLayoutMap();
        }

        return NormalizePanelLayoutMap(parsed);
        // END_BLOCK_LOAD_PANEL_LAYOUT_MAP
    }

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

    private SettingsModel NormalizeSettings(SettingsModel model)
    {
        // START_BLOCK_NORMALIZE_SETTINGS
        model.ExcelTemplatePath = ResolveTemplatePath(model.ExcelTemplatePath);
        return model;
        // END_BLOCK_NORMALIZE_SETTINGS
    }

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
                new("ВВОД", "\u041e\u041b\u0421_\u0412\u0412\u041e\u0414"),
                new("QS", "\u041e\u041b\u0421_\u0412\u0412\u041e\u0414"),
                new("QF", "\u041e\u041b\u0421_\u0410\u0412\u0422\u041e\u041c\u0410\u0422"),
                new("АВТОМАТ", "\u041e\u041b\u0421_\u0410\u0412\u0422\u041e\u041c\u0410\u0422"),
                new("ВА", "\u041e\u041b\u0421_\u0410\u0412\u0422\u041e\u041c\u0410\u0422"),
                new("УЗО", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("RCD", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("АВДТ", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("ДИФ", "\u041e\u041b\u0421_\u0423\u0417\u041e"),
                new("ДИФАВТОМАТ", "\u041e\u041b\u0421_\u0423\u0417\u041e")
            }
        };
        // END_BLOCK_CREATE_DEFAULT_PANEL_LAYOUT_MAP
    }

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

    private static string GetSettingsPath()
    {
        // START_BLOCK_GET_SETTINGS_PATH
        return ResolveStoragePath(SettingsFileName);
        // END_BLOCK_GET_SETTINGS_PATH
    }

    private static string GetInstallTypeRulesPath()
    {
        // START_BLOCK_GET_INSTALL_RULES_PATH
        return ResolveStoragePath(InstallTypeRulesFileName);
        // END_BLOCK_GET_INSTALL_RULES_PATH
    }

    private static string GetPanelLayoutMapPath()
    {
        // START_BLOCK_GET_PANEL_LAYOUT_MAP_PATH
        return ResolveStoragePath(PanelLayoutMapFileName);
        // END_BLOCK_GET_PANEL_LAYOUT_MAP_PATH
    }

    private static string ResolveStoragePath(string fileName)
    {
        // START_BLOCK_RESOLVE_STORAGE_PATH
        return Path.Combine(GetPluginDirectory(), fileName);
        // END_BLOCK_RESOLVE_STORAGE_PATH
    }

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

    private static string NormalizeTag(string? raw, string fallback)
    {
        // START_BLOCK_NORMALIZE_TAG
        return string.IsNullOrWhiteSpace(raw) ? fallback : raw.Trim();
        // END_BLOCK_NORMALIZE_TAG
    }

    private static string? NormalizeVisibilityValue(string? raw)
    {
        // START_BLOCK_NORMALIZE_VISIBILITY_VALUE
        return string.IsNullOrWhiteSpace(raw) ? null : raw.Trim();
        // END_BLOCK_NORMALIZE_VISIBILITY_VALUE
    }
}

