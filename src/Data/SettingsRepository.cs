// FILE: src/Data/SettingsRepository.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Load and save Settings.json and plugin profiles.
//   SCOPE: Local JSON file persistence for settings.
//   DEPENDS: M-JSON, M-LOG
//   LINKS: M-SETTINGS, M-JSON, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   LoadSettings - Reads settings from Settings.json.
//   SaveSettings - Persists settings to Settings.json.
// END_MODULE_MAP

using ElTools.Integrations;
using ElTools.Models;
using ElTools.Services;
using System.Linq;

namespace ElTools.Data;

public class SettingsRepository
{
    private const string FileName = "Settings.json";
    private const string InstallTypeRulesFileName = "InstallTypeRules.json";
    private readonly JsonAdapter _json = new();
    private readonly LogService _log = new();

    public SettingsModel LoadSettings()
    {
        // START_BLOCK_LOAD_SETTINGS
        if (!File.Exists(FileName))
        {
            _log.Write("Файл Settings.json не найден, используются настройки по умолчанию.");
            return new SettingsModel();
        }

        string raw = File.ReadAllText(FileName);
        return _json.Deserialize<SettingsModel>(raw) ?? new SettingsModel();
        // END_BLOCK_LOAD_SETTINGS
    }

    public void SaveSettings(SettingsModel model)
    {
        // START_BLOCK_SAVE_SETTINGS
        string raw = _json.Serialize(model);
        File.WriteAllText(FileName, raw);
        _log.Write("Настройки сохранены в Settings.json.");
        // END_BLOCK_SAVE_SETTINGS
    }

    public InstallTypeRuleSet LoadInstallTypeRules()
    {
        // START_BLOCK_LOAD_INSTALL_TYPE_RULES
        if (!File.Exists(InstallTypeRulesFileName))
        {
            InstallTypeRuleSet defaults = new SettingsModel().InstallTypeRules;
            SaveInstallTypeRules(defaults);
            return defaults;
        }

        string raw = File.ReadAllText(InstallTypeRulesFileName);
        InstallTypeRuleSet? parsed = _json.Deserialize<InstallTypeRuleSet>(raw);
        if (parsed is null || parsed.Rules.Count == 0)
        {
            _log.Write("Файл правил прокладки некорректен, применены значения по умолчанию.");
            return new SettingsModel().InstallTypeRules;
        }

        return new InstallTypeRuleSet(parsed.Default, parsed.Rules.OrderBy(x => x.Priority).ToList());
        // END_BLOCK_LOAD_INSTALL_TYPE_RULES
    }

    public void SaveInstallTypeRules(InstallTypeRuleSet rules)
    {
        // START_BLOCK_SAVE_INSTALL_TYPE_RULES
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
        string raw = _json.Serialize(normalized);
        File.WriteAllText(InstallTypeRulesFileName, raw);
        _log.Write("Правила типа прокладки сохранены.");
        // END_BLOCK_SAVE_INSTALL_TYPE_RULES
    }

    public string OpenInstallTypeConfig()
    {
        // START_BLOCK_OPEN_INSTALL_TYPE_CONFIG
        if (!File.Exists(InstallTypeRulesFileName))
        {
            SaveInstallTypeRules(new SettingsModel().InstallTypeRules);
        }

        return Path.GetFullPath(InstallTypeRulesFileName);
        // END_BLOCK_OPEN_INSTALL_TYPE_CONFIG
    }
}
