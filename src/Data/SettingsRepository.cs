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

namespace ElTools.Data;

public class SettingsRepository
{
    private const string FileName = "Settings.json";
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
}
