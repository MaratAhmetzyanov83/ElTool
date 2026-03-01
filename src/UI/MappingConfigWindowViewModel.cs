// FILE: src/UI/MappingConfigWindowViewModel.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Provide Russian UI state and hints for mapping profile configuration.
//   SCOPE: Load, edit, validate, and save mapping rules from Settings.json.
//   DEPENDS: M-SETTINGS, M-LOGGING, M-MODELS
//   LINKS: M-MAP-CONFIG-VM, M-SETTINGS, M-LOGGING, M-MODELS
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   LoadProfile - Loads active profile and fills editable rules collection.
//   AddRule - Adds an empty mapping rule row.
//   RemoveSelectedRule - Removes selected rule row.
//   SaveProfile - Validates and persists profile to Settings.json.
//   MappingRuleItem - Editable mapping rule row model for UI binding.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Implemented Russian UI ViewModel for mapping configuration.
// END_CHANGE_SUMMARY

using System.Collections.ObjectModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElTools.Data;
using ElTools.Models;
using ElTools.Services;

namespace ElTools.UI;

public sealed partial class MappingConfigWindowViewModel : ObservableObject
{
    private readonly SettingsRepository _settingsRepository = new();
    private readonly LogService _log = new();
    private SettingsModel _settings = new();

    [ObservableProperty]
    private string _activeProfileName = "Default";

    [ObservableProperty]
    private string _statusMessage = "Р В Р’В Р Р†Р вЂљРЎС™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚Сћ Р В Р’В Р РЋРІР‚Сњ Р В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’ВµР В Р’В Р СћРІР‚ВР В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚СњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р РЋРІР‚ВР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В РІР‚в„– Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В» Р В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В Р РЏ.";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RemoveSelectedRuleCommand))]
    private MappingRuleItem? _selectedRule;

    public MappingConfigWindowViewModel()
    {
        LoadProfile();
    }

    public ObservableCollection<MappingRuleItem> Rules { get; } = new();

    public string WindowTitle => "Р В Р’В Р РЋРЎС™Р В Р’В Р вЂ™Р’В°Р В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р Р†РІР‚С›РІР‚вЂњР В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В° Р В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В Р РЏ Р В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СњР В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В ";
    public string ProfileNameLabel => "Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ°";
    public string RulesSectionLabel => "Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’В° Р В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В Р РЏ";
    public string SourceBlockColumn => "Р В Р’В Р вЂ™Р’ВР В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљР’В¦Р В Р’В Р РЋРІР‚СћР В Р’В Р СћРІР‚ВР В Р’В Р В РІР‚В¦Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р’В Р Р†РІР‚С›РІР‚вЂњ Р В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚Сњ";
    public string TargetBlockColumn => "Р В Р’В Р вЂ™Р’В¦Р В Р’В Р вЂ™Р’ВµР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚СћР В Р’В Р Р†РІР‚С›РІР‚вЂњ Р В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚Сњ";
    public string SourceHeightTagColumn => "Р В Р’В Р РЋРЎвЂєР В Р’В Р вЂ™Р’ВµР В Р’В Р РЋРІР‚вЂњ Р В Р’В Р В РІР‚В Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р Р†Р вЂљРІвЂћвЂ“ (Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљР Р‹Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р’В Р РЋРІР‚Сњ)";
    public string TargetHeightTagColumn => "Р В Р’В Р РЋРЎвЂєР В Р’В Р вЂ™Р’ВµР В Р’В Р РЋРІР‚вЂњ Р В Р’В Р В РІР‚В Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р Р†Р вЂљРІвЂћвЂ“ (Р В Р Р‹Р Р†Р вЂљР’В Р В Р’В Р вЂ™Р’ВµР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ°)";

    public string ProfileNameHint => "Р В Р’В Р В РІвЂљВ¬Р В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В°Р В Р’В Р вЂ™Р’В¶Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’Вµ Р В Р’В Р РЋРІР‚ВР В Р’В Р РЋР’ВР В Р Р‹Р В Р РЏ Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р РЏ. Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚ВР В Р’В Р РЋР’ВР В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™: Default, Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’ВµР В Р’В Р РЋРІР‚СњР В Р Р‹Р Р†Р вЂљРЎв„ў-Р В Р’В Р РЋРІР‚в„ў.";
    public string SourceBlockHint => "Р В Р’В Р вЂ™Р’ВР В Р’В Р РЋР’ВР В Р Р‹Р В Р РЏ Р В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В°, Р В Р’В Р РЋРІР‚СњР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р РЋРІР‚СћР В Р Р‹Р В РІР‚С™Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р’В Р Р†РІР‚С›РІР‚вЂњ Р В Р’В Р В РІР‚В¦Р В Р Р‹Р РЋРІР‚СљР В Р’В Р вЂ™Р’В¶Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚Сћ Р В Р’В Р вЂ™Р’В·Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋР’ВР В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В Р вЂ°.";
    public string TargetBlockHint => "Р В Р’В Р вЂ™Р’ВР В Р’В Р РЋР’ВР В Р Р‹Р В Р РЏ Р В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В° Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В· Р В Р’В Р РЋРІР‚СњР В Р’В Р РЋРІР‚СћР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚СћР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚вЂњР В Р’В Р РЋРІР‚Сћ Р В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦Р В Р’В Р СћРІР‚ВР В Р’В Р вЂ™Р’В°Р В Р Р‹Р В РІР‚С™Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В°.";
    public string HeightTagHint => "Р В Р’В Р РЋРІР‚С”Р В Р’В Р вЂ™Р’В±Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р Р‹Р Р†Р вЂљР Р‹Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚Сћ Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ°Р В Р’В Р вЂ™Р’В·Р В Р Р‹Р РЋРІР‚СљР В Р’В Р вЂ™Р’ВµР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РЎвЂњР В Р Р‹Р В Р РЏ Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’ВµР В Р’В Р РЋРІР‚вЂњ Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў.";
    public string SaveHint => "Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљР’В¦Р В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В Р вЂ° Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’В° Р В Р’В Р В РІР‚В  Settings.json Р В Р’В Р РЋРІР‚В Р В Р Р‹Р В РЎвЂњР В Р’В Р СћРІР‚ВР В Р’В Р вЂ™Р’ВµР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’В°Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В Р вЂ° Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ° Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚СњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В Р В Р’В Р В РІР‚В¦Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р’В Р РЋР’В.";
    public string ReloadHint => "Р В Р’В Р РЋРЎСџР В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р Р†Р вЂљР Р‹Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’В°Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В Р вЂ° Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ° Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В· Р В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р вЂ™Р’В°Р В Р’В Р Р†РІР‚С›РІР‚вЂњР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’В° Р В Р’В Р РЋРІР‚В Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р РЋР’ВР В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В Р вЂ° Р В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљР’В¦Р В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р В РІР‚В¦Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р’В Р вЂ™Р’Вµ Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В·Р В Р’В Р РЋР’ВР В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р В Р РЏ.";

    [RelayCommand]
    // START_CONTRACT: LoadProfile
    //   PURPOSE: Load profile.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: LoadProfile

    private void LoadProfile()
    {
        // START_BLOCK_VM_LOAD_PROFILE
        _settings = _settingsRepository.LoadSettings();
        string resolvedProfileName = string.IsNullOrWhiteSpace(_settings.ActiveProfile)
            ? "Default"
            : _settings.ActiveProfile.Trim();

        MappingProfile? active = _settings.MappingProfiles.FirstOrDefault(p => p.Name == resolvedProfileName);
        if (active is null)
        {
            active = new MappingProfile(resolvedProfileName, Array.Empty<MappingRule>());
            _settings.MappingProfiles.Add(active);
        }

        ActiveProfileName = active.Name;
        Rules.Clear();

        foreach (MappingRule rule in active.Rules)
        {
            Rules.Add(MappingRuleItem.FromRule(rule));
        }

        if (Rules.Count == 0)
        {
            Rules.Add(new MappingRuleItem());
        }

        StatusMessage = $"Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ° \"{ActiveProfileName}\" Р В Р’В Р вЂ™Р’В·Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂњР В Р Р‹Р В РІР‚С™Р В Р Р‹Р РЋРІР‚СљР В Р’В Р вЂ™Р’В¶Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦. Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»: {Rules.Count}.";
        _log.Write("[MappingConfigWindowViewModel][LoadProfile][VM_LOAD_PROFILE] Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ° Р В Р’В Р РЋР’ВР В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚вЂњР В Р’В Р вЂ™Р’В° Р В Р’В Р вЂ™Р’В·Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂњР В Р Р‹Р В РІР‚С™Р В Р Р‹Р РЋРІР‚СљР В Р’В Р вЂ™Р’В¶Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦.");
        // END_BLOCK_VM_LOAD_PROFILE
    }

    [RelayCommand]
    // START_CONTRACT: AddRule
    //   PURPOSE: Add rule.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: AddRule

    private void AddRule()
    {
        // START_BLOCK_VM_ADD_RULE
        var item = new MappingRuleItem();
        Rules.Add(item);
        SelectedRule = item;
        StatusMessage = "Р В Р’В Р Р†Р вЂљРЎСљР В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚Сћ Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚СћР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’Вµ Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚Сћ.";
        // END_BLOCK_VM_ADD_RULE
    }

    // START_CONTRACT: CanRemoveSelectedRule
    //   PURPOSE: Check whether remove selected rule.
    //   INPUTS: none
    //   OUTPUTS: { bool - true when method can check whether remove selected rule }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: CanRemoveSelectedRule

    private bool CanRemoveSelectedRule()
    {
        // START_BLOCK_VM_CAN_REMOVE_RULE
        return SelectedRule is not null;
        // END_BLOCK_VM_CAN_REMOVE_RULE
    }

    [RelayCommand(CanExecute = nameof(CanRemoveSelectedRule))]
    // START_CONTRACT: RemoveSelectedRule
    //   PURPOSE: Remove selected rule.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: RemoveSelectedRule

    private void RemoveSelectedRule()
    {
        // START_BLOCK_VM_REMOVE_RULE
        if (SelectedRule is null)
        {
            return;
        }

        Rules.Remove(SelectedRule);
        SelectedRule = null;

        if (Rules.Count == 0)
        {
            Rules.Add(new MappingRuleItem());
        }

        StatusMessage = "Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚Сћ Р В Р Р‹Р РЋРІР‚СљР В Р’В Р СћРІР‚ВР В Р’В Р вЂ™Р’В°Р В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚Сћ.";
        // END_BLOCK_VM_REMOVE_RULE
    }

    [RelayCommand]
    // START_CONTRACT: SaveProfile
    //   PURPOSE: Persist profile.
    //   INPUTS: none
    //   OUTPUTS: { void - no return value }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: SaveProfile

    private void SaveProfile()
    {
        // START_BLOCK_VM_SAVE_PROFILE
        if (!ValidateBeforeSave(out string validationError))
        {
            StatusMessage = validationError;
            _log.Write($"[MappingConfigWindowViewModel][SaveProfile][VM_SAVE_PROFILE] {validationError}");
            return;
        }

        string profileName = ActiveProfileName.Trim();
        var normalizedRules = Rules
            .Select(r => r.ToRule())
            .Where(r => !string.IsNullOrWhiteSpace(r.SourceBlockName) && !string.IsNullOrWhiteSpace(r.TargetBlockName))
            .ToList();

        _settings.ActiveProfile = profileName;

        int existingIndex = _settings.MappingProfiles.FindIndex(p => p.Name == profileName);
        var updatedProfile = new MappingProfile(profileName, normalizedRules);
        if (existingIndex >= 0)
        {
            _settings.MappingProfiles[existingIndex] = updatedProfile;
        }
        else
        {
            _settings.MappingProfiles.Add(updatedProfile);
        }

        _settingsRepository.SaveSettings(_settings);
        StatusMessage = $"Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ° \"{profileName}\" Р В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљР’В¦Р В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦. Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»: {normalizedRules.Count}.";
        _log.Write("[MappingConfigWindowViewModel][SaveProfile][VM_SAVE_PROFILE] Р В Р’В Р РЋРЎСџР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р вЂ° Р В Р’В Р РЋР’ВР В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚ВР В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚вЂњР В Р’В Р вЂ™Р’В° Р В Р Р‹Р В РЎвЂњР В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљР’В¦Р В Р Р‹Р В РІР‚С™Р В Р’В Р вЂ™Р’В°Р В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В¦.");
        // END_BLOCK_VM_SAVE_PROFILE
    }

    // START_CONTRACT: ValidateBeforeSave
    //   PURPOSE: Validate before save.
    //   INPUTS: { error: out string - method parameter }
    //   OUTPUTS: { bool - true when method can validate before save }
    //   SIDE_EFFECTS: Reads CAD/runtime/config state and may emit diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: ValidateBeforeSave

    private bool ValidateBeforeSave(out string error)
    {
        // START_BLOCK_VM_VALIDATE_BEFORE_SAVE
        if (string.IsNullOrWhiteSpace(ActiveProfileName))
        {
            error = "Р В Р’В Р вЂ™Р’ВР В Р’В Р РЋР’ВР В Р Р‹Р В Р РЏ Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р Р‹Р Р†Р вЂљРЎвЂєР В Р’В Р РЋРІР‚ВР В Р’В Р вЂ™Р’В»Р В Р Р‹Р В Р РЏ Р В Р’В Р В РІР‚В¦Р В Р’В Р вЂ™Р’Вµ Р В Р’В Р РЋР’ВР В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’В¶Р В Р’В Р вЂ™Р’ВµР В Р Р‹Р Р†Р вЂљРЎв„ў Р В Р’В Р вЂ™Р’В±Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В Р вЂ° Р В Р’В Р РЋРІР‚вЂќР В Р Р‹Р РЋРІР‚СљР В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р’В Р РЋР’В.";
            return false;
        }

        for (int i = 0; i < Rules.Count; i++)
        {
            MappingRuleItem rule = Rules[i];
            if (string.IsNullOrWhiteSpace(rule.SourceBlockName))
            {
                error = $"Р В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В° {i + 1}: Р В Р’В Р вЂ™Р’В·Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’В»Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’Вµ Р В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’Вµ \"Р В Р’В Р вЂ™Р’ВР В Р Р‹Р В РЎвЂњР В Р Р‹Р Р†Р вЂљР’В¦Р В Р’В Р РЋРІР‚СћР В Р’В Р СћРІР‚ВР В Р’В Р В РІР‚В¦Р В Р Р‹Р Р†Р вЂљРІвЂћвЂ“Р В Р’В Р Р†РІР‚С›РІР‚вЂњ Р В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚Сњ\".";
                return false;
            }

            if (string.IsNullOrWhiteSpace(rule.TargetBlockName))
            {
                error = $"Р В Р’В Р В Р вЂ№Р В Р Р‹Р Р†Р вЂљРЎв„ўР В Р Р‹Р В РІР‚С™Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚СњР В Р’В Р вЂ™Р’В° {i + 1}: Р В Р’В Р вЂ™Р’В·Р В Р’В Р вЂ™Р’В°Р В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’В»Р В Р’В Р В РІР‚В¦Р В Р’В Р РЋРІР‚ВР В Р Р‹Р Р†Р вЂљРЎв„ўР В Р’В Р вЂ™Р’Вµ Р В Р’В Р РЋРІР‚вЂќР В Р’В Р РЋРІР‚СћР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’Вµ \"Р В Р’В Р вЂ™Р’В¦Р В Р’В Р вЂ™Р’ВµР В Р’В Р вЂ™Р’В»Р В Р’В Р вЂ™Р’ВµР В Р’В Р В РІР‚В Р В Р’В Р РЋРІР‚СћР В Р’В Р Р†РІР‚С›РІР‚вЂњ Р В Р’В Р вЂ™Р’В±Р В Р’В Р вЂ™Р’В»Р В Р’В Р РЋРІР‚СћР В Р’В Р РЋРІР‚Сњ\".";
                return false;
            }
        }

        error = string.Empty;
        return true;
        // END_BLOCK_VM_VALIDATE_BEFORE_SAVE
    }
}

public sealed partial class MappingRuleItem : ObservableObject
{
    [ObservableProperty]
    private string _sourceBlockName = string.Empty;

    [ObservableProperty]
    private string _targetBlockName = string.Empty;

    [ObservableProperty]
    private string _heightSourceTag = "Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў";

    [ObservableProperty]
    private string _heightTargetTag = "Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў";

    // START_CONTRACT: FromRule
    //   PURPOSE: From rule.
    //   INPUTS: { rule: MappingRule - method parameter }
    //   OUTPUTS: { MappingRuleItem - result of from rule }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: FromRule

    public static MappingRuleItem FromRule(MappingRule rule)
    {
        // START_BLOCK_MAPPING_RULE_FROM_RULE
        return new MappingRuleItem
        {
            SourceBlockName = rule.SourceBlockName,
            TargetBlockName = rule.TargetBlockName,
            HeightSourceTag = string.IsNullOrWhiteSpace(rule.HeightSourceTag) ? "Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў" : rule.HeightSourceTag,
            HeightTargetTag = string.IsNullOrWhiteSpace(rule.HeightTargetTag) ? "Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў" : rule.HeightTargetTag
        };
        // END_BLOCK_MAPPING_RULE_FROM_RULE
    }

    // START_CONTRACT: ToRule
    //   PURPOSE: To rule.
    //   INPUTS: none
    //   OUTPUTS: { MappingRule - result of to rule }
    //   SIDE_EFFECTS: May read or update CAD/runtime/config state and diagnostics.
    //   LINKS: M-MAP-CONFIG-VM
    // END_CONTRACT: ToRule

    public MappingRule ToRule()
    {
        // START_BLOCK_MAPPING_RULE_TO_RULE
        return new MappingRule(
            SourceBlockName.Trim(),
            TargetBlockName.Trim(),
            string.IsNullOrWhiteSpace(HeightSourceTag) ? "Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў" : HeightSourceTag.Trim(),
            string.IsNullOrWhiteSpace(HeightTargetTag) ? "Р В Р’В Р Р†Р вЂљРІвЂћСћР В Р’В Р вЂ™Р’В«Р В Р’В Р В Р вЂ№Р В Р’В Р РЋРІР‚С”Р В Р’В Р РЋРЎвЂєР В Р’В Р РЋРІР‚в„ў" : HeightTargetTag.Trim());
        // END_BLOCK_MAPPING_RULE_TO_RULE
    }
}