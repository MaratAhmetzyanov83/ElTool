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
    private string _statusMessage = "Готово к редактированию правил соответствия.";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RemoveSelectedRuleCommand))]
    private MappingRuleItem? _selectedRule;

    public MappingConfigWindowViewModel()
    {
        LoadProfile();
    }

    public ObservableCollection<MappingRuleItem> Rules { get; } = new();

    public string WindowTitle => "Настройка соответствия блоков";
    public string ProfileNameLabel => "Профиль";
    public string RulesSectionLabel => "Правила соответствия";
    public string SourceBlockColumn => "Исходный блок";
    public string TargetBlockColumn => "Целевой блок";
    public string SourceHeightTagColumn => "Тег высоты (источник)";
    public string TargetHeightTagColumn => "Тег высоты (цель)";

    public string ProfileNameHint => "Укажите имя профиля. Пример: Default, Проект-А.";
    public string SourceBlockHint => "Имя блока, который нужно заменить.";
    public string TargetBlockHint => "Имя блока из корпоративного стандарта.";
    public string HeightTagHint => "Обычно используется тег ВЫСОТА.";
    public string SaveHint => "Сохранить правила в Settings.json и сделать профиль активным.";
    public string ReloadHint => "Перечитать профиль из файла и отменить несохраненные изменения.";

    [RelayCommand]
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

        StatusMessage = $"Профиль \"{ActiveProfileName}\" загружен. Правил: {Rules.Count}.";
        _log.Write("[MappingConfigWindowViewModel][LoadProfile][VM_LOAD_PROFILE] Профиль маппинга загружен.");
        // END_BLOCK_VM_LOAD_PROFILE
    }

    [RelayCommand]
    private void AddRule()
    {
        // START_BLOCK_VM_ADD_RULE
        var item = new MappingRuleItem();
        Rules.Add(item);
        SelectedRule = item;
        StatusMessage = "Добавлено новое правило.";
        // END_BLOCK_VM_ADD_RULE
    }

    private bool CanRemoveSelectedRule()
    {
        // START_BLOCK_VM_CAN_REMOVE_RULE
        return SelectedRule is not null;
        // END_BLOCK_VM_CAN_REMOVE_RULE
    }

    [RelayCommand(CanExecute = nameof(CanRemoveSelectedRule))]
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

        StatusMessage = "Правило удалено.";
        // END_BLOCK_VM_REMOVE_RULE
    }

    [RelayCommand]
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
        StatusMessage = $"Профиль \"{profileName}\" сохранен. Правил: {normalizedRules.Count}.";
        _log.Write("[MappingConfigWindowViewModel][SaveProfile][VM_SAVE_PROFILE] Профиль маппинга сохранен.");
        // END_BLOCK_VM_SAVE_PROFILE
    }

    private bool ValidateBeforeSave(out string error)
    {
        // START_BLOCK_VM_VALIDATE_BEFORE_SAVE
        if (string.IsNullOrWhiteSpace(ActiveProfileName))
        {
            error = "Имя профиля не может быть пустым.";
            return false;
        }

        for (int i = 0; i < Rules.Count; i++)
        {
            MappingRuleItem rule = Rules[i];
            if (string.IsNullOrWhiteSpace(rule.SourceBlockName))
            {
                error = $"Строка {i + 1}: заполните поле \"Исходный блок\".";
                return false;
            }

            if (string.IsNullOrWhiteSpace(rule.TargetBlockName))
            {
                error = $"Строка {i + 1}: заполните поле \"Целевой блок\".";
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
    private string _heightSourceTag = "ВЫСОТА";

    [ObservableProperty]
    private string _heightTargetTag = "ВЫСОТА";

    public static MappingRuleItem FromRule(MappingRule rule)
    {
        // START_BLOCK_MAPPING_RULE_FROM_RULE
        return new MappingRuleItem
        {
            SourceBlockName = rule.SourceBlockName,
            TargetBlockName = rule.TargetBlockName,
            HeightSourceTag = string.IsNullOrWhiteSpace(rule.HeightSourceTag) ? "ВЫСОТА" : rule.HeightSourceTag,
            HeightTargetTag = string.IsNullOrWhiteSpace(rule.HeightTargetTag) ? "ВЫСОТА" : rule.HeightTargetTag
        };
        // END_BLOCK_MAPPING_RULE_FROM_RULE
    }

    public MappingRule ToRule()
    {
        // START_BLOCK_MAPPING_RULE_TO_RULE
        return new MappingRule(
            SourceBlockName.Trim(),
            TargetBlockName.Trim(),
            string.IsNullOrWhiteSpace(HeightSourceTag) ? "ВЫСОТА" : HeightSourceTag.Trim(),
            string.IsNullOrWhiteSpace(HeightTargetTag) ? "ВЫСОТА" : HeightTargetTag.Trim());
        // END_BLOCK_MAPPING_RULE_TO_RULE
    }
}

