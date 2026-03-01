// FILE: src/UI/PanelLayoutConfigWindowViewModel.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Provide editable UI state for panel layout mapping and selector rules.
//   SCOPE: Load, edit, validate, and save PanelLayoutMap.json; pick source/layout blocks from drawing via injected callbacks.
//   DEPENDS: M-SETTINGS, M-LOGGING, M-MODELS, M-CONFIG
//   LINKS: M-PANEL-LAYOUT-CONFIG-VM, M-SETTINGS, M-LOGGING, M-MODELS, M-CONFIG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   LoadMap - Loads panel layout map configuration and fills editable collections.
//   PickAndAddSelectorRule - Picks source and layout blocks from drawing and appends selector rule.
//   SaveMap - Validates and persists edited panel layout map.
//   PanelLayoutSelectorRuleItem - Editable selector rule row model for UI binding.
//   PanelLayoutLegacyRuleItem - Editable legacy fallback rule row model for UI binding.
// END_MODULE_MAP
//
// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added view-model for panel layout configuration UI and drawing-based selector rule picking.
// END_CHANGE_SUMMARY

using System.Collections.ObjectModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ElTools.Data;
using ElTools.Models;
using ElTools.Services;
using ElTools.Shared;

namespace ElTools.UI;

public sealed partial class PanelLayoutConfigWindowViewModel : ObservableObject
{
    private readonly SettingsRepository _settingsRepository = new();
    private readonly LogService _log = new();
    private readonly Func<OlsSourceSignature?> _pickSourceSignature;
    private readonly Func<string?> _pickLayoutBlockName;

    [ObservableProperty]
    private string _statusMessage = "Готово к настройке компоновки щита.";

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RemoveSelectedSelectorRuleCommand))]
    [NotifyCanExecuteChangedFor(nameof(PickSourceForSelectedSelectorRuleCommand))]
    [NotifyCanExecuteChangedFor(nameof(PickLayoutForSelectedSelectorRuleCommand))]
    private PanelLayoutSelectorRuleItem? _selectedSelectorRule;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RemoveSelectedLegacyRuleCommand))]
    [NotifyCanExecuteChangedFor(nameof(PickLayoutForSelectedLegacyRuleCommand))]
    private PanelLayoutLegacyRuleItem? _selectedLegacyRule;

    [ObservableProperty]
    private int _defaultModulesPerRow = 24;

    [ObservableProperty]
    private string _deviceTag = PluginConfig.PanelLayout.DeviceTag;

    [ObservableProperty]
    private string _modulesTag = PluginConfig.PanelLayout.ModulesTag;

    [ObservableProperty]
    private string _groupTag = PluginConfig.PanelLayout.GroupTag;

    [ObservableProperty]
    private string _noteTag = PluginConfig.PanelLayout.NoteTag;

    public PanelLayoutConfigWindowViewModel(
        Func<OlsSourceSignature?> pickSourceSignature,
        Func<string?> pickLayoutBlockName)
    {
        _pickSourceSignature = pickSourceSignature;
        _pickLayoutBlockName = pickLayoutBlockName;
        LoadMap();
    }

    public ObservableCollection<PanelLayoutSelectorRuleItem> SelectorRules { get; } = new();
    public ObservableCollection<PanelLayoutLegacyRuleItem> LegacyRules { get; } = new();

    public string WindowTitle => "Настройка компоновки щита";
    public string SelectorSectionTitle => "SelectorRules: SOURCE (блок + видимость) -> LAYOUT";
    public string LegacySectionTitle => "Legacy LayoutMap: АППАРАТ -> LAYOUT";
    public string AttributeTagsTitle => "Теги атрибутов ОЛС";

    [RelayCommand]
    private void LoadMap()
    {
        // START_BLOCK_VM_LOAD_MAP
        PanelLayoutMapConfig map = _settingsRepository.LoadPanelLayoutMap();

        DefaultModulesPerRow = map.DefaultModulesPerRow > 0 ? map.DefaultModulesPerRow : 24;
        DeviceTag = map.AttributeTags?.Device ?? PluginConfig.PanelLayout.DeviceTag;
        ModulesTag = map.AttributeTags?.Modules ?? PluginConfig.PanelLayout.ModulesTag;
        GroupTag = map.AttributeTags?.Group ?? PluginConfig.PanelLayout.GroupTag;
        NoteTag = map.AttributeTags?.Note ?? PluginConfig.PanelLayout.NoteTag;

        SelectorRules.Clear();
        foreach (PanelLayoutSelectorRule rule in map.SelectorRules)
        {
            SelectorRules.Add(PanelLayoutSelectorRuleItem.FromRule(rule));
        }

        LegacyRules.Clear();
        foreach (PanelLayoutMapRule rule in map.LayoutMap)
        {
            LegacyRules.Add(PanelLayoutLegacyRuleItem.FromRule(rule));
        }

        if (SelectorRules.Count == 0)
        {
            SelectorRules.Add(new PanelLayoutSelectorRuleItem());
        }

        if (LegacyRules.Count == 0)
        {
            LegacyRules.Add(new PanelLayoutLegacyRuleItem());
        }

        StatusMessage = $"Загружено правил: selector={SelectorRules.Count}, legacy={LegacyRules.Count}.";
        _log.Write("[PanelLayoutConfigWindowViewModel][LoadMap][VM_LOAD_MAP] Конфигурация компоновки загружена.");
        // END_BLOCK_VM_LOAD_MAP
    }

    [RelayCommand]
    private void AddSelectorRule()
    {
        // START_BLOCK_VM_ADD_SELECTOR_RULE
        var rule = new PanelLayoutSelectorRuleItem();
        SelectorRules.Add(rule);
        SelectedSelectorRule = rule;
        StatusMessage = "Добавлено новое selector-правило.";
        // END_BLOCK_VM_ADD_SELECTOR_RULE
    }

    private bool CanRemoveSelectedSelectorRule()
    {
        // START_BLOCK_VM_CAN_REMOVE_SELECTOR_RULE
        return SelectedSelectorRule is not null;
        // END_BLOCK_VM_CAN_REMOVE_SELECTOR_RULE
    }

    [RelayCommand(CanExecute = nameof(CanRemoveSelectedSelectorRule))]
    private void RemoveSelectedSelectorRule()
    {
        // START_BLOCK_VM_REMOVE_SELECTOR_RULE
        if (SelectedSelectorRule is null)
        {
            return;
        }

        SelectorRules.Remove(SelectedSelectorRule);
        SelectedSelectorRule = null;
        if (SelectorRules.Count == 0)
        {
            SelectorRules.Add(new PanelLayoutSelectorRuleItem());
        }

        StatusMessage = "Selector-правило удалено.";
        // END_BLOCK_VM_REMOVE_SELECTOR_RULE
    }

    private bool CanPickSelectorRule()
    {
        // START_BLOCK_VM_CAN_PICK_SELECTOR_RULE
        return SelectedSelectorRule is not null;
        // END_BLOCK_VM_CAN_PICK_SELECTOR_RULE
    }

    [RelayCommand(CanExecute = nameof(CanPickSelectorRule))]
    private void PickSourceForSelectedSelectorRule()
    {
        // START_BLOCK_VM_PICK_SOURCE_FOR_SELECTED_RULE
        if (SelectedSelectorRule is null)
        {
            return;
        }

        OlsSourceSignature? signature = _pickSourceSignature.Invoke();
        if (signature is null)
        {
            StatusMessage = "Выбор SOURCE отменен.";
            return;
        }

        SelectedSelectorRule.SourceBlockName = signature.SourceBlockName;
        SelectedSelectorRule.VisibilityValue = signature.VisibilityValue ?? string.Empty;
        StatusMessage = $"SOURCE установлен: {signature.SourceBlockName}|{(signature.VisibilityValue ?? "*")}.";
        // END_BLOCK_VM_PICK_SOURCE_FOR_SELECTED_RULE
    }

    [RelayCommand(CanExecute = nameof(CanPickSelectorRule))]
    private void PickLayoutForSelectedSelectorRule()
    {
        // START_BLOCK_VM_PICK_LAYOUT_FOR_SELECTED_RULE
        if (SelectedSelectorRule is null)
        {
            return;
        }

        string? layoutBlockName = _pickLayoutBlockName.Invoke();
        if (string.IsNullOrWhiteSpace(layoutBlockName))
        {
            StatusMessage = "Выбор LAYOUT отменен.";
            return;
        }

        SelectedSelectorRule.LayoutBlockName = layoutBlockName.Trim();
        StatusMessage = $"LAYOUT установлен: {SelectedSelectorRule.LayoutBlockName}.";
        // END_BLOCK_VM_PICK_LAYOUT_FOR_SELECTED_RULE
    }

    [RelayCommand]
    private void PickAndAddSelectorRule()
    {
        // START_BLOCK_VM_PICK_AND_ADD_SELECTOR_RULE
        OlsSourceSignature? signature = _pickSourceSignature.Invoke();
        if (signature is null)
        {
            StatusMessage = "Выбор SOURCE отменен.";
            return;
        }

        string? layoutBlockName = _pickLayoutBlockName.Invoke();
        if (string.IsNullOrWhiteSpace(layoutBlockName))
        {
            StatusMessage = "Выбор LAYOUT отменен.";
            return;
        }

        var rule = new PanelLayoutSelectorRuleItem
        {
            Priority = 100,
            SourceBlockName = signature.SourceBlockName,
            VisibilityValue = signature.VisibilityValue ?? string.Empty,
            LayoutBlockName = layoutBlockName.Trim(),
            FallbackModules = null
        };

        SelectorRules.Add(rule);
        SelectedSelectorRule = rule;
        StatusMessage = $"Добавлена связь: {rule.SourceBlockName}|{(string.IsNullOrWhiteSpace(rule.VisibilityValue) ? "*" : rule.VisibilityValue)} -> {rule.LayoutBlockName}.";
        // END_BLOCK_VM_PICK_AND_ADD_SELECTOR_RULE
    }

    [RelayCommand]
    private void AddLegacyRule()
    {
        // START_BLOCK_VM_ADD_LEGACY_RULE
        var rule = new PanelLayoutLegacyRuleItem();
        LegacyRules.Add(rule);
        SelectedLegacyRule = rule;
        StatusMessage = "Добавлено новое legacy-правило.";
        // END_BLOCK_VM_ADD_LEGACY_RULE
    }

    private bool CanRemoveSelectedLegacyRule()
    {
        // START_BLOCK_VM_CAN_REMOVE_LEGACY_RULE
        return SelectedLegacyRule is not null;
        // END_BLOCK_VM_CAN_REMOVE_LEGACY_RULE
    }

    [RelayCommand(CanExecute = nameof(CanRemoveSelectedLegacyRule))]
    private void RemoveSelectedLegacyRule()
    {
        // START_BLOCK_VM_REMOVE_LEGACY_RULE
        if (SelectedLegacyRule is null)
        {
            return;
        }

        LegacyRules.Remove(SelectedLegacyRule);
        SelectedLegacyRule = null;
        if (LegacyRules.Count == 0)
        {
            LegacyRules.Add(new PanelLayoutLegacyRuleItem());
        }

        StatusMessage = "Legacy-правило удалено.";
        // END_BLOCK_VM_REMOVE_LEGACY_RULE
    }

    private bool CanPickLegacyLayout()
    {
        // START_BLOCK_VM_CAN_PICK_LEGACY_LAYOUT
        return SelectedLegacyRule is not null;
        // END_BLOCK_VM_CAN_PICK_LEGACY_LAYOUT
    }

    [RelayCommand(CanExecute = nameof(CanPickLegacyLayout))]
    private void PickLayoutForSelectedLegacyRule()
    {
        // START_BLOCK_VM_PICK_LAYOUT_FOR_SELECTED_LEGACY_RULE
        if (SelectedLegacyRule is null)
        {
            return;
        }

        string? layoutBlockName = _pickLayoutBlockName.Invoke();
        if (string.IsNullOrWhiteSpace(layoutBlockName))
        {
            StatusMessage = "Выбор LAYOUT отменен.";
            return;
        }

        SelectedLegacyRule.LayoutBlockName = layoutBlockName.Trim();
        StatusMessage = $"LAYOUT установлен: {SelectedLegacyRule.LayoutBlockName}.";
        // END_BLOCK_VM_PICK_LAYOUT_FOR_SELECTED_LEGACY_RULE
    }

    [RelayCommand]
    private void SaveMap()
    {
        // START_BLOCK_VM_SAVE_MAP
        if (!ValidateBeforeSave(out string validationError))
        {
            StatusMessage = validationError;
            _log.Write($"[PanelLayoutConfigWindowViewModel][SaveMap][VM_SAVE_MAP] {validationError}");
            return;
        }

        int normalizedModulesPerRow = DefaultModulesPerRow is > 0 and <= 72 ? DefaultModulesPerRow : 24;
        var selectorRules = SelectorRules
            .Select(x => x.ToRule())
            .Where(x => !string.IsNullOrWhiteSpace(x.SourceBlockName) && !string.IsNullOrWhiteSpace(x.LayoutBlockName))
            .OrderBy(x => x.Priority)
            .ThenBy(x => x.SourceBlockName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(x => x.VisibilityValue ?? string.Empty, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var legacyRules = LegacyRules
            .Select(x => x.ToRule())
            .Where(x => !string.IsNullOrWhiteSpace(x.DeviceKey) && !string.IsNullOrWhiteSpace(x.LayoutBlockName))
            .ToList();

        var config = new PanelLayoutMapConfig
        {
            Version = "2.0",
            DefaultModulesPerRow = normalizedModulesPerRow,
            AttributeTags = new PanelLayoutAttributeTags
            {
                Device = NormalizeTag(DeviceTag, PluginConfig.PanelLayout.DeviceTag),
                Modules = NormalizeTag(ModulesTag, PluginConfig.PanelLayout.ModulesTag),
                Group = NormalizeTag(GroupTag, PluginConfig.PanelLayout.GroupTag),
                Note = NormalizeTag(NoteTag, PluginConfig.PanelLayout.NoteTag)
            },
            SelectorRules = selectorRules,
            LayoutMap = legacyRules
        };

        _settingsRepository.SavePanelLayoutMap(config);
        StatusMessage = $"Сохранено: selector={selectorRules.Count}, legacy={legacyRules.Count}, модулей в ряду={normalizedModulesPerRow}.";
        _log.Write("[PanelLayoutConfigWindowViewModel][SaveMap][VM_SAVE_MAP] Конфигурация компоновки сохранена.");
        // END_BLOCK_VM_SAVE_MAP
    }

    private bool ValidateBeforeSave(out string error)
    {
        // START_BLOCK_VM_VALIDATE_BEFORE_SAVE
        if (DefaultModulesPerRow <= 0 || DefaultModulesPerRow > 72)
        {
            error = "Количество модулей в ряду должно быть от 1 до 72.";
            return false;
        }

        for (int i = 0; i < SelectorRules.Count; i++)
        {
            PanelLayoutSelectorRuleItem rule = SelectorRules[i];
            bool hasAny = !string.IsNullOrWhiteSpace(rule.SourceBlockName)
                || !string.IsNullOrWhiteSpace(rule.LayoutBlockName)
                || !string.IsNullOrWhiteSpace(rule.VisibilityValue)
                || rule.FallbackModules.HasValue;
            if (!hasAny)
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(rule.SourceBlockName))
            {
                error = $"SelectorRules строка {i + 1}: заполните SourceBlockName.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(rule.LayoutBlockName))
            {
                error = $"SelectorRules строка {i + 1}: заполните LayoutBlockName.";
                return false;
            }

            if (rule.Priority < 0)
            {
                error = $"SelectorRules строка {i + 1}: Priority должен быть >= 0.";
                return false;
            }

            if (rule.FallbackModules is < 1 or > 72)
            {
                error = $"SelectorRules строка {i + 1}: FallbackModules должен быть в диапазоне 1..72.";
                return false;
            }
        }

        for (int i = 0; i < LegacyRules.Count; i++)
        {
            PanelLayoutLegacyRuleItem rule = LegacyRules[i];
            bool hasAny = !string.IsNullOrWhiteSpace(rule.DeviceKey)
                || !string.IsNullOrWhiteSpace(rule.LayoutBlockName)
                || rule.FallbackModules.HasValue;
            if (!hasAny)
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(rule.DeviceKey))
            {
                error = $"LayoutMap строка {i + 1}: заполните DeviceKey.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(rule.LayoutBlockName))
            {
                error = $"LayoutMap строка {i + 1}: заполните LayoutBlockName.";
                return false;
            }

            if (rule.FallbackModules is < 1 or > 72)
            {
                error = $"LayoutMap строка {i + 1}: FallbackModules должен быть в диапазоне 1..72.";
                return false;
            }
        }

        error = string.Empty;
        return true;
        // END_BLOCK_VM_VALIDATE_BEFORE_SAVE
    }

    private static string NormalizeTag(string? raw, string fallback)
    {
        // START_BLOCK_VM_NORMALIZE_TAG
        return string.IsNullOrWhiteSpace(raw) ? fallback : raw.Trim();
        // END_BLOCK_VM_NORMALIZE_TAG
    }
}

public sealed partial class PanelLayoutSelectorRuleItem : ObservableObject
{
    [ObservableProperty]
    private int _priority = 100;

    [ObservableProperty]
    private string _sourceBlockName = string.Empty;

    [ObservableProperty]
    private string _visibilityValue = string.Empty;

    [ObservableProperty]
    private string _layoutBlockName = string.Empty;

    [ObservableProperty]
    private int? _fallbackModules;

    public static PanelLayoutSelectorRuleItem FromRule(PanelLayoutSelectorRule rule)
    {
        // START_BLOCK_SELECTOR_ITEM_FROM_RULE
        return new PanelLayoutSelectorRuleItem
        {
            Priority = rule.Priority < 0 ? 0 : rule.Priority,
            SourceBlockName = rule.SourceBlockName,
            VisibilityValue = rule.VisibilityValue ?? string.Empty,
            LayoutBlockName = rule.LayoutBlockName,
            FallbackModules = rule.FallbackModules is > 0 and <= 72 ? rule.FallbackModules : null
        };
        // END_BLOCK_SELECTOR_ITEM_FROM_RULE
    }

    public PanelLayoutSelectorRule ToRule()
    {
        // START_BLOCK_SELECTOR_ITEM_TO_RULE
        int? normalizedFallback = FallbackModules is > 0 and <= 72 ? FallbackModules : null;
        string? normalizedVisibility = string.IsNullOrWhiteSpace(VisibilityValue) ? null : VisibilityValue.Trim();
        return new PanelLayoutSelectorRule(
            Priority < 0 ? 0 : Priority,
            SourceBlockName.Trim(),
            normalizedVisibility,
            LayoutBlockName.Trim(),
            normalizedFallback);
        // END_BLOCK_SELECTOR_ITEM_TO_RULE
    }
}

public sealed partial class PanelLayoutLegacyRuleItem : ObservableObject
{
    [ObservableProperty]
    private string _deviceKey = string.Empty;

    [ObservableProperty]
    private string _layoutBlockName = string.Empty;

    [ObservableProperty]
    private int? _fallbackModules;

    public static PanelLayoutLegacyRuleItem FromRule(PanelLayoutMapRule rule)
    {
        // START_BLOCK_LEGACY_ITEM_FROM_RULE
        return new PanelLayoutLegacyRuleItem
        {
            DeviceKey = rule.DeviceKey,
            LayoutBlockName = rule.LayoutBlockName,
            FallbackModules = rule.FallbackModules is > 0 and <= 72 ? rule.FallbackModules : null
        };
        // END_BLOCK_LEGACY_ITEM_FROM_RULE
    }

    public PanelLayoutMapRule ToRule()
    {
        // START_BLOCK_LEGACY_ITEM_TO_RULE
        int? normalizedFallback = FallbackModules is > 0 and <= 72 ? FallbackModules : null;
        return new PanelLayoutMapRule(DeviceKey.Trim(), LayoutBlockName.Trim(), normalizedFallback);
        // END_BLOCK_LEGACY_ITEM_TO_RULE
    }
}
