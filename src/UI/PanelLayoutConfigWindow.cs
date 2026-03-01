// FILE: src/UI/PanelLayoutConfigWindow.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Render panel layout configuration window with editable selector/legacy mappings.
//   SCOPE: Build WPF UI, bind to PanelLayoutConfigWindowViewModel, and run drawing pick actions from UI buttons.
//   DEPENDS: M-PANEL-LAYOUT-CONFIG-VM, M-CAD-CONTEXT
//   LINKS: M-PANEL-LAYOUT-CONFIG-UI, M-PANEL-LAYOUT-CONFIG-VM, M-CAD-CONTEXT
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   BuildLayout - Creates full window layout and binds controls to view-model.
//   CreateSelectorGrid - Creates editable selector rules table.
//   CreateLegacyGrid - Creates editable legacy layout map table.
// END_MODULE_MAP

using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ElTools.UI;

public sealed class PanelLayoutConfigWindow : Window
{
    private readonly PanelLayoutConfigWindowViewModel _viewModel;

    public PanelLayoutConfigWindow(PanelLayoutConfigWindowViewModel viewModel)
    {
        // START_BLOCK_PANEL_LAYOUT_WINDOW_INIT
        _viewModel = viewModel;
        DataContext = viewModel;
        Width = 1280;
        Height = 760;
        MinWidth = 1100;
        MinHeight = 640;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.CanResize;
        SetBinding(TitleProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.WindowTitle)));
        Content = BuildLayout();
        // END_BLOCK_PANEL_LAYOUT_WINDOW_INIT
    }

    private UIElement BuildLayout()
    {
        // START_BLOCK_PANEL_LAYOUT_BUILD_LAYOUT
        var root = new Grid { Margin = new Thickness(12) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var top = new Grid { Margin = new Thickness(0, 0, 0, 10) };
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(70) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(14) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(120) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(14) });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        top.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var modulesLabel = new TextBlock { Text = "Модулей в ряду:", VerticalAlignment = VerticalAlignment.Center };
        Grid.SetColumn(modulesLabel, 0);
        top.Children.Add(modulesLabel);

        var modulesBox = new TextBox { VerticalAlignment = VerticalAlignment.Center, Height = 26, Margin = new Thickness(6, 0, 0, 0) };
        modulesBox.SetBinding(TextBox.TextProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.DefaultModulesPerRow))
        {
            Mode = BindingMode.TwoWay,
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        });
        Grid.SetColumn(modulesBox, 1);
        top.Children.Add(modulesBox);

        var deviceLabel = new TextBlock { Text = "АППАРАТ:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0) };
        Grid.SetColumn(deviceLabel, 3);
        top.Children.Add(deviceLabel);
        var deviceBox = CreateTagTextBox(nameof(PanelLayoutConfigWindowViewModel.DeviceTag));
        Grid.SetColumn(deviceBox, 4);
        top.Children.Add(deviceBox);

        var modulesTagLabel = new TextBlock { Text = "МОДУЛЕЙ:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10, 0, 6, 0) };
        Grid.SetColumn(modulesTagLabel, 5);
        top.Children.Add(modulesTagLabel);
        var modulesTagBox = CreateTagTextBox(nameof(PanelLayoutConfigWindowViewModel.ModulesTag));
        Grid.SetColumn(modulesTagBox, 6);
        top.Children.Add(modulesTagBox);

        var groupTagLabel = new TextBlock { Text = "ГРУППА:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10, 0, 6, 0) };
        Grid.SetColumn(groupTagLabel, 7);
        top.Children.Add(groupTagLabel);
        var groupTagBox = CreateTagTextBox(nameof(PanelLayoutConfigWindowViewModel.GroupTag));
        Grid.SetColumn(groupTagBox, 8);
        top.Children.Add(groupTagBox);

        var noteTagLabel = new TextBlock { Text = "ПРИМЕЧАНИЕ:", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(10, 0, 6, 0) };
        Grid.SetColumn(noteTagLabel, 9);
        top.Children.Add(noteTagLabel);
        var noteTagBox = CreateTagTextBox(nameof(PanelLayoutConfigWindowViewModel.NoteTag));
        Grid.SetColumn(noteTagBox, 10);
        top.Children.Add(noteTagBox);

        var pickPairButton = new Button
        {
            Content = "Выбрать связь из чертежа",
            Height = 28,
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(10, 0, 10, 0)
        };
        pickPairButton.Click += (_, _) => ExecuteWithHiddenWindow(() => _viewModel.PickAndAddSelectorRuleCommand.Execute(null));
        Grid.SetColumn(pickPairButton, 12);
        top.Children.Add(pickPairButton);

        var reloadButton = new Button
        {
            Content = "Обновить",
            Height = 28,
            Margin = new Thickness(0, 0, 8, 0),
            Padding = new Thickness(10, 0, 10, 0)
        };
        reloadButton.SetBinding(Button.CommandProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.LoadMapCommand)));
        Grid.SetColumn(reloadButton, 13);
        top.Children.Add(reloadButton);

        var saveButton = new Button
        {
            Content = "Сохранить",
            Height = 28,
            Padding = new Thickness(10, 0, 10, 0)
        };
        saveButton.SetBinding(Button.CommandProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.SaveMapCommand)));
        Grid.SetColumn(saveButton, 14);
        top.Children.Add(saveButton);

        Grid.SetRow(top, 0);
        root.Children.Add(top);

        var selectorPanel = BuildSelectorPanel();
        Grid.SetRow(selectorPanel, 1);
        root.Children.Add(selectorPanel);

        var legacyPanel = BuildLegacyPanel();
        Grid.SetRow(legacyPanel, 2);
        root.Children.Add(legacyPanel);

        var statusText = new TextBlock
        {
            Margin = new Thickness(0, 8, 0, 0),
            TextWrapping = TextWrapping.Wrap
        };
        statusText.SetBinding(TextBlock.TextProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.StatusMessage)));
        Grid.SetRow(statusText, 3);
        root.Children.Add(statusText);

        return root;
        // END_BLOCK_PANEL_LAYOUT_BUILD_LAYOUT
    }

    private UIElement BuildSelectorPanel()
    {
        // START_BLOCK_PANEL_LAYOUT_BUILD_SELECTOR_PANEL
        var panel = new Grid { Margin = new Thickness(0, 0, 0, 8) };
        panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        panel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var title = new TextBlock
        {
            FontSize = 14,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 6)
        };
        title.SetBinding(TextBlock.TextProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.SelectorSectionTitle)));
        Grid.SetRow(title, 0);
        panel.Children.Add(title);

        DataGrid grid = CreateSelectorGrid();
        Grid.SetRow(grid, 1);
        panel.Children.Add(grid);

        var actions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(0, 8, 0, 0)
        };

        var addButton = new Button { Content = "Добавить", Width = 120, Height = 28, Margin = new Thickness(0, 0, 8, 0) };
        addButton.SetBinding(Button.CommandProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.AddSelectorRuleCommand)));

        var removeButton = new Button { Content = "Удалить", Width = 120, Height = 28, Margin = new Thickness(0, 0, 8, 0) };
        removeButton.SetBinding(Button.CommandProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.RemoveSelectedSelectorRuleCommand)));

        var pickSourceButton = new Button { Content = "Выбрать SOURCE", Width = 150, Height = 28, Margin = new Thickness(0, 0, 8, 0) };
        pickSourceButton.Click += (_, _) => ExecuteWithHiddenWindow(() => _viewModel.PickSourceForSelectedSelectorRuleCommand.Execute(null));

        var pickLayoutButton = new Button { Content = "Выбрать LAYOUT", Width = 150, Height = 28 };
        pickLayoutButton.Click += (_, _) => ExecuteWithHiddenWindow(() => _viewModel.PickLayoutForSelectedSelectorRuleCommand.Execute(null));

        actions.Children.Add(addButton);
        actions.Children.Add(removeButton);
        actions.Children.Add(pickSourceButton);
        actions.Children.Add(pickLayoutButton);
        Grid.SetRow(actions, 2);
        panel.Children.Add(actions);

        return panel;
        // END_BLOCK_PANEL_LAYOUT_BUILD_SELECTOR_PANEL
    }

    private UIElement BuildLegacyPanel()
    {
        // START_BLOCK_PANEL_LAYOUT_BUILD_LEGACY_PANEL
        var panel = new Grid();
        panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        panel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var title = new TextBlock
        {
            FontSize = 14,
            FontWeight = FontWeights.SemiBold,
            Margin = new Thickness(0, 0, 0, 6)
        };
        title.SetBinding(TextBlock.TextProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.LegacySectionTitle)));
        Grid.SetRow(title, 0);
        panel.Children.Add(title);

        DataGrid grid = CreateLegacyGrid();
        Grid.SetRow(grid, 1);
        panel.Children.Add(grid);

        var actions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            Margin = new Thickness(0, 8, 0, 0)
        };

        var addButton = new Button { Content = "Добавить", Width = 120, Height = 28, Margin = new Thickness(0, 0, 8, 0) };
        addButton.SetBinding(Button.CommandProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.AddLegacyRuleCommand)));

        var removeButton = new Button { Content = "Удалить", Width = 120, Height = 28, Margin = new Thickness(0, 0, 8, 0) };
        removeButton.SetBinding(Button.CommandProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.RemoveSelectedLegacyRuleCommand)));

        var pickLayoutButton = new Button { Content = "Выбрать LAYOUT", Width = 150, Height = 28 };
        pickLayoutButton.Click += (_, _) => ExecuteWithHiddenWindow(() => _viewModel.PickLayoutForSelectedLegacyRuleCommand.Execute(null));

        actions.Children.Add(addButton);
        actions.Children.Add(removeButton);
        actions.Children.Add(pickLayoutButton);
        Grid.SetRow(actions, 2);
        panel.Children.Add(actions);

        return panel;
        // END_BLOCK_PANEL_LAYOUT_BUILD_LEGACY_PANEL
    }

    private static DataGrid CreateSelectorGrid()
    {
        // START_BLOCK_PANEL_LAYOUT_CREATE_SELECTOR_GRID
        var grid = new DataGrid
        {
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            HeadersVisibility = DataGridHeadersVisibility.Column,
            IsReadOnly = false,
            SelectionMode = DataGridSelectionMode.Single,
            SelectionUnit = DataGridSelectionUnit.FullRow
        };

        grid.SetBinding(ItemsControl.ItemsSourceProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.SelectorRules)));
        grid.SetBinding(DataGrid.SelectedItemProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.SelectedSelectorRule))
        {
            Mode = BindingMode.TwoWay
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "Priority",
            Binding = new Binding(nameof(PanelLayoutSelectorRuleItem.Priority))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(0.5, DataGridLengthUnitType.Star)
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "SourceBlockName",
            Binding = new Binding(nameof(PanelLayoutSelectorRuleItem.SourceBlockName))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(1.4, DataGridLengthUnitType.Star)
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "VisibilityValue",
            Binding = new Binding(nameof(PanelLayoutSelectorRuleItem.VisibilityValue))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(1.0, DataGridLengthUnitType.Star)
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "LayoutBlockName",
            Binding = new Binding(nameof(PanelLayoutSelectorRuleItem.LayoutBlockName))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(1.4, DataGridLengthUnitType.Star)
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "FallbackModules",
            Binding = new Binding(nameof(PanelLayoutSelectorRuleItem.FallbackModules))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(0.7, DataGridLengthUnitType.Star)
        });

        return grid;
        // END_BLOCK_PANEL_LAYOUT_CREATE_SELECTOR_GRID
    }

    private static DataGrid CreateLegacyGrid()
    {
        // START_BLOCK_PANEL_LAYOUT_CREATE_LEGACY_GRID
        var grid = new DataGrid
        {
            AutoGenerateColumns = false,
            CanUserAddRows = false,
            CanUserDeleteRows = false,
            HeadersVisibility = DataGridHeadersVisibility.Column,
            IsReadOnly = false,
            SelectionMode = DataGridSelectionMode.Single,
            SelectionUnit = DataGridSelectionUnit.FullRow
        };

        grid.SetBinding(ItemsControl.ItemsSourceProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.LegacyRules)));
        grid.SetBinding(DataGrid.SelectedItemProperty, new Binding(nameof(PanelLayoutConfigWindowViewModel.SelectedLegacyRule))
        {
            Mode = BindingMode.TwoWay
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "DeviceKey",
            Binding = new Binding(nameof(PanelLayoutLegacyRuleItem.DeviceKey))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(1.2, DataGridLengthUnitType.Star)
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "LayoutBlockName",
            Binding = new Binding(nameof(PanelLayoutLegacyRuleItem.LayoutBlockName))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(1.6, DataGridLengthUnitType.Star)
        });

        grid.Columns.Add(new DataGridTextColumn
        {
            Header = "FallbackModules",
            Binding = new Binding(nameof(PanelLayoutLegacyRuleItem.FallbackModules))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(0.8, DataGridLengthUnitType.Star)
        });

        return grid;
        // END_BLOCK_PANEL_LAYOUT_CREATE_LEGACY_GRID
    }

    private static TextBox CreateTagTextBox(string propertyName)
    {
        // START_BLOCK_PANEL_LAYOUT_CREATE_TAG_TEXTBOX
        var textBox = new TextBox { Height = 26, VerticalAlignment = VerticalAlignment.Center };
        textBox.SetBinding(TextBox.TextProperty, new Binding(propertyName)
        {
            Mode = BindingMode.TwoWay,
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        });
        return textBox;
        // END_BLOCK_PANEL_LAYOUT_CREATE_TAG_TEXTBOX
    }

    private void ExecuteWithHiddenWindow(Action action)
    {
        // START_BLOCK_PANEL_LAYOUT_EXECUTE_WITH_HIDDEN_WINDOW
        Hide();
        try
        {
            action.Invoke();
        }
        finally
        {
            Show();
            Activate();
        }
        // END_BLOCK_PANEL_LAYOUT_EXECUTE_WITH_HIDDEN_WINDOW
    }
}
