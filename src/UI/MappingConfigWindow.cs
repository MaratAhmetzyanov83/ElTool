// FILE: src/UI/MappingConfigWindow.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Render Russian mapping configuration UI and bind it to MappingConfigWindowViewModel.
//   SCOPE: Build modal WPF window with profile editor, rules grid, commands, and status bar.
//   DEPENDS: M-MAP-CONFIG-VM
//   LINKS: M-MAP-CONFIG-UI, M-MAP-CONFIG-VM
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   BuildLayout - Creates window layout and data bindings.
//   CreateGrid - Creates editable rules table.
// END_MODULE_MAP

// START_CHANGE_SUMMARY
//   LAST_CHANGE: v1.0.0 - Added missing CHANGE_SUMMARY block for GRACE integrity refresh.
// END_CHANGE_SUMMARY


using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace ElTools.UI;

public sealed class MappingConfigWindow : Window
{
    public MappingConfigWindow(MappingConfigWindowViewModel viewModel)
    {
        // START_BLOCK_WINDOW_INIT
        DataContext = viewModel;
        Width = 1000;
        Height = 640;
        MinWidth = 900;
        MinHeight = 540;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
        ResizeMode = ResizeMode.CanResize;
        SetBinding(TitleProperty, new Binding(nameof(MappingConfigWindowViewModel.WindowTitle)));
        Content = BuildLayout();
        // END_BLOCK_WINDOW_INIT
    }

    // START_CONTRACT: BuildLayout
    //   PURPOSE: Build layout.
    //   INPUTS: none
    //   OUTPUTS: { UIElement - result of build layout }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-MAP-CONFIG-UI
    // END_CONTRACT: BuildLayout

    private UIElement BuildLayout()
    {
        // START_BLOCK_BUILD_LAYOUT
        var root = new Grid { Margin = new Thickness(12) };
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var profileRow = new Grid();
        profileRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        profileRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(260) });
        profileRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        profileRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        var profileLabel = new TextBlock
        {
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(0, 0, 8, 0)
        };
        profileLabel.SetBinding(TextBlock.TextProperty, new Binding(nameof(MappingConfigWindowViewModel.ProfileNameLabel)));
        Grid.SetColumn(profileLabel, 0);

        var profileTextBox = new TextBox { Margin = new Thickness(0, 0, 12, 0), VerticalAlignment = VerticalAlignment.Center };
        profileTextBox.SetBinding(TextBox.TextProperty, new Binding(nameof(MappingConfigWindowViewModel.ActiveProfileName))
        {
            Mode = BindingMode.TwoWay,
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
        });
        profileTextBox.SetBinding(ToolTipProperty, new Binding(nameof(MappingConfigWindowViewModel.ProfileNameHint)));
        Grid.SetColumn(profileTextBox, 1);

        var reloadButton = new Button
        {
            Content = "Р В РЎвЂєР В Р’В±Р В Р вЂ¦Р В РЎвЂўР В Р вЂ Р В РЎвЂР РЋРІР‚С™Р РЋР Р‰",
            Width = 110,
            Height = 30,
            Margin = new Thickness(0, 0, 8, 0)
        };
        reloadButton.SetBinding(Button.CommandProperty, new Binding(nameof(MappingConfigWindowViewModel.LoadProfileCommand)));
        reloadButton.SetBinding(ToolTipProperty, new Binding(nameof(MappingConfigWindowViewModel.ReloadHint)));
        Grid.SetColumn(reloadButton, 2);

        var saveButton = new Button
        {
            Content = "Р В Р Р‹Р В РЎвЂўР РЋРІР‚В¦Р РЋР вЂљР В Р’В°Р В Р вЂ¦Р В РЎвЂР РЋРІР‚С™Р РЋР Р‰",
            Width = 120,
            Height = 30
        };
        saveButton.SetBinding(Button.CommandProperty, new Binding(nameof(MappingConfigWindowViewModel.SaveProfileCommand)));
        saveButton.SetBinding(ToolTipProperty, new Binding(nameof(MappingConfigWindowViewModel.SaveHint)));
        Grid.SetColumn(saveButton, 3);

        profileRow.Children.Add(profileLabel);
        profileRow.Children.Add(profileTextBox);
        profileRow.Children.Add(reloadButton);
        profileRow.Children.Add(saveButton);
        Grid.SetRow(profileRow, 0);
        root.Children.Add(profileRow);

        var editorLayout = new Grid { Margin = new Thickness(0, 12, 0, 12) };
        editorLayout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
        editorLayout.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
        editorLayout.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

        var sectionTitle = new TextBlock { FontSize = 14, FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 8) };
        sectionTitle.SetBinding(TextBlock.TextProperty, new Binding(nameof(MappingConfigWindowViewModel.RulesSectionLabel)));
        Grid.SetRow(sectionTitle, 0);
        editorLayout.Children.Add(sectionTitle);

        DataGrid rulesGrid = CreateGrid();
        Grid.SetRow(rulesGrid, 1);
        editorLayout.Children.Add(rulesGrid);

        var actions = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Left,
            Margin = new Thickness(0, 8, 0, 0)
        };

        var addButton = new Button { Content = "Р В РІР‚СњР В РЎвЂўР В Р’В±Р В Р’В°Р В Р вЂ Р В РЎвЂР РЋРІР‚С™Р РЋР Р‰ Р В РЎвЂ”Р РЋР вЂљР В Р’В°Р В Р вЂ Р В РЎвЂР В Р’В»Р В РЎвЂў", Width = 150, Height = 28, Margin = new Thickness(0, 0, 8, 0) };
        addButton.SetBinding(Button.CommandProperty, new Binding(nameof(MappingConfigWindowViewModel.AddRuleCommand)));

        var removeButton = new Button { Content = "Р В Р в‚¬Р В РўвЂР В Р’В°Р В Р’В»Р В РЎвЂР РЋРІР‚С™Р РЋР Р‰ Р В Р вЂ Р РЋРІР‚в„–Р В Р’В±Р РЋР вЂљР В Р’В°Р В Р вЂ¦Р В Р вЂ¦Р В РЎвЂўР В Р’Вµ", Width = 150, Height = 28 };
        removeButton.SetBinding(Button.CommandProperty, new Binding(nameof(MappingConfigWindowViewModel.RemoveSelectedRuleCommand)));

        actions.Children.Add(addButton);
        actions.Children.Add(removeButton);
        Grid.SetRow(actions, 2);
        editorLayout.Children.Add(actions);

        Grid.SetRow(editorLayout, 1);
        root.Children.Add(editorLayout);

        var statusText = new TextBlock
        {
            Margin = new Thickness(0, 4, 0, 0),
            TextWrapping = TextWrapping.Wrap
        };
        statusText.SetBinding(TextBlock.TextProperty, new Binding(nameof(MappingConfigWindowViewModel.StatusMessage)));
        Grid.SetRow(statusText, 2);
        root.Children.Add(statusText);

        return root;
        // END_BLOCK_BUILD_LAYOUT
    }

    // START_CONTRACT: CreateGrid
    //   PURPOSE: Create grid.
    //   INPUTS: none
    //   OUTPUTS: { DataGrid - result of create grid }
    //   SIDE_EFFECTS: May modify CAD entities, configuration files, runtime state, or diagnostics.
    //   LINKS: M-MAP-CONFIG-UI
    // END_CONTRACT: CreateGrid

    private static DataGrid CreateGrid()
    {
        // START_BLOCK_CREATE_GRID
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

        grid.SetBinding(ItemsControl.ItemsSourceProperty, new Binding(nameof(MappingConfigWindowViewModel.Rules)));
        grid.SetBinding(DataGrid.SelectedItemProperty, new Binding(nameof(MappingConfigWindowViewModel.SelectedRule))
        {
            Mode = BindingMode.TwoWay
        });

        var sourceColumn = new DataGridTextColumn
        {
            Header = "Р В Р’ВР РЋР С“Р РЋРІР‚В¦Р В РЎвЂўР В РўвЂР В Р вЂ¦Р РЋРІР‚в„–Р В РІвЂћвЂ“ Р В Р’В±Р В Р’В»Р В РЎвЂўР В РЎвЂќ",
            Binding = new Binding(nameof(MappingRuleItem.SourceBlockName))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(1, DataGridLengthUnitType.Star)
        };

        var targetColumn = new DataGridTextColumn
        {
            Header = "Р В Р’В¦Р В Р’ВµР В Р’В»Р В Р’ВµР В Р вЂ Р В РЎвЂўР В РІвЂћвЂ“ Р В Р’В±Р В Р’В»Р В РЎвЂўР В РЎвЂќ",
            Binding = new Binding(nameof(MappingRuleItem.TargetBlockName))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(1, DataGridLengthUnitType.Star)
        };

        var sourceTagColumn = new DataGridTextColumn
        {
            Header = "Р В РЎС›Р В Р’ВµР В РЎвЂ“ Р В Р вЂ Р РЋРІР‚в„–Р РЋР С“Р В РЎвЂўР РЋРІР‚С™Р РЋРІР‚в„– (Р В РЎвЂР РЋР С“Р РЋРІР‚С™Р В РЎвЂўР РЋРІР‚РЋР В Р вЂ¦Р В РЎвЂР В РЎвЂќ)",
            Binding = new Binding(nameof(MappingRuleItem.HeightSourceTag))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(0.8, DataGridLengthUnitType.Star)
        };

        var targetTagColumn = new DataGridTextColumn
        {
            Header = "Р В РЎС›Р В Р’ВµР В РЎвЂ“ Р В Р вЂ Р РЋРІР‚в„–Р РЋР С“Р В РЎвЂўР РЋРІР‚С™Р РЋРІР‚в„– (Р РЋРІР‚В Р В Р’ВµР В Р’В»Р РЋР Р‰)",
            Binding = new Binding(nameof(MappingRuleItem.HeightTargetTag))
            {
                Mode = BindingMode.TwoWay,
                UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            },
            Width = new DataGridLength(0.8, DataGridLengthUnitType.Star)
        };

        grid.Columns.Add(sourceColumn);
        grid.Columns.Add(targetColumn);
        grid.Columns.Add(sourceTagColumn);
        grid.Columns.Add(targetTagColumn);
        return grid;
        // END_BLOCK_CREATE_GRID
    }
}