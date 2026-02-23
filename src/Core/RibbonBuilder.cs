// FILE: src/Core/RibbonBuilder.cs
// VERSION: 1.0.0
// START_MODULE_CONTRACT
//   PURPOSE: Create and manage Ribbon tab, panel, and command buttons.
//   SCOPE: Builds EOM PRO tab with command buttons.
//   DEPENDS: M-COMMANDS, M-LOG
//   LINKS: M-RIBBON, M-COMMANDS, M-LOG
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   BuildRibbon - Creates ribbon tab, panel, and buttons.
// END_MODULE_MAP

using System.Windows.Input;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using ElTools.Services;

namespace ElTools.Core;

public class RibbonBuilder
{
    private const string TabId = "ElTools.Tab";
    private readonly LogService _log = new();

    public void BuildRibbon()
    {
        // START_BLOCK_BUILD_RIBBON
        RibbonControl? ribbon = ComponentManager.Ribbon;
        if (ribbon is null)
        {
            _log.Write("Лента недоступна.");
            return;
        }

        RibbonTab? existing = ribbon.Tabs.FirstOrDefault(t => t.Id == TabId);
        if (existing is not null)
        {
            ribbon.Tabs.Remove(existing);
        }

        var tab = new RibbonTab { Id = TabId, Title = "ЭОМ ПРО" };
        var panelSource = new RibbonPanelSource { Title = "Инструменты" };
        panelSource.Items.Add(CreateButton("Замена блоков", "EOM_MAP ", "Замена блоков по профилю соответствия."));
        panelSource.Items.Add(CreateButton("Настройка маппинга", "EOM_MAPCFG ", "Открыть окно настройки правил соответствия блоков."));
        panelSource.Items.Add(CreateButton("Трассировка", "EOM_TRACE ", "Трассировка кабеля и расчет длины."));
        panelSource.Items.Add(CreateButton("Спецификация", "EOM_SPEC ", "Формирование спецификации по EOM_DATA."));

        tab.Panels.Add(new RibbonPanel { Source = panelSource });
        ribbon.Tabs.Add(tab);
        _log.Write("Лента ElTools создана.");
        // END_BLOCK_BUILD_RIBBON
    }

    private static RibbonButton CreateButton(string text, string command, string tooltip)
    {
        // START_BLOCK_CREATE_RIBBON_BUTTON
        return new RibbonButton
        {
            Text = text,
            ShowText = true,
            ToolTip = tooltip,
            CommandParameter = command,
            CommandHandler = new RibbonCommandHandler()
        };
        // END_BLOCK_CREATE_RIBBON_BUTTON
    }

    private sealed class RibbonCommandHandler : ICommand
    {
        public event EventHandler? CanExecuteChanged
        {
            add { }
            remove { }
        }

        public bool CanExecute(object? parameter)
        {
            // START_BLOCK_RIBBON_CAN_EXECUTE
            return Application.DocumentManager.MdiActiveDocument is not null;
            // END_BLOCK_RIBBON_CAN_EXECUTE
        }

        public void Execute(object? parameter)
        {
            // START_BLOCK_RIBBON_EXECUTE
            if (parameter is not string command)
            {
                return;
            }

            Document? doc = Application.DocumentManager.MdiActiveDocument;
            if (doc is null)
            {
                return;
            }

            doc.SendStringToExecute(command, true, false, true);
            // END_BLOCK_RIBBON_EXECUTE
        }
    }
}
