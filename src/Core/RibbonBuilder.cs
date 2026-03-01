// FILE: src/Core/RibbonBuilder.cs
// VERSION: 1.2.0
// START_MODULE_CONTRACT
//   PURPOSE: Create and manage Ribbon tab, panel, and command buttons.
//   SCOPE: Builds EOM PRO tab with command buttons.
//   DEPENDS: M-ENTRY-COMMANDS, M-CONFIG, M-LOGGING
//   LINKS: M-RIBBON, M-ENTRY-COMMANDS, M-CONFIG, M-LOGGING
// END_MODULE_CONTRACT
//
// START_MODULE_MAP
//   BuildRibbon - Creates ribbon tab, panel, and buttons.
//   BuildBasePanel - Adds core operational commands.
//   BuildExcelPanel - Adds Excel import/export and drawing commands.
//   BuildControlPanel - Adds validation and configuration commands.
// END_MODULE_MAP

using System.Windows.Input;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.Windows;
using ElTools.Services;
using ElTools.Shared;

namespace ElTools.Core;

public class RibbonBuilder
{
    private const string TabId = "ElTools.Tab";
    private readonly LogService _log = new();

    public bool BuildRibbon()
    {
        // START_BLOCK_BUILD_RIBBON
        RibbonControl? ribbon = ComponentManager.Ribbon;
        if (ribbon is null)
        {
            _log.Write("Лента недоступна.");
            return false;
        }

        RibbonTab? existing = ribbon.Tabs.FirstOrDefault(t => t.Id == TabId);
        if (existing is not null)
        {
            ribbon.Tabs.Remove(existing);
        }

        var tab = new RibbonTab { Id = TabId, Title = "ЭОМ ПРО" };
        tab.Panels.Add(BuildBasePanel());
        tab.Panels.Add(BuildExcelPanel());
        tab.Panels.Add(BuildControlPanel());
        ribbon.Tabs.Add(tab);
        _log.Write("Лента ElTools создана.");
        return true;
        // END_BLOCK_BUILD_RIBBON
    }

    private static RibbonPanel BuildBasePanel()
    {
        // START_BLOCK_BUILD_BASE_PANEL
        var panelSource = new RibbonPanelSource { Title = "Базовые" };
        panelSource.Items.Add(CreateButton(
            "Замена блоков",
            PluginConfig.Commands.Map,
            "Назначение: заменяет блоки по выбранному образцу.\nТребуется: исходный и целевой блок на чертеже.\nРезультат: массовая замена с сохранением позиции/поворота."));
        panelSource.Items.Add(CreateButton(
            "Трассировка",
            PluginConfig.Commands.Trace,
            "Назначение: строит ответвления от магистрали и считает длины.\nТребуется: базовая полилиния и блоки-нагрузки.\nРезультат: линии трассы и расчет длины кабеля."));
        panelSource.Items.Add(CreateButton(
            "Спецификация",
            PluginConfig.Commands.Spec,
            "Назначение: формирует ведомость и экспортные файлы.\nТребуется: заполненные атрибуты и трассы/группы.\nРезультат: таблица AutoCAD, CSV и Excel INPUT."));
        panelSource.Items.Add(CreateButton(
            "Обновить расчеты",
            PluginConfig.Commands.Update,
            "Назначение: пересчитывает агрегаты по группам.\nТребуется: линии и нагрузки с назначенными группами.\nРезультат: актуальные суммы длин и мощностей."));
        return new RibbonPanel { Source = panelSource };
        // END_BLOCK_BUILD_BASE_PANEL
    }

    private static RibbonPanel BuildExcelPanel()
    {
        // START_BLOCK_BUILD_EXCEL_PANEL
        var panelSource = new RibbonPanelSource { Title = "Excel/Схемы" };
        panelSource.Items.Add(CreateButton(
            "Экспорт Excel",
            PluginConfig.Commands.ExportExcel,
            "Назначение: выгружает INPUT для расчета.\nТребуется: путь к шаблону в настройках.\nРезультат: файл *.INPUT.csv рядом с шаблоном."));
        panelSource.Items.Add(CreateButton(
            "Импорт Excel",
            PluginConfig.Commands.ImportExcel,
            "Назначение: загружает результаты расчета.\nТребуется: файл *.OUTPUT.csv рядом с шаблоном.\nРезультат: кэш OUTPUT и отчет CSV."));
        panelSource.Items.Add(CreateButton(
            "Построить ОЛС",
            PluginConfig.Commands.BuildOls,
            "Назначение: строит однолинейную схему.\nТребуется: импортированный OUTPUT и точка вставки.\nРезультат: блоки/подписи ОЛС по строкам OUTPUT."));
        panelSource.Items.Add(CreateButton(
            "Связать визуализацию",
            PluginConfig.Commands.BindPanelLayoutVisualization,
            "Назначение: создает правило SOURCE(блок+видимость) -> блок визуализации.\nТребуется: выбрать исходный блок ОЛС и блок визуализации.\nРезультат: сохранение SelectorRules в PanelLayoutMap.json."));
        panelSource.Items.Add(CreateButton(
            "Настройка компоновки",
            PluginConfig.Commands.PanelLayoutConfig,
            "Назначение: открывает UI настройки компоновки щита.\nТребуется: доступ к чертежу и PanelLayoutMap.json.\nРезультат: редактирование SelectorRules/legacy LayoutMap, выбор связей SOURCE->LAYOUT из UI."));
        panelSource.Items.Add(CreateButton(
            "Компоновка щита",
            PluginConfig.Commands.BuildPanelLayout,
            "Назначение: раскладывает аппараты по модулям из выделенной ОЛС.\nТребуется: правила SelectorRules (или legacy LayoutMap) и МОДУЛЕЙ/FallbackModules.\nРезультат: графическая компоновка щита по DIN-рядам."));
        return new RibbonPanel { Source = panelSource };
        // END_BLOCK_BUILD_EXCEL_PANEL
    }

    private static RibbonPanel BuildControlPanel()
    {
        // START_BLOCK_BUILD_CONTROL_PANEL
        var panelSource = new RibbonPanelSource { Title = "Контроль/Настройки" };
        panelSource.Items.Add(CreateButton(
            "Проверка",
            PluginConfig.Commands.Validate,
            "Назначение: проверяет данные проекта на ошибки.\nТребуется: линии/нагрузки в модели.\nРезультат: список проблем с переходом к объекту."));
        panelSource.Items.Add(CreateButton(
            "Активная группа",
            PluginConfig.Commands.ActiveGroup,
            "Назначение: задает группу для новых линий.\nТребуется: код группы.\nРезультат: автоназначение группы при создании line/polyline."));
        panelSource.Items.Add(CreateButton(
            "Назначить группу",
            PluginConfig.Commands.AssignGroup,
            "Назначение: массово пишет группу в выбранные линии.\nТребуется: код группы и выбор LINE/LWPOLYLINE.\nРезультат: XData группы для выделенных объектов."));
        panelSource.Items.Add(CreateButton(
            "Типы прокладки",
            PluginConfig.Commands.InstallTypeSettings,
            "Назначение: открывает конфиг правил прокладки.\nТребуется: доступ к файлу настроек.\nРезультат: путь к InstallTypeRules.json в командной строке."));
        panelSource.Items.Add(CreateButton(
            "Настройка замен",
            PluginConfig.Commands.MapConfig,
            "Назначение: открывает окно настроек соответствий блоков.\nТребуется: доступ к UI плагина.\nРезультат: редактирование профилей замены."));
        return new RibbonPanel { Source = panelSource };
        // END_BLOCK_BUILD_CONTROL_PANEL
    }

    private static RibbonButton CreateButton(string text, string command, string tooltip)
    {
        // START_BLOCK_CREATE_RIBBON_BUTTON
        return new RibbonButton
        {
            Text = text,
            ShowText = true,
            ToolTip = tooltip,
            CommandParameter = CreateCommandMacro(command),
            CommandHandler = new RibbonCommandHandler()
        };
        // END_BLOCK_CREATE_RIBBON_BUTTON
    }

    private static string CreateCommandMacro(string command)
    {
        // START_BLOCK_CREATE_COMMAND_MACRO
        return $"_.{command} ";
        // END_BLOCK_CREATE_COMMAND_MACRO
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
            return true;
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
            if (doc is not null)
            {
                string macro = command.EndsWith(" ", StringComparison.Ordinal) ? command : command + " ";
                doc.SendStringToExecute(macro, true, false, false);
            }
            // END_BLOCK_RIBBON_EXECUTE
        }
    }
}

