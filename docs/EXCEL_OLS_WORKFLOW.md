# Excel OLS Workflow Reference

Updated: 2026-03-01

## Sample Workbook (User Reference)

Primary sample file for workflow understanding:

`C:\Users\marat\Desktop\Проект\04 Проект\№ 25-ЮН-ИНЖ 20ПИР-ЭОМ\Однолинейка ЩР ЭОМ.xlsx`

Target worksheet for direct import fallback:

`В Акад`

## Operational Flow

1. `EOM_ЭКСПОРТ_EXCEL` writes `<template>.INPUT.csv`.
2. External Excel calculation updates workbook.
3. `EOM_ИМПОРТ_EXCEL` reads output rows and refreshes cache.
4. `EOM_ПОСТРОИТЬ_ОЛС` builds drawing from imported rows.

## Direct Workbook Fallback Mapping (`В Акад`)

Data rows start from Excel row `3`.

Column mapping used by `CsvExcelGateway.ReadOutputRowsFromWorkbook`:

- `D` (4): `shield`
- `E` (5): `group` / line id
- `G` (7): `note` (consumer name/description)
- `K` (11): `circuitBreaker`
- `M` (13): `cable`

Derived/default values when building `ExcelOutputRow`:

- `phase`: `string.Empty`
- `cablesCount`: `1`
- `length`: `0`
- `circuitBreaker`: defaults to `QF` if column `K` is empty

Row is skipped when:

- any of `shield`, `group`, `cable` is empty or non-meaningful;
- `group == "0"`;
- value looks like formula/error/header marker (`#...`, `=...`, known header captions).

## Real Row Cases (From Sample)

Accepted row example:

- Row `15`: `D=ЩР`, `E=R.1.01.1`, `K=R.1.01.1`, `M=ВВГнг(А)-LS (3х2,5) ...`.
- Result: imported as one OLS row.

Skipped row examples:

- Row `3`: `E=0`, `M` empty. Result: skipped (`group == "0"` and missing cable).
- Row `8`: `E=KV-1`, `M=#N/A`. Result: skipped (`#...` marker is treated as non-meaningful).

Header/formula row:

- Row `2`: captions/formulas (`Щит`, `Номер линии`, links from `Разбивка_по_щитам` and `Нагрузка_щитов`).
- Result: skipped as metadata row.

## How To Use This Table For New OLS Projects

1. Put project workbook path into settings (`EOM_EXCEL_PATH` or Ribbon button `Путь Excel`).
2. Keep the source data in Excel and ensure worksheet `В Акад` is generated and saved.
3. Use rows where `D`, `E`, and `M` are meaningful (not empty, not `#N/A`, not headers).
4. Run `EOM_ИМПОРТ_EXCEL` to refresh output cache.
5. Run `EOM_ПОСТРОИТЬ_ОЛС` and choose insertion point and optional shield filter.

## Quick Validation Checklist

Before running `EOM_ИМПОРТ_EXCEL`:

1. Confirm Excel path is set correctly (`EOM_EXCEL_PATH`).
2. Confirm workbook exists and has sheet `В Акад`.
3. Confirm workbook was saved after recalculation (formula caches are up to date).
4. Confirm target rows in `В Акад` have values in `D`, `E`, `M`.
