# CLI Command Reference

> Auto-generated from the built `excelcli` runtime. Use these exact command and parameter names.

- [analysis](commands/analysis.md)
- [batch](commands/batch.md)
- [calculationmode](commands/calculationmode.md)
- [chart](commands/chart.md)
- [chartconfig](commands/chartconfig.md)
- [conditionalformat](commands/conditionalformat.md)
- [connection](commands/connection.md)
- [datamodel](commands/datamodel.md)
- [datamodelrelationship](commands/datamodelrelationship.md)
- [diag](commands/diag.md)
- [drawing](commands/drawing.md)
- [namedrange](commands/namedrange.md)
- [pivottable](commands/pivottable.md)
- [pivottablecalc](commands/pivottablecalc.md)
- [pivottablefield](commands/pivottablefield.md)
- [powerquery](commands/powerquery.md)
- [pythoninexcel](commands/pythoninexcel.md)
- [querytable](commands/querytable.md)
- [range](commands/range.md)
- [rangeedit](commands/rangeedit.md)
- [rangeformat](commands/rangeformat.md)
- [rangelink](commands/rangelink.md)
- [screenshot](commands/screenshot.md)
- [service](commands/service.md)
- [session](commands/session.md)
- [sheet](commands/sheet.md)
- [slicer](commands/slicer.md)
- [table](commands/table.md)
- [tablecolumn](commands/tablecolumn.md)
- [vba](commands/vba.md)
- [window](commands/window.md)
- [workbook](commands/workbook.md)
- [worksheetstyle](commands/worksheetstyle.md)
- [xmlmap](commands/xmlmap.md)

## Common Pitfalls

- `--values-file` requires an existing JSON or CSV file; use `--values` for inline JSON.
- `--timeout` ranges are action-specific: session open/create/test accepts 10-3600; Power Query refresh/refresh-all accepts 0-2147483 (0 keeps the default); other generated timeout actions accept 1-2147483.
- `pythoninexcel get-result --max-wait-seconds` must be at least 1 and shorter than the session operation timeout.
- `--values` and list parameters use JSON arrays; range values use a two-dimensional array.
- Power Query operations may take 30 seconds or longer; use a deliberate data-operation timeout or 0 for the default.
