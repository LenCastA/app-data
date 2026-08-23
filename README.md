# Horext Public Updater

## Seed CSV format

Files in `seeds/` use one consistent CSV format:

- UTF-8 without a byte-order mark (BOM)
- comma delimiter
- RFC 4180 quoting (`""` represents a quote inside a quoted value)
- one unique, non-empty name for every column
- a final newline; LF and CRLF line endings are both accepted

Text containing commas, quotes, or newlines must be quoted according to CSV rules. Generate or edit these files with a CSV-aware tool; do not construct rows by joining values with commas.

## Import an hourly-load workbook

Install `openpyxl`, then run the importer with the workbook and destination seed. Passing the current seed as `--previous` preserves the known teacher/type for source rows marked only as `Solo para PC`.

```powershell
python -m pip install openpyxl
python scripts/import_hourly_load.py "C:\path\CARGA-HORARIA.xlsx" "seeds\hl_Carga Horaria 2026-2_2026-2_I.csv" --previous "seeds\hl_Carga Horaria 2026-2_2026-2_I.csv"
python scripts/validate_seeds.py
```
