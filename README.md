# Clinicaccelerator-test

Small Excel to JSON converter for the clinic scoreboard test.

## How to run

Install the dependency:

```bash
pip install -r requirements.txt
```

Run the converter from this folder:

```bash
python convert.py "Scoreboard Test.xlsx" output.json
```

Both arguments are optional. If you do not pass them, the script uses `Scoreboard Test.xlsx` and writes `output.json`.

If the input file is missing or cannot be read as an Excel workbook, the script prints a short error message and exits with a non-zero status code.

## JSON shape

The JSON keeps one entry per sheet. Inside each sheet, I split the data into two main parts:

- `columns`: the map of what every Excel column means.
- `records`: the actual weekly rows, keyed by the generated metric IDs from `columns`.

I chose this because the spreadsheet is wide and a lot of columns share similar names. If every row repeated all of the header details, the JSON would be much larger and harder to scan. If the output only had raw arrays, it would be smaller but annoying to query. This structure keeps the metric definitions in one place and keeps the weekly data simple.

```json
{
  "source_file": "Scoreboard Test.xlsx",
  "sheets": [
    {
      "name": "SCOREBOARD",
      "merged_ranges": ["AJ1:AP1"],
      "columns": [
        {
          "id": "total_revenue_all_services_b",
          "kind": "metric",
          "column": "B",
          "label": "Total Revenue - All Services",
          "focus": "Financial",
          "source": "EMR",
          "role": "J"
        },
        {
          "id": "revenue_collected_4_wk_avg_i",
          "kind": "metric",
          "column": "I",
          "label": "Revenue Collected (4 wk avg)",
          "focus": "Collection",
          "source": "Ratio",
          "role": "J"
        }
      ],
      "records": [
        {
          "row": 8,
          "week": "2026-02-16",
          "values": {
            "total_revenue_all_services_b": {
              "column": "B",
              "label": "Total Revenue - All Services",
              "value": 40454.28
            },
            "revenue_collected_4_wk_avg_i": {
              "column": "I",
              "label": "Revenue Collected (4 wk avg)",
              "value": {
                "value": 1.00441053,
                "formula": "=if(B8=\"\",\"Formula\",sum(H8:H10)/sum(B8:B10))"
              }
            }
          }
        }
      ]
    }
  ]
}
```

The `id` field is the main lookup key. It is based on the cleaned metric label plus the Excel column letter, such as `utilization_dw` or `total_revenue_all_services_b`. The column letter is important because the sheet has repeated labels like `Utilization`, `PVA (4 wk avg)`, and `TP Utilization`. Without the column letter, those values would overwrite each other or need awkward numbering.

Each weekly record only includes values that actually exist for that week. The full column list is still available above it, so a dashboard or another script can always look up the missing context when needed.

## Messy spreadsheet choices

Merged cells are recorded in `merged_ranges`, and their visible value is carried into the affected header cells. For example, the `PHONE PERFORMANCE` merged header applies across several phone-related columns. JSON has no real merged-cell concept, so copying the visible value into each affected column makes the output easier to query while still recording the original merged range.

Spacer columns stay in the `columns` list as `kind: "spacer"`. I kept them because the instruction said not to lose headings or data, and spacer columns are part of the original layout. The trade-off is that the column list is a little longer, but the weekly `records` stay clean because empty spacer columns are not repeated in every row.

Formula cells keep the cached Excel result and the original formula:

```json
{
  "value": 1.00441053,
  "formula": "=if(B8=\"\",\"Formula\",sum(H8:H10)/sum(B8:B10))"
}
```

This is a trade-off too. The script does not try to recalculate formulas itself, because Excel formulas can be messy and Python would not always match Excel exactly. Instead, it keeps the last value saved in the workbook and the formula text. That gives a dashboard the number it needs, while still letting a developer audit where it came from.

Dates are written as ISO strings like `2026-02-16`, which is easier for code to sort and filter than Excel's internal date format. Odd source cells like hyperlinks are preserved as formula objects when Excel stores them that way, so the link text and the formula are both still available.

Blank cells are not repeated in each weekly record. This makes `output.json` smaller and easier to read. The trade-off is that a missing key can mean "blank in the spreadsheet", so code reading the file should use the `columns` list as the source of all possible metrics.

With another couple hours, I would add a small validation report that flags duplicate labels, broken formulas, and columns that have data but very little header context.
