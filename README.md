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

## JSON shape

The JSON keeps one entry per sheet. Each sheet has column definitions first, then weekly records:

```json
{
  "source_file": "Scoreboard Test.xlsx",
  "sheets": [
    {
      "name": "SCOREBOARD",
      "columns": [
        {
          "id": "total_revenue_all_services_b",
          "kind": "metric",
          "column": "B",
          "label": "Total Revenue - All Services",
          "focus": "Financial",
          "source": "EMR",
          "role": "J"
        }
      ],
      "records": [
        {
          "week": "2026-02-16",
          "values": {
            "total_revenue_all_services_b": {
              "column": "B",
              "label": "Total Revenue - All Services",
              "value": 40454.28
            }
          }
        }
      ]
    }
  ]
}
```

I used this shape because dashboards usually need two things: a stable list of metric definitions and a simple set of dated rows to query. Metric IDs include the Excel column letter so repeated names like `Utilization` do not overwrite each other.

## Messy spreadsheet choices

Merged cells are recorded in `merged_ranges`, and their visible value is carried into the affected header cells. Spacer columns stay in the `columns` list as `kind: "spacer"` so the original layout is still represented, but they are left out of weekly `values` unless they contain data.

Formula cells keep the cached Excel result and the original formula:

```json
{
  "value": 1.00441053,
  "formula": "=if(B8=\"\",\"Formula\",sum(H8:H10)/sum(B8:B10))"
}
```

Dates are written as ISO strings, and odd source cells like hyperlinks are preserved as formula objects when Excel stores them that way. Blank cells are not repeated in each weekly record, which keeps the output easier to use without throwing away the column map.

With another couple hours, I would add a small validation report that flags duplicate labels, broken formulas, and columns that have data but very little header context.
