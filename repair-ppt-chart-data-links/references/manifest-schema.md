# Manifest and update schema

## Inventory authorization

`chart-inventory.json` is the scan result. Mutation requires both:

```json
{
  "selection_confirmed": true,
  "charts": [{"chart_id": "slide-15-chart-1", "selected": true}]
}
```

Exact old external target paths are grouping keys. Never group solely by filename.

## Updates

Normalize user-supplied values into:

```json
{
  "updates_confirmed": true,
  "updates": [
    {
      "chart_id": "slide-15-chart-1",
      "sheet": "Sheet1",
      "target_cell": "K28",
      "old_value": 167.8,
      "new_value": 183.3806,
      "unit": "亿吨",
      "source_note": "用户提供"
    }
  ],
  "text_replacements": [
    {
      "slide": 15,
      "old_text": "2025年1–11月",
      "new_text": "2025年",
      "confirmed": true
    }
  ]
}
```

Before applying, verify chart, series/sheet, cell, value type, unit, category period, and old value. Reject ambiguous matches and business inconsistencies; never silently correct them.

`repair-manifest.json` is the audit record and contains source hash, selected charts, workbook groups, recovered cells, update provenance, text confirmations, output paths, and `relink_allowed`.

Each recovered reference records `cache_indexes`, expanded `values`, formula cells and any accepted `intentional_gap`. Each numeric series receives a `data_mappings` item with the exact value range, category range, series index, point indexes and `cache_excel_match` result.

Office and portable behavioral probes additionally record `series_name`. The OOXML `series_index` remains an audit index; the Office verifier must resolve the live COM series by exact name and reject duplicate or missing matches.

After relinking, `native_relink` records the recipe manifest, temporary Office paste results, transplanted chart parts and cleanup result. The only supported relink method is:

```text
native Excel chart carrier + Office PasteSourceFormatting + OOXML relationship transplant
```

Portable mode writes a separate `bundle-manifest.json` with a `bundle_id`, master/workbook SHA-256 values, each authorized `chart_rel_part`, its workbook path relative to the bundle, and one named numeric behavioral probe. The launcher resolves these relative paths to current-machine absolute paths; the PPT itself does not rely on relative-link fallback.
