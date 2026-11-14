# Configuration

Everything you need to adapt the engine lives in the `CONFIGURATION` block at the
top of `src/ConstrainedAllocator.bas`. You change **values**, never the algorithm.

## Column / sheet layout

| Constant | Default | Meaning |
|----------|---------|---------|
| `SRC_SHEET` | `""` | `""` = use the active sheet; otherwise the sheet name. |
| `KEY_COL` | `4` (D) | Column that holds the totals-row marker. |
| `KEY_ROW_TEXT` | `"Total"` | Text in `KEY_COL` that marks the totals row. |
| `PRODUCT_START_ROW` | `4` | First item data row. |
| `NAME_COL` | `3` (C) | Display name; used only by the protected-item check. |
| `QTY_COL` | `15` (O) | **Adjustable** quantity — the macro writes here. |
| `ORIG_QTY_COL` | `16` (P) | Original quantity backup (ratio guards). |
| `BOX_COL` | `14` (N) | Package size; quantity must be a multiple of this. |
| `UNIT_VAL_COL` | `29` (AC) | Per-package value (usually a formula). |
| `UNIT_PRICE_COL` | `30` (AD) | Price per value-unit; items pricier than `PRICE_SKIP` are skipped. |
| `PENALTY_COL` | `31` (AE) | Penalty per unit (the attribute objective). |
| `PENALTY_ITEM_COL` | `18` (R) | Penalty per item (= penalty/unit × qty). |
| `TOTAL_VAL_COL` | `8` (H) | Per-item value; summed at the totals row. |
| `TARGET_VAL_COL` | `9` (I) | Desired aggregate total (row below the totals row). |
| `LINK_COL` | `28` (AB) | Optional; a unique link triggers a micro-adjust. |

## Tuning constants

| Constant | Default | Meaning |
|----------|---------|---------|
| `UNIT_VALUE_MIN` / `UNIT_VALUE_MAX` | `3.35` / `35.7` | Allowed per-package value band. |
| `BAND_HI_MIN` / `BAND_HI_MAX` | `3.35` / `12` | Natural-distribution band for high-penalty items. |
| `BAND_LO_MIN` / `BAND_LO_MAX` | `22` / `32` | Natural-distribution band for low-penalty items. |
| `PENALTY_CEILING` | `0.88` | Aggregate penalty/unit must stay at or below this. |
| `PENALTY_HIGH` | `0.35` | Threshold that classifies an item as "high penalty". |
| `PENALTY_MID_HI` | `0.6` | Top of the "mid penalty" band. |
| `PRICE_SKIP` | `40` | Skip items pricier than this (per value-unit). |
| `TOL_COARSE` | `300` | Accept window for the coarse phase. |
| `TOL_PRECISE` | `100` | Accept window for the precise phase. |
| `TIME_LIMIT_SEC` | `45` | Hard wall-clock cap (anti-infinite-loop). |
| `MAX_ITERATIONS` | `500` | Hard iteration cap. |
| `MIN_QTY_RATIO` / `MAX_QTY_RATIO` | `0.3` / `3.5` | Quantity may not drop below 30% / exceed 350% of original. |
| `FORCED_RATIO` / `FORCED_MIN` | `0.1` / `50` | Forced-reduce stage removes ~10% of the gap (min 50). |
| `HILITE_COLOR` | `255` | Cell colour used to mark adjusted quantities. |
| `DEBUG_ON` | `False` | Set `True` to flood `Debug.Print`. |

## Protected items

Edit `IsProtectedItem` to list keywords for rows the engine must never change.
The shipped defaults are placeholders (`ProtectedA`, `ProtectedB`, …).

## Second domain example: budget allocation

Rename the concepts and the engine solves a completely different problem:

- **Total value** &rarr; *total budget to spend*.
- **Penalty per unit** &rarr; *cost premium per unit* (high = overpriced).
- **Per-unit value band** &rarr; *per-line spend bounds* (avoid tiny or monstrous
  line items).
- **Package size** &rarr; *minimum order multiple*.
- **Goal** &rarr; spend the whole budget while keeping the high-premium share low
  and every line item within sane spend bounds.

No code changes are required — only the constants and your data. That is the
point of the clean-room, domain-independent design.
