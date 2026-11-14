# heuristic-constraint-optimizer

A **multi-objective, constraint-aware heuristic optimiser written in VBA** for
Excel. It adjusts the *quantity* of each item in a table so that two goals
are satisfied at once, while honouring a set of hard per-item constraints:

- **Goal A (aggregate constraint):** the total value of all items should be as
  close as possible to a target total.
- **Goal B (attribute objective):** the *aggregate penalty per unit* (cost, risk,
  or any per-unit score you define) should stay at or below a ceiling.

**Hard per-item constraints**

- Each item's per-unit value must stay inside `[UNIT_VALUE_MIN, UNIT_VALUE_MAX]`.
- Quantity must be a multiple of the item's package size.
- Quantity may not drop below `MIN_QTY_RATIO` (nor exceed `MAX_QTY_RATIO`) of its
  original quantity.
- Items you mark as *protected* are never changed.

## Why this is interesting

The engine is a small, dependency-free implementation of techniques usually found
only in heavier optimisation libraries:

1. **Multi-objective search** — it balances a physical total against an attribute
   objective that can conflict with it.
2. **Bounded random walk** — it probes inside safe thresholds with `Rnd()` to
   avoid getting stuck on a bad local solution.
3. **Multi-stage fallback** — coarse normalisation &rarr; precise heuristic loop
   &rarr; unique-item formula micro-adjust &rarr; penalty optimisation &rarr;
   forced reduce (guarantees a usable result even when goals conflict).
4. **Anti-infinite-loop guard** — a wall-clock time limit, a max-iteration cap,
   and a no-progress bailout keep the macro from hanging.

## What's in the box

| File | Purpose |
|------|---------|
| `src/ConstrainedAllocator.bas` | The engine (clean-room, domain-independent). |
| `docs/Configuration.md` | Every constant explained + a second domain example. |
| `examples/sample-data.xlsx` | A fictional workbook that mirrors the expected layout. |
| `README.md` / `LICENSE` / `PUSH_GUIDE.md` | Docs. |
| `.gitignore` / `.gitattributes` | Repo hygiene. |

## Quick start

1. Open `examples/sample-data.xlsx` in Excel.
2. Import `src/ConstrainedAllocator.bas` into the VBA editor (`Alt+F11` &rarr;
   *File &rarr; Import File*).
3. Run `AdjustAllocation` for a fast, deterministic pass, then `PreciseTune` for
   the heuristic refinement.

> The sample numbers are invented and have no real-world meaning. They only show
> the layout so you can drop in your own data and tune the constants.

## Porting to your own problem

Edit the `CONFIGURATION` block at the top of `ConstrainedAllocator.bas`:

- Map your columns to the `..._COL` constants.
- Set the per-unit bounds, penalty thresholds, tolerances, and loop guards.
- Put your own protected-item keywords in `IsProtectedItem`.

See `docs/Configuration.md` for the full list and a worked second example
(budget allocation).

## Non-affiliation

This repository is an independent, general-purpose algorithm project. It is **not**
affiliated with, endorsed by, or derived from any commercial product, employer, or
proprietary system. All code here is an original clean-room re-implementation of a
generic optimisation pattern; it contains no third-party data, branding, or
business logic.

## Legal

- The VBA source is released under the **MIT License** — Copyright (c) 2026
  ivy92888-glitch. The license covers only the author's original contribution in
  this repository.
- This project does **not** redistribute, reference, or attempt to license any
  third party's confidential information, trade secrets, or intellectual
  property. If you use this code inside a proprietary system, you are responsible
  for ensuring your use complies with your own organisation's policies.
- No personal data, customer records, or proprietary datasets are included. The
  example workbook is entirely fictional.
