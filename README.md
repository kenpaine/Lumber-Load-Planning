# Centerbeam Lumber Car Layout Project

A planning tool for loading 72‑ft centerbeam rail cars with dimensional lumber.
It fills each car row to **exactly 72 ft**, uses every pack when the math allows,
groups like packs into columns, and tracks every pack by **product**, **length**,
and **grade**.

---

## Screenshots

### Browser app — `centerbeam_layout_planner.html`

![Centerbeam Lumber Layout Planner — browser app](html_app_screenshot.png)

### Excel workbook — `Centerbeam_Lumber_Layout_Planner.xlsm`

Macro‑enabled workbook with one‑click **Solve Layout** / **Clear Grid** / **Clear All**
buttons on the Planner sheet.

![Centerbeam Lumber Layout Planner — Excel workbook](excel_screenshot.png)

---

## What's in this project

| File | What it is |
|------|-----------|
| `Centerbeam_Lumber_Layout_Planner.xlsm` | **The main tool (recommended).** Macro‑enabled workbook with built‑in **Solve Layout / Clear Grid / Clear All** buttons. Click *Enable Content* on open. |
| `Centerbeam_Lumber_Layout_Planner.xlsx` | Same workbook **without macros** (no buttons) — use if your environment blocks macros. |
| `CenterbeamSolver.bas` | The VBA module behind the buttons (already embedded in the `.xlsm`; kept here as source). |
| `centerbeam_layout_planner.html` | **Browser app** with the full feature set (open in any browser, no install). Line‑item inventory (product × length × grade), the same column‑stacking solver, a proportional visual car layout, Row Detail, a printable one‑page Manifest with pick list, and the length‑colored Pattern Library. |
| `source/` | Python scripts that generate the workbook (for maintenance / regeneration). |

---

## The model

Every pack of lumber is described by three attributes:

- **Product (cross‑section):** 2x4, 2x6, 2x8, 2x10, 4x4, 4x6, 6x6
- **Length:** 8, 10, 12, 14, 16, 18, 20 ft
- **Grade:** 1, 2, 3, 4, 2P, MSR

A car has **10 or 14 rows**; each row must total **exactly 72 ft** end‑to‑end.
Only **length** affects the 72‑ft fit — cross‑section and grade ride along on each
pack and drive the color scheme and the manifest breakdowns.

---

## Using the workbook

### Planner sheet

1. Set **Number of Rows** (10 or 14).
2. Choose how the car is loaded with the **Single product/grade?** toggle:
   - **Yes (single product/grade)** — most cars. Pick the **Product** and **Grade**
     once in the *Single Product / Grade* box, then fill the inventory with just
     **Length + Packs** on each line. On Solve, the blank Product/Grade cells are
     filled in for you. (Any value you do type on a line is kept, so you can still
     mix in the odd different pack.)
   - **No (mixed)** — fill Product + Length + Grade + Packs on every line, as before.
3. Fill the **line‑item inventory** (dropdowns provided). Lineal ft, Placed, and
   Remaining compute automatically.
4. The **Layout Grid** shows the load, one pack per slot, each row summing to 72 ft.
   - **Fill color = product**
   - **Bold colored border = grade**
   - Cell text is the full code `product-length-grade` (e.g. `2x6-14-2`).
5. The **Load Status** box reports utilization, full rows, packs placed/unplaced,
   and an overall verdict.

A solved example load is included so the sheet is populated on open.

### Manifest sheet

Auto‑generates from the Planner:

- **Car Layout** grid (mirror of the Planner, colored by product, bordered by grade)
- **Summary by Product**, **Summary by Grade**
- **Placed Packs — Product × Grade** matrix

Print the Manifest in **landscape** for a clean one‑page sheet.

### Pattern Library sheet

Reference list of every length combination that sums to 72 ft.

---

## One‑click Solve / Clear buttons (macros)

The `.xlsm` already has three buttons wired up on the Planner sheet — just open it and
click **Enable Content**:

- **Solve Layout** — reads the line‑item inventory, fills every selected row to 72 ft, and
  assigns product+grade grouped so like packs stack in columns. In **single product/grade**
  mode it first fills any blank Product/Grade cells from the defaults you picked.
- **Clear Grid** — clears the layout grid only.
- **Clear All** — empties the grid and the line items.

Behind the buttons is the `CenterbeamSolver` VBA module (`SolveLayout`, `ClearGrid`,
`ClearAll`), kept in source form as `CenterbeamSolver.bas`.

**Using the no‑macro `.xlsx` instead?** It has no buttons. You can either add them
yourself (**Developer → Insert → Button (Form Control)**, assign `SolveLayout` /
`ClearGrid` / `ClearAll` after importing `CenterbeamSolver.bas` via **Alt+F11 →
File → Import File…**), or just clear by hand: select the grid and line‑item cells and
press **Delete**.

---

## How the solver works

1. **Length partition.** It enumerates every multiset of lengths that sums to 72 ft
   ("row patterns"), biases toward homogeneous/repeatable patterns, then uses
   backtracking to pick one pattern per row so the per‑length pack counts match
   inventory exactly. This guarantees each chosen row is exactly 72 ft and fills the
   selected number of rows when a solution exists.
2. **Column stacking.** Chosen rows are ordered so identical rows are adjacent and the
   longest packs sit left, so like sizes line up in the same columns.
3. **Product + grade assignment.** For each length, packs are queued grouped by product
   then grade and dealt into that length's slots, so matching product/grade packs stack
   together.

The same algorithm exists in three places (kept in sync):
the workbook's pre‑solved example, the `CenterbeamSolver.bas` macro, and the
Python in `source/`.

---

## Regenerating the workbook (mainteners)

Requires Python 3 with `openpyxl`.

```bash
cd source
python3 build_v3.py        # builds the Planner sheet -> ../cb3.xlsx (writes to /home/claude in original env; adjust paths)
python3 build_v3_p2.py     # adds Manifest, Pattern Library, and the VBA sheet
```

`solver_v3.py` holds the solver and the default example; edit the product list,
grade list, colors, or the default load there. (Paths in the build scripts point at
the original build environment — adjust the input/output paths to your setup.)

---

## Notes & limits

- The 72‑ft fit depends only on **length**. Changing product or grade without changing
  the length mix keeps the layout shape the same and only changes colors/labels — which
  is correct for end‑to‑end loading.
- Perfect top‑to‑bottom column alignment holds **within** a block of identical rows;
  where two blocks meet, the arrangement changes because the packs differ.
- If an inventory's lengths can't be partitioned into exact 72‑ft rows, the solver
  fills as many exact rows as possible and flags the rest as unplaced.

---

*Color key — Product (fill):* 2x4 blue · 2x6 green · 2x8 amber · 2x10 purple · 4x4 red · 4x6 teal · 6x6 brown
*Color key — Grade (border):* 1 green · 2 blue · 3 orange · 4 red · 2P purple · MSR teal

---

## License

Copyright (c) 2026 Ken Paine

This project is licensed under the **GNU General Public License v3.0** — see the
[LICENSE](LICENSE) file for the full text.
