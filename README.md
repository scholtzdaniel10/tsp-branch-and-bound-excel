# TSP Branch-and-Bound Excel

This repository contains materials to build a TSP Branch & Bound prototype in Excel:

- VBA modules (modules/*.bas) implementing:
  - A best-first Branch & Bound driver for TSP
  - MST lower bound (Prim) and an Assignment/Hungarian lower bound
  - Utilities: nearest-neighbour heuristic, copy-as-answer exporter

- Example distance matrices (examples/symmetric.csv, examples/asymmetric.csv)

- INSTRUCTIONS.md with step-by-step setup (how to assemble the .xlsm workbook, name ranges, import modules, and wire buttons).

## How to use
1. Follow INSTRUCTIONS.md to create the .xlsm workbook in Excel and import the VBA modules.
2. Load one of the example matrices and name the range `Dist`.
3. Set `nCities`, `StartCity`, and choose `BoundMethod` (MST or Hungarian).
4. Run the `RunTSPBranchAndBound` macro and use `CopyValidBranches` to get a printable exam-ready iteration table.

## Notes
- The Hungarian bound is stronger but slower; MST is the fast default.
- The macros are designed for n up to ~10–12 for reasonable run times in Excel.