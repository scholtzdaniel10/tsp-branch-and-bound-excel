# Setup Instructions — Build the TSP_BranchAndBound.xlsm workbook

Follow these steps in Excel (Windows recommended) to create the macro-enabled workbook.

### A) Create workbook and sheets
1. Open Excel -> New workbook.
2. Add and rename sheets:
   - TSP_Symmetric
   - TSP_Asymmetric
   - Examples
   - TSP_Dashboard (optional)
3. Save as: `TSP_BranchAndBound.xlsm` (Excel Macro-Enabled Workbook).

### B) Paste the example matrices
1. Open `examples/symmetric.csv` and copy the numeric 6x6 block into the Examples sheet (or paste directly into TSP_Symmetric).
2. Open `examples/asymmetric.csv` and paste into Examples (or TSP_Asymmetric).

### C) Create named ranges / input cells on TSP_Symmetric (repeat for TSP_Asymmetric)
1. Choose a clear area for inputs (suggest top-left).
   - Cell C4: enter number of cities (nCities) — e.g., 6
   - Cell C5: enter start city index (StartCity) — e.g., 1
   - Cell C9 and onward: paste the numeric n×n distance matrix for this sheet.
   - Cell C16: BoundMethod (create a dropdown with values: MST,Hungarian)
   - Cell C17: maxNodes (0 = unlimited)
   - Cell C18: maxTime (seconds, 0 = unlimited)
   - Cell C19: prec (precision, default 3)
2. Create named ranges:
   - Select the numeric n×n matrix and define the name `Dist` (Formulas -> Define Name).
   - Select cell C4 and name it `nCities`.
   - Select cell C5 and name it `StartCity`.
   - Select cell C17 and name it `maxNodes`.
   - Select cell C18 and name it `maxTime`.
   - Select cell C19 and name it `prec`.
   - Select cell C16 and name it `BoundMethod`.

### D) Prepare iteration log area
1. Reserve row 30 downward for the iteration log.
2. Header row at row 30 (or adjust the `rowIterStart` constant in VBA):
   - A: NodeID
   - B: Parent
   - C: FixedTour
   - D: Excluded
   - E: LB
   - F: Branch
   - G: IncumbentTour
   - H: IncumbentVal
   - I: Prune
   - J: Notes

### E) Add VBA modules
1. Developer -> Visual Basic -> Insert -> Module.
2. For each .bas file in `modules/` (TSP_BranchAndBound.bas and TSP_Hungarian_and_LB.bas), copy the contents and paste into separate modules in the VBA editor.
3. Save.

### F) Add buttons (optional)
1. Developer -> Insert -> Button (Form Control).
2. Assign `RunTSPBranchAndBound` to one button.
3. Add another button and assign `CopyValidBranches`.
4. (Optional) Add Reset button mapped to a short macro that clears iteration log rows.

### G) Test run
1. On TSP_Symmetric set:
   - nCities = 6
   - StartCity = 1
   - BoundMethod = MST (or Hungarian)
   - prec = 3
2. Click Run button or run macro `RunTSPBranchAndBound`.
3. The Iteration table will populate from row 30 onward.
4. Click `CopyValidBranches` to create a compact exam-style sheet "TSP_CopyAnswer".

### H) Packaging for GitHub
1. Create local folder:
   mkdir tsp-branch-and-bound-excel
   cd tsp-branch-and-bound-excel
2. Copy the following files into the folder:
   - README.md
   - INSTRUCTIONS.md
   - modules/TSP_BranchAndBound.bas
   - modules/TSP_Hungarian_and_LB.bas
   - examples/symmetric.csv
   - examples/asymmetric.csv
   - (Optionally) add the built `TSP_BranchAndBound.xlsm` if you created it locally.
3. Initialize git and push:
   git init
   git add .
   git commit -m "Initial commit — TSP B&B Excel prototype files"
   (create repo on GitHub via web UI or gh CLI; set visibility to private)
   git remote add origin https://github.com/scholtzdaniel10/tsp-branch-and-bound-excel.git
   git branch -M main
   git push -u origin main
