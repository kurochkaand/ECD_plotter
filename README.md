# ECD_plotter

PowerShell cheat sheet. Python 3.12+; run from the project folder.

[Setup](docs/REFERENCE.md#setup) | [SESCA installation](docs/REFERENCE.md#estimate-secondary-structure-with-sesca)

## Merge, convert, and plot

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --subtract "H2O" --plot --smooth 50 --ht-threshold 500
```

First run creates `concentration_table.txt`. Set `pathlength_in_mm`,
`number_of_amino_acids`, and each sample's actual molecule `concentration_in_mM`, then rerun.

**Output:** `<folder-name>-combined.xlsx` beside the data folder, including `MRE` and `Epsilon`.

## SESCA

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --plot
```

Reads `MRE`; adds `SESCA_results` and `SESCA_fit` to the same workbook.
Add `--shared-scaling` for one automatic factor, or `--shared-scaling 1.33334` to fix it.

## Spectral components

**Input:** separate numeric workbook, `Sheet1`; headers in row 1, wavelengths in A,
at least two selected spectra in B onward. Copy numeric `MRE` values for MRE analysis.

```powershell
.\.venv\Scripts\python.exe spectral_component_analysis.py --file "X:\Data\corrected.xlsx" --components 2 --no-show
```

**Output:** CSV tables and PNG plots in `component_analysis_output`.
Omit `--components 2` for automatic selection; omit `--no-show` to display plots.

## Remember

- Close the combined workbook in Excel before rerunning merge or SESCA.
- After table edits, rerun merge, then SESCA. Every merge replaces the workbook.
- Each script accepts `--help`.

[Full reference](docs/REFERENCE.md) | [Spectral component analysis guide](docs/SPECTRAL_COMPONENT_ANALYSIS.md)
