# ECD_plotter

PowerShell cheat sheet. Python 3.12+; run from the project folder.

[Setup](docs/REFERENCE.md#setup) | [SESCA installation](docs/REFERENCE.md#estimate-secondary-structure-with-sesca)

## Merge, convert, and plot

```powershell
python merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --subtract "H2O" --plot --smooth 50 --ht-threshold 500
python merge.py --folder "X:\Cations_and_PGA\ECD-Szeged\K.Andrii\260722-PLGA-pH-titrat" --plot --subtract water --ht-threshold 450
```

First run creates `concentration_table.txt`. Set `pathlength_in_mm`,
`number_of_amino_acids`, and each sample's actual molecule `concentration_in_mM`, then rerun.

**Output:** `<folder-name>-combined.xlsx` beside the data folder, including `MRE` and `Epsilon`.

## SESCA

```powershell
python deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --plot
```

Reads `MRE`; adds `SESCA_results` and `SESCA_fit` to the same workbook.
Add `--shared-scaling` for one automatic factor, or `--shared-scaling 1.33334` to fix it.

## MRE component count (SVD)

```powershell
python component_analysis_mre.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --plot
```

Reads `MRE` from `<folder-name>-combined.xlsx` beside the data folder. Analyzes
190–250 nm by default, excluding any wavelength row with a blank in any spectrum.
Adds `MRE_SVD_results`, `MRE_SVD_parameters`, and `MRE_SVD_input` to the same workbook.
Reports the noise-threshold rank and the 99.5% spectral-energy rank, without a
four-component cap. No concentration table or additional normalization is needed.
Use `--min-wavelength 195 --max-wavelength 245` to change the range; omit `--plot`
to save results without opening plots. [Details](docs/REFERENCE.md#estimate-component-count-from-mre).

## Spectral components

**Input:** separate numeric workbook, `Sheet1`; headers in row 1, wavelengths in A,
at least two selected spectra in B onward. Copy numeric `MRE` values for MRE analysis.

```powershell
python spectral_component_analysis.py --file "X:\Data\corrected.xlsx" --components 2 --no-show
```

**Output:** CSV tables and PNG plots in `component_analysis_output`.
Omit `--components 2` for automatic selection; omit `--no-show` to display plots.

## Remember

- Close the combined workbook in Excel before running merge, SESCA, or MRE SVD.
- After table edits, rerun merge, then any required SESCA or MRE SVD analyses. Every merge replaces the workbook.
- Each script accepts `--help`.

[Full reference](docs/REFERENCE.md) | [Spectral component analysis guide](docs/SPECTRAL_COMPONENT_ANALYSIS.md)
