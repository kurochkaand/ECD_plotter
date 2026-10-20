# ECD_plotter

Cheat sheet for merging ECD exports, SESCA deconvolution, and spectral component analysis.
Requires Python 3.12+. Run commands in PowerShell from the project folder.

## Setup

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

For SESCA fitting, also follow the [SESCA installation instructions](docs/REFERENCE.md#convert-to-mre-and-estimate-secondary-structure-with-sesca).

## Data folder

Set this to the folder containing your instrument CSV exports for the next two commands:

```powershell
$folder = "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd"
```

## Merge and plot

```powershell
.\.venv\Scripts\python.exe merge.py --folder $folder --subtract "H2O" --plot --smooth 50 --ht-threshold 500
```

- `--subtract "H2O"`: reference filename prefix; must match exactly one CSV. Omit for no subtraction.
- `--smooth 50`: smooth CD with a 50-point window. Omit for no smoothing.
- `--ht-threshold 500`: leave CD blank where HT exceeds 500 V. Omit for no HT filtering.

The first run creates `concentration_table.txt` inside the data folder. Edit its
concentration coefficients, pH, and short names; remove rows to exclude samples.
Rerun the same command to apply edits. Add rows manually to include new CSVs.

**Output:** `<folder-name>-combined.xlsx` beside the data folder, with `CD`, `HT`,
`Abs`, `CD_subtr`, and `Abs_subtr` sheets.

## Convert to MRE and run SESCA

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder $folder --path-in-mm 1 --num-of-aa 20 --concentration-in-mM 0.1 --plot
```

Replace `1`, `20`, and `0.1` with your cell path length in mm, residues per molecule,
and peptide/protein molecule concentration in mM before dilution (coefficient `1.0`).

- Add `--shared-scaling` for one automatic scaling factor across the series, or `--shared-scaling 1.33334` to fix it.
- Replace `--plot` with `--mre-only` to convert without SESCA; omit shared scaling in this mode.

**Output:** `MRE`, `MRE_parameters`, `SESCA_results`, and `SESCA_fit` sheets in the
same combined workbook. `--mre-only` writes the first two and removes old SESCA sheets.

## Analyze spectral components

Prepare a separate numeric workbook with `Sheet1`: headers in row 1, wavelengths
in column A, and at least two selected spectra in B onward. Copy the corrected
columns from `CD_subtr` or use `MRE`; omit raw columns and the zero separator.
Open and save the source workbook in Excel before copying formula results.

```powershell
.\.venv\Scripts\python.exe spectral_component_analysis.py --file "X:\Data\corrected.xlsx" --components 2 --no-show
```

- Omit `--components 2` for automatic component selection.
- Omit `--no-show` to display plots; add `--no-mcr` to run SVD/PCA only.
- Add `--min-wavelength 190 --max-wavelength 260` to set the analysis range.

**Output:** CSV tables and PNG plots in `component_analysis_output` under the
current working folder. Use `--output "X:\Data\analysis"` to choose another folder.

## Workflow reminders

- Close the combined workbook in Excel before rerunning merge or deconvolution.
- Editing `concentration_table.txt` requires rerunning merge, then deconvolution.
- Merge recreates the entire workbook, so rerun deconvolution afterward.
- Close plot windows to return to the terminal. Each script accepts `--help`.

## Detailed documentation

- [Merge and deconvolution reference](docs/REFERENCE.md)
- [Spectral component analysis guide](docs/SPECTRAL_COMPONENT_ANALYSIS.md)
