# Spectral Component Analysis

Analyze an Excel workbook on Windows using SVD, PCA, and optional non-negative
MCR-ALS. See [the project README](README.md) for Python installation and setup.

## Input workbook

- Excel sheet name: `Sheet1`.
- Row 1: headers.
- Column A: wavelength values, ascending or descending.
- Columns B onward: measured spectra, with at least two spectral columns.
- At least five wavelength points must remain after filtering.

The analysis reads every column after A as a measured spectrum. Prepare a workbook
containing only the spectra you want to analyze. The merge script's `CD_subtr`
sheet includes both raw and corrected data and a zero separator, so select the
appropriate columns into a separate `Sheet1` first. If merging with `--subtract`,
open and save the workbook in Excel to calculate its formulas before preparing
the numeric input for analysis.

## Configure the script

By default, the script looks for `corrected.xlsx` in the current working folder.
To use another default, edit `EXCEL_PATH` in `spectral_component_analysis.py`:

```python
EXCEL_PATH = Path(r"X:\Data\corrected.xlsx")
MIN_WAVELENGTH = 190.0
MAX_WAVELENGTH = 260.0
```

Set `SHEET_NAME` if the worksheet has a different name. Wavelength bounds are
inclusive, and their order does not matter. Duplicate wavelengths are averaged.
Small gaps are interpolated; a spectrum with more than 5% missing values in the
selected interval is rejected.

Leave `N_COMPONENTS = None` for automatic selection. The default
`COMPONENT_SELECTION = "gavish_donoho"` uses a singular-value noise threshold.
The alternative `"variance"` selects enough components to reach
`CUMULATIVE_VARIANCE_TARGET`, which defaults to 0.995. Selection is capped by
`MAX_COMPONENTS`, currently 4, and the available matrix dimensions.

To force the total number of components, set `N_COMPONENTS = 2` or use
`--components 2`. The current implementation estimates all components from the
measured spectra; it does not accept fixed pure-reference columns.

## Run in PowerShell

```powershell
.\.venv\Scripts\python.exe spectral_component_analysis.py --file "X:\Data\corrected.xlsx" --min-wavelength 190 --max-wavelength 260 --components 2
```

Supported command-line overrides:

| Option | Purpose |
| --- | --- |
| `--file` | Input workbook path. |
| `--min-wavelength` | Lower wavelength bound in nm. |
| `--max-wavelength` | Upper wavelength bound in nm. |
| `--components` | Force the total component count. |
| `--output` | Results folder; defaults to `component_analysis_output`. |
| `--no-show` | Save figures without opening plot windows. |
| `--no-mcr` | Skip MCR-ALS and run SVD/PCA only. |

## Analysis workflow

1. Read and validate the workbook, then select and sort wavelengths.
2. Estimate numerical rank using SVD of the uncentered spectral matrix by default.
3. Perform PCA, mean-centered across spectra by default.
4. For non-negative input, optionally fit MCR-ALS spectra and contributions.
5. Save CSV tables and diagnostic PNG plots, then display plots unless disabled.

MCR-ALS is skipped when the input contains materially negative values. Signed ECD
and difference spectra can still be analyzed using SVD/PCA. Do not shift them
merely to force non-negativity.

## Outputs

The results folder contains:

- `filtered_input_spectra.csv`
- `svd_diagnostics.csv`
- `pca_loadings.csv`
- `pca_scores.csv`
- `pca_reconstructed_spectra.csv`
- `mcr_component_spectra.csv`, when MCR-ALS runs
- `mcr_contributions.csv`, when MCR-ALS runs
- `mcr_reconstructed_spectra.csv`, when MCR-ALS runs
- PNG plots of input spectra, singular values, explained energy, PCA loadings,
  scores, and residuals, plus MCR spectra, contributions, and reconstruction when
  available

## Scientific limitations

Numerical spectral rank estimates independent patterns rather than proving the
number of chemical species. Strong spectral overlap, fixed concentration ratios,
baseline drift, scattering, and noise can change the apparent rank. MCR spectra
and contributions are not chemically unique without additional constraints or
reference information. Contributions use relative scaling and do not establish
absolute concentrations.
