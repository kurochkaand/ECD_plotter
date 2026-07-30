# Spectral Component Analysis with Known Pure References

## Input workbook

- Excel sheet name: `Sheet1`
- Row 1: headers
- Column A: wavelength values, ascending or descending
- Columns B onward: measured spectra
- One or more spectral columns may be declared as pure-component references

## Configure the script

Open `spectral_component_analysis.py` and edit:

```python
EXCEL_PATH = Path(r"/Users/andriikurochka/Documents/PGA_methals/K.Andrii/corrected.xlsx")
MIN_WAVELENGTH = 190.0
MAX_WAVELENGTH = 260.0
```

The letters refer to real Excel column positions, not header names. Column A is reserved for wavelength. Columns listed in `KNOWN_COMPONENT_COLUMNS` are treated as fixed pure spectra and excluded from the mixture/sample set.

Leave this setting as `None` for automatic estimation of additional components:

```python
N_UNKNOWN_COMPONENTS = None
```

To force a result, use the number of **unknown components only**:

```python
N_UNKNOWN_COMPONENTS = 2
```

Thus, with known columns B and C and two unknown components, the final model contains four components in total.

## Install and run

```bash
python -m pip install -r requirements_spectral_analysis.txt
python spectral_component_analysis.py
```

Command-line override:

```bash
python spectral_component_analysis.py \
  --file spectra.xlsx \
  --known-columns B C \
  --min-wavelength 300 \
  --max-wavelength 700 \
  --unknown-components 2
```

Use `--no-show` to save figures without opening windows. Use `--no-mcr` to run the fixed-reference fit and residual SVD/PCA without MCR-ALS.

## Analysis workflow

1. The known pure spectra are fitted to every remaining sample using linear least squares with non-negative scale coefficients by default.
2. The known-only reconstruction is subtracted from each mixture.
3. SVD of the residual matrix estimates how many additional independent spectral patterns remain.
4. PCA describes the residual variation and provides loading and score plots.
5. For non-negative data, fixed-reference MCR-ALS jointly refits all contributions while keeping the known B/C spectra exactly unchanged and estimating only the unknown spectra.

The automatic procedure can return zero unknown components when the residual signal is below `MIN_UNKNOWN_SIGNAL_FRACTION`. The default is 0.005, corresponding to 0.5% of the total mixture-matrix norm. Lower it when very weak additional species are important, but expect greater sensitivity to noise and baseline artifacts.

## Main outputs

- `known_reference_spectra.csv`
- `known_component_coefficients.csv`
- `known_only_reconstructed_spectra.csv`
- `residual_after_known_components.csv`
- `residual_svd_diagnostics.csv`
- `unknown_residual_pca_loadings.csv`
- `unknown_residual_pca_scores.csv`
- `fixed_mcr_component_spectra.csv`
- `fixed_mcr_contributions.csv`
- `fixed_mcr_reconstructed_spectra.csv`
- diagnostic PNG plots

The known-component coefficients are scale factors relative to the measured pure-reference spectra. They are directly proportional to concentration only when path length, reference concentration, and instrumental scaling are appropriately controlled.

## Scientific limitations

Residual rank estimates independent spectral patterns rather than proving the number of chemical species. Strong spectral overlap, fixed concentration ratios, reactions that do not vary independently, baseline drift, scattering, and noise can change the apparent rank. Unknown MCR spectra are constrained and constituent-like but are not mathematically unique without further chemical constraints.
