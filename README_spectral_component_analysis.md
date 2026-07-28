# Spectral Component Analysis CLI

## Input workbook

- Excel sheet name: `Sheet1`
- Row 1: headers
- Column A: wavelength values, ascending or descending
- Columns B onward: one measured spectrum per column

## Configure

Open `spectral_component_analysis.py` and edit:

```python
EXCEL_PATH = Path(r"/Users/andriikurochka/Documents/PGA_methals/K.Andrii/corrected.xlsx")
MIN_WAVELENGTH = 190.0
MAX_WAVELENGTH = 260.0
```

You can also set `N_COMPONENTS` to an integer to override automatic selection.

## Install and run

```bash
python -m pip install -r requirements_spectral_analysis.txt
python spectral_component_analysis.py
```

Optional override without editing the file:

```bash
python spectral_component_analysis.py --file spectra.xlsx \
  --min-wavelength 300 --max-wavelength 700 --components 3
```

Use `--no-show` to save figures without opening windows, or `--no-mcr` to run only SVD/PCA.

## Methods

- **SVD rank estimation:** estimates the number of independent spectral patterns using the Gavish–Donoho singular-value threshold and also reports the count required to explain 99.5% of spectral energy.
- **PCA:** supplies scores, loading spectra, explained variance, reconstructions, and residual diagnostics.
- **MCR-ALS:** when all values are non-negative, estimates non-negative constituent-like spectra and their relative contributions using alternating non-negative least squares.

The output folder contains PNG figures and CSV tables.

## Scientific limitation

A numerical component count is not automatically identical to the number of chemical species. Co-linear component spectra, fixed concentration ratios, baseline drift, scattering, and noise may change the apparent rank. MCR-ALS solutions can also be rotationally ambiguous unless reference spectra, concentration constraints, selectivity regions, or a known chemical model are supplied.
