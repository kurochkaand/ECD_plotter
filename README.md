# ECD_plotter

Windows tools for merging instrument ECD exports and analyzing spectral components.
Requires Python 3.12 or newer.

## Setup

Run these commands in PowerShell from the project folder:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Merge instrument exports

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Szeged\K.Andrii\260730-PLGA-titration-Zn"
```

Alternatively, edit `DEFAULT_FOLDER` in `merge.py` and run it without `--folder`.

The script reads CSV files directly inside the selected folder, ordered by Windows
creation time, with filename order used to break ties. Each file must contain an
`XYDATA` block with wavelength, CD, and HT columns followed by
`##### Extended Information`. Exactly one CSV filename must start with `water`
(case insensitive); it supplies the water reference.

On the first run, the script creates `concentration_table.txt` in the data folder
with a concentration coefficient of `1.0` and `Ph` of `7.0` for each sample.
Edit the tab-separated table and rerun the command to apply the coefficients.
Existing two-column tables without `Ph` are also accepted. Coefficients must be
finite and greater than zero. `Ph` is retained as metadata and is not used in
the correction.

The workbook is saved beside the data folder as `<folder-name>-combined.xlsx`:

- `RAW`: wavelength plus the original CD and HT columns for every input file.
- `SUMMARY`: wavelength, water CD, raw sample CD, a zero separator, then corrected
  sample CD calculated as `(sample - water) / concentration_coefficient`.

All input wavelength axes must match exactly. Rerunning replaces the output workbook.
Close it in Excel before rerunning.

## Analyze spectral components

```powershell
.\.venv\Scripts\python.exe spectral_component_analysis.py --file "X:\Data\corrected.xlsx" --components 2 --no-show
```

See [the analysis guide](README_spectral_component_analysis.md) for input layout,
configuration, outputs, and scientific limitations.

## Check the code

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```
