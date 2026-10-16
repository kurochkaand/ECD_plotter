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

The script reads CSV files directly inside the selected folder, ordered by the
acquisition date in cell B5 and time in cell B6, earliest first. Dates use `YY/MM/DD`
(for example, `26/10/14` means October 14, 2026) or `YYYY/MM/DD`. Times can use
`18:41:36` or `6:41:36 PM`. Filename order breaks ties. Missing or invalid timestamps
produce an error identifying the file and cell. Each file must contain an
`XYDATA` block with wavelength, CD, HT, and Abs columns followed by
`##### Extended Information`. Older three-column exports are also accepted, with
blank Abs values.
Every CSV is treated as a sample.

On the first run, the script creates `concentration_table.txt` in the data folder
with a concentration coefficient of `1.0` for each sample. `Ph` is extracted from
the filename: `pH7.43` gives `7.43`, and `pH7.9` gives `7.9`. A space after `pH`
is accepted. Files without pH in the name, such as `H2O_1mm_185-300nm_4sc_20nm-1.csv`,
use the default `7.0`.
Edit the tab-separated table and rerun the command to apply the coefficients.
Each run sorts the table rows into acquisition order, keeping existing coefficients
and `Ph` values attached to their spectrum names. All workbook sheets use this
same spectrum order for their columns.
Missing `Ph` columns or blank values in existing tables are populated the same way.
To replace existing pH values with those from the filenames while keeping your
concentration coefficients, add `--refresh-ph`:

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14" --refresh-ph
```

Values for filenames without pH are preserved when refreshing. Coefficients must be
finite and greater than zero. `Ph` is retained as metadata and is not used in
the correction.

The workbook is saved beside the data folder as `<folder-name>-combined.xlsx`:

- `CD`: wavelength plus the original CD signal for each input file, in mdeg.
- `HT`: wavelength plus the original HT signal for each input file, in V.
- `Abs`: wavelength plus the original absorbance for each input file.
- `CD_subtr`: wavelength, raw sample CD, a zero separator, then normalized
  sample CD calculated as `sample / concentration_coefficient`, or reference-
  subtracted CD formulas when `--subtract` is supplied.

Each channel sheet starts with `wavelength_nm`, followed by one column per spectrum
named after the CSV filename without its extension. The `RAW` sheet is replaced
by these three channel sheets. Only CD is concentration-normalized in `CD_subtr`.

To subtract a reference spectrum, provide its filename prefix:

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14" --subtract "H2O"
```

The prefix must match exactly one spectrum, ignoring case. The reference remains
in the raw columns of `CD_subtr`; after the zero separator, each other spectrum
gets Excel formulas for `(sample - reference) / concentration_coefficient`.
The reference itself is omitted from the corrected block. With a coefficient of
`1.0`, a formula might be `=B2-$O2`; with a coefficient of `2.0`, it becomes
`=(B2-$O2)/2.0`. The reference column is fixed, and each wavelength uses its own
row. The raw `CD`, `HT`, and `Abs` sheets retain their original signals.

Excel calculates the formulas when the workbook opens. Open and save the workbook
in Excel before reading calculated subtraction values with Python or another tool
that relies on stored formula results.

All input wavelength axes must match the earliest acquired CSV. A mismatch
reports the first differing data row, physical CSV line numbers, expected and
actual wavelengths, reference filename, and row counts. Rerunning replaces the
output workbook. Close it in Excel before rerunning.

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
