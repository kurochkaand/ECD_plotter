# ECD_plotter

Windows tools for merging instrument ECD exports and analyzing spectral components.
Requires Python 3.12 or newer.

## Setup

Run these commands in PowerShell from the project folder:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## My latest usecase

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --subtract "H2O" --plot --smooth 50 --ht-threshold 500
```

## Merge instrument exports

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Szeged\K.Andrii\260730-PLGA-titration-Zn"
```

Alternatively, edit `DEFAULT_FOLDER` in `merge.py` and run it without `--folder`.

Supported flags for `merge.py` (all optional):

| Flag                     | Purpose                                                                                                                                                            | Default                                             |
| ------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ | --------------------------------------------------- |
| `--folder PATH`          | Folder containing the CSV exports. Quote paths containing spaces.                                                                                                  | `DEFAULT_FOLDER` configured in `merge.py`.          |
| `--refresh-ph`           | Replace existing pH values with values from filenames, preserving concentration coefficients and pH values for filenames without pH.                               | Preserve existing pH values; infer missing values.  |
| `--subtract PREFIX`      | Subtract the CD and Abs reference whose filename starts with `PREFIX`, ignoring case, using Excel formulas. Exactly one spectrum must match.                       | No reference subtraction.                           |
| `--plot`                 | Display corrected CD and Abs spectra in separate panels after saving the workbook. Applies concentration normalization and any requested reference subtraction.    | No plot window.                                     |
| `--smooth [POINTS]`      | Smooth every CD spectrum, including the subtraction reference, before subtraction and concentration normalization. Savitzky-Golay, polynomial order 2, no padding. | No smoothing; `--smooth` alone uses 50 points.      |
| `--ht-threshold [VALUE]` | Blank CD values where that spectrum's HT exceeds `VALUE` V, before smoothing, subtraction, and normalization.                                                      | No HT filtering; `--ht-threshold` alone uses 500 V. |
| `-h`, `--help`           | Display usage and supported flags, then exit.                                                                                                                      | Merge normally.                                     |

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
Extra tabs may be used to align columns in the header and data rows; blank lines
are ignored. Each run writes the table back with one tab between columns.
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

- `CD`: wavelength plus the CD signal for each input file, in mdeg, with optional HT exclusions.
- `HT`: wavelength plus the original HT signal for each input file, in V.
- `Abs`: wavelength plus the original absorbance for each input file.
- `CD_subtr`: wavelength, raw sample CD, a zero separator, then normalized
  sample CD calculated as `sample / concentration_coefficient`, or reference-
  subtracted CD formulas when `--subtract` is supplied.
- `Abs_subtr`: the same layout for absorbance, calculated as
  `sample_abs / concentration_coefficient`, or Excel formulas for
  `(sample_abs - reference_abs) / concentration_coefficient` with `--subtract`.
  The reference and coefficients are the same as for CD. Its measured absorbance
  remains in the input block; the reference is omitted from the corrected block.

With `--smooth`, the input CD block before the zero separator contains the smoothed
signals. Subtraction formulas reference these smoothed values. The `CD` sheet
retains unsmoothed measurements, with any requested HT exclusions; `HT` and `Abs`
retain their original values.

Each channel sheet starts with `wavelength_nm`, followed by one column per spectrum
named after the CSV filename without its extension. The `RAW` sheet is replaced
by these three channel sheets. CD and absorbance are concentration-normalized in
`CD_subtr` and `Abs_subtr`, respectively. Smoothing and HT filtering apply to CD;
absorbance correction uses the original absorbance measurements. Missing sample
or reference absorbance values leave the corresponding corrected cells blank.

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
row. HT exclusions apply to the `CD` sheet as well; `HT` and `Abs` retain their
original signals.

Excel calculates the formulas when the workbook opens. Open and save the workbook
in Excel before reading calculated subtraction values with Python or another tool
that relies on stored formula results.

Add `--plot` to display corrected CD and absorbance spectra in two panels sharing
the wavelength axis in nm. The y-axes show normalized CD in mdeg and normalized
absorbance in AU. Curves and legend entries follow acquisition order, with the same
sample colors in both panels. With `--subtract`, both panels show the subtracted,
concentration-normalized spectra and omit the reference itself. The plot uses the
CSV values directly, so Excel does not need to calculate the workbook first.
Missing absorbance values produce gaps; spectra with no corrected absorbance
values are omitted from that panel. If none are available, only CD is plotted.

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14" --subtract "H2O" --plot
```

The workbook is saved before the plot window opens. Close the window to return to
the terminal.

Use `--smooth 50` to apply Savitzky-Golay smoothing with a 50-point window and
polynomial order 2. `--smooth` without a number also selects 50 points. The window
must contain at least 3 points and cannot exceed the number of wavelength rows.
Both even and odd windows are supported; 50 is used exactly, without rounding.

The boundary condition is **None**: no artificial padding, reflection, or repeated
edge values. Interior fits are evaluated at the original sample positions; an even
50-point window uses 25 preceding and 24 following points. At the edges, a quadratic
is fitted to the first or last full window of measured values and evaluated at each
original position. This preserves the wavelength alignment and output row count.
See [Origin's definition of None](https://docs.originlab.com/origin-help/smooth-dialog/)
and [SciPy's Savitzky-Golay coefficient options](https://docs.scipy.org/doc/scipy/reference/generated/scipy.signal.savgol_coeffs.html).

Smoothing is applied separately to each sample and reference **before subtraction**,
then concentration normalization is applied. The plot uses these same processed
signals, while the Excel subtraction cells remain formulas.

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --smooth 50 --subtract "H2O" --plot
```

Use `--ht-threshold 500` to exclude CD readings whose corresponding HT is **above**
500 V. `--ht-threshold` without a value also selects 500 V; omitting the flag
disables HT filtering. Values equal to the threshold are retained. Thresholds must
be finite and non-negative.

Filtering applies separately to each spectrum, including the subtraction reference.
Affected CD cells are left blank in `CD` and `CD_subtr`; wavelength rows and other
spectra at the same wavelength are retained. If a sample or reference CD cell is
excluded, the corresponding subtraction result is also blank, with no formula
treating the blank as zero. Plots leave gaps at these points. CSV files, HT values,
and absorbance values are preserved.

With `--smooth`, HT filtering happens first. Smoothing operates on each contiguous
valid run, keeping excluded points blank. Runs shorter than the smoothing window
are left unsmoothed; the requested window is not reduced and gaps are not filled.

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --ht-threshold 500 --smooth 50 --subtract "H2O" --plot
```

All input wavelength axes must match the earliest acquired CSV. A mismatch
reports the first differing data row, physical CSV line numbers, expected and
actual wavelengths, reference filename, and row counts. Rerunning replaces the
output workbook. Close it in Excel before rerunning.

## Convert to MRE and estimate secondary structure with SESCA

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --path-in-mm 1 --num-of-aa 20 --concentration-in-mM 0.1
```

This command converts the corrected spectra to MRE and runs the official
`SESCA_deconv.py` engine for each sample. To install SESCA in the default local
directory `.sesca/SESCA_v097`, download and extract the
[official SESCA 0.97 distribution](https://www.mpinat.mpg.de/4787098/SESCA_v097.zip):

```powershell
New-Item -ItemType Directory -Path .sesca -Force | Out-Null
Invoke-WebRequest -Uri "https://www.mpinat.mpg.de/4787098/SESCA_v097.zip" -OutFile ".sesca/SESCA_v097.zip"
Expand-Archive -LiteralPath ".sesca/SESCA_v097.zip" -DestinationPath ".sesca"
```

The distribution is kept outside version control and retains its upstream GPLv3
license. The adapter uses the deconvolution module directly with the NumPy and
SciPy already listed in `requirements.txt`; SESCA's setup script, PDB files, DSSP,
and DISICL are not needed for this workflow. To use another installation, pass
`--sesca-dir "X:\Tools\SESCA_v097"` or set the `SESCA_HOME` environment variable to
its root directory. An explicit flag takes precedence over the environment, which
takes precedence over the local `.sesca/SESCA_v097` directory.

Replace the example residue count and concentration with your sample's values.
All four flags are required. Path length and concentration must be finite and
greater than zero; the amino acid count must be a positive integer.
The concentration is the **peptide/protein molecule concentration in mM before
dilution**, corresponding to a concentration coefficient of `1.0` in `merge.py`.
It is not the metal concentration or the concentration of amino acid residues.
Relative dilution coefficients alone cannot establish absolute concentration.

The script finds `<folder-name>-combined.xlsx` beside the data folder. If the
workbook is missing, the error suggests running `merge.py --folder ...` first.
It reads the corrected spectra **after the zero separator in `CD_subtr`**, keeping
the subtraction, dilution normalization, optional smoothing, and HT exclusions
already applied by `merge.py`. Concentration coefficients are not applied again.
When preparing the combined workbook, use `merge.py --subtract PREFIX` if buffer
subtraction is needed. Only the corrected sample columns are converted; the
subtracted reference is omitted. Without subtraction, all normalized spectra,
including any buffer present as a sample, are converted.

Mean residue ellipticity (MRE), in `deg cm^2 dmol^-1` per residue, is calculated as:

```text
concentration_in_M = concentration_in_mM / 1000
MRE = corrected_CD_mdeg / (path_in_mm * concentration_in_M * num_of_aa)
```

For example, `-20 mdeg` at `1 mm`, `0.1 mM`, and `20` residues gives MRE of
`-10000 deg cm^2 dmol^-1`. See the
[CD conversion protocol](https://pmc.ncbi.nlm.nih.gov/articles/PMC2728378/).

The script evaluates subtraction formulas generated by `merge.py` directly, so
Excel does not need to open or calculate the workbook first. Other formulas and
invalid measurements produce an error before saving. Excluded measurements stay
blank; wavelength order, spectrum order, and existing sheets are preserved.

Results are saved in the same combined workbook:

- `MRE`: wavelength in nm followed by one numeric MRE column per corrected spectrum.
- `MRE_parameters`: path length, residue count, concentration in M, units, formula,
  source sheet, SESCA settings, random seed, and engine/basis file hashes.
- `SESCA_results`: one row per spectrum with structure fractions, intensity scaling
  factor, RMSD in kMRE, normalized RMSD, fitted wavelength range, and point count.
  Fractions are stored as numbers from 0 to 1 and displayed as percentages.
- `SESCA_fit`: one row per fitted wavelength and spectrum, containing the original
  observed MRE, scaled observed MRE, fitted MRE, and residual MRE
  (`scaled_observed_MRE - fitted_MRE`).

The default basis is **DS-dT**, whose three classes are `Alpha`, `Beta`, and `Coil`.
Class names always come from the selected SESCA basis. `Coil` includes structures
outside the alpha and beta classes; it is not a separate measurement of disorder.
Fitting defaults to **190–250 nm** and uses only valid measured wavelengths matching
the basis grid (normally 1 nm). Values at other wavelengths remain in `MRE` but are
not fitted. Wavelengths within `0.000001 nm` of a grid point count as matches.
Excluded points remain absent from fitting, with no interpolation across gaps.
Each sample must have at least `max(10, 2 * number_of_classes)` matched points,
at least 20 nm coverage, and an independent basis matrix. Blank, all-zero, or
insufficient spectra get `status=skipped` with a reason and blank fractions.

The adapter uses SESCA's default nonnegative fit followed by coefficient
normalization, which allows intensity rescaling. The scaling factor multiplies
the observed spectrum before comparison with the normalized fitted spectrum.
For example, `scaling_factor=0.5` means the fitted structure matches half the
measured intensity. Examine this factor alongside RMSD when assessing the fit.
Each fit uses 10 random starts and up to 5000 optimization iterations per start;
seed `0` makes reruns reproducible within the same software environment.
The input passed to SESCA is **kMRE = MRE / 1000**, matching its basis library units.

Optional controls:

| Flag | Purpose | Default |
| --- | --- | --- |
| `--mre-only` | Convert without SESCA; remove any old SESCA output sheets. | Run SESCA. |
| `--sesca-dir PATH` | SESCA root containing `scripts/` and `libs/`. | `SESCA_HOME`, then local installation. |
| `--sesca-basis NAME_OR_PATH` | Standalone basis name or basis file path. | `DS-dT`. |
| `--sesca-min-nm VALUE` | Lower fit wavelength bound. | `190`. |
| `--sesca-max-nm VALUE` | Upper fit wavelength bound. | `250`. |
| `--sesca-repeats INTEGER` | Number of fit restarts. | `10`. |
| `--sesca-seed INTEGER` | Random seed between 0 and 4294967295. | `0`. |

For example, select the six-class `DS5-4` basis or a different fit range:

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --path-in-mm 1 --num-of-aa 20 --concentration-in-mM 0.1 --sesca-basis DS5-4 --sesca-min-nm 195 --sesca-max-nm 250
```

Use standalone basis sets for this workflow. Mixed backbone/side-chain sets need
sequence-dependent correction spectra and are not configured automatically here.
These outputs are SESCA deconvolution estimates; Bayesian estimation and
class-specific uncertainty intervals are not part of this command.
See the [SESCA project and publications](https://www.mpinat.mpg.de/sesca).

Rerunning replaces these four output sheets. A fitting or save error leaves the
original workbook intact. Close the workbook in Excel before
running. Rerunning `merge.py` recreates the entire workbook, so run
`deconvolution.py` again afterward. To retain the earlier conversion-only behavior,
add `--mre-only`.

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
