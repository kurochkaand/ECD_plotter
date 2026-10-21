# ECD_plotter reference

For everyday commands, see the [cheat sheet](../README.md).

Windows tools for merging instrument ECD exports and analyzing spectral components.
Requires Python 3.12 or newer.

## Setup

Run these commands in PowerShell from the project folder:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## My latest usecases

```powershell
python merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --subtract "H2O" --plot --smooth 50 --ht-threshold 500
python deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --plot
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
| `--subtract PREFIX`      | Subtract the CD and absorbance reference whose filename starts with `PREFIX`, ignoring case, before conversion. Exactly one spectrum must match.                  | No reference subtraction.                           |
| `--plot`                 | Display CD in MRE and absorbance in epsilon per residue in separate panels after saving the workbook.                                                            | No plot window.                                     |
| `--smooth [POINTS]`      | Smooth every CD spectrum, including the subtraction reference, before subtraction and concentration normalization. Savitzky-Golay, polynomial order 2, no padding. | No smoothing; `--smooth` alone uses 50 points.      |
| `--ht-threshold [VALUE]` | Blank CD and epsilon where sample or reference HT exceeds `VALUE` V, before smoothing, subtraction, and conversion.                                              | No HT filtering; `--ht-threshold` alone uses 500 V. |
| `-h`, `--help`           | Display usage and supported flags, then exit.                                                                                                                      | Merge normally.                                     |

The script reads selected CSV files directly inside the data folder, ordered by the
acquisition date in cell B5 and time in cell B6, earliest first. Dates use `YY/MM/DD`
(for example, `26/10/14` means October 14, 2026) or `YYYY/MM/DD`. Times can use
`18:41:36` or `6:41:36 PM`. Filename order breaks ties. Missing or invalid timestamps
produce an error identifying the file and cell. Each file must contain an
`XYDATA` block with wavelength, CD, HT, and Abs columns followed by
`##### Extended Information`. Older three-column exports are also accepted, with
blank epsilon values.
When no concentration table exists, every CSV is treated as a sample except the
reference selected by `--subtract`.

On the first run, the script creates `concentration_table.txt` in the data folder
with `pathlength_in_mm = 1`, `number_of_amino_acids = 1`, and an actual
`concentration_in_mM` of `1.0` for each sample. MRE is calculated immediately
using these defaults; edit them to match your samples and rerun. `pH` is extracted
from the filename: `pH7.43` gives `7.43`, and `pH7.9` gives `7.9`. A space after `pH`
is accepted. Files without pH in the name, such as `H2O_1mm_185-300nm_4sc_20nm-1.csv`,
use the default `7.0`.
The table is the source of truth for sample selection, MRE settings, concentrations,
pH, and short names. Once it exists, `merge.py` reads it without rewriting it, preserving
your tabs, blank lines, row order, and other formatting exactly. Extra tabs may
be used to align columns in the header and data rows; blank lines are ignored.
Edit the table and rerun `merge.py` to apply your changes, then rerun
`deconvolution.py` if needed. Workbook columns and plots still follow acquisition
time, regardless of the order of rows in your table.

Delete a sample's row to exclude its CSV from every workbook sheet, plot, and
subsequent deconvolution. New or other unlisted CSVs are ignored without reading
their timestamps or spectral data. Add their rows manually to include them.
Table entries must identify existing CSVs by their full filename stem, without
`.csv`. At least one sample must remain selected. To regenerate defaults for all
current CSVs, delete `concentration_table.txt` and rerun `merge.py`.

The table also includes a `short name` column. Initial defaults are generated by
removing the longest common prefix and suffix from the sample filename stems and
trimming surrounding spaces, underscores, and hyphens. For example,
`peptide_pH7.43_1mm` and `peptide_pH5.77_1mm` become `7.43` and `5.77`.
The reference selected by `--subtract` has no row in a newly created table and
keeps its full filename stem in the workbook inputs. It is loaded separately,
only for subtraction; an existing reference row in an older table is ignored.
A single remaining sample keeps its full stem;
full stems are also retained if shortening would create an empty or duplicate
label or a reserved workbook header (`zero` or `wavelength_nm`). You can edit
short names freely, provided they remain unique and do not conflict with reserved
headers or the reference's filename stem. Two- and three-column sample tables with
the required shared settings remain supported: missing short names and pH values use inferred defaults in
memory, without adding columns or filling cells in the file. Concentrations must
be finite and greater than zero. Existing files missing either shared setting or
using the old `concentration_coefficient` column produce an error with migration
instructions; enter actual concentrations in mM rather than relative coefficients.
`pH` is retained as metadata and is not used in the correction.

The workbook is saved beside the data folder as `<folder-name>-combined.xlsx`:

- `CD`: wavelength plus the CD signal for each input file, in mdeg, with optional HT exclusions.
- `HT`: wavelength plus the original HT signal for each input file, in V.
- `CD_subtr`: wavelength, raw sample CD, a zero separator, then normalized
  sample CD calculated as `sample / concentration_in_mM`, or reference-
  subtracted CD formulas when `--subtract` is supplied.
- `MRE`: numeric mean residue ellipticity for each selected sample.
- `MRE_parameters`: shared settings, each sample concentration in mM and M,
  MRE and epsilon units and conversion details, the absorbance reference, and HT cutoff.
- `Epsilon`: wavelength plus numeric absorption coefficients in `M^-1 cm^-1`
  per residue for each selected sample, after any reference subtraction and HT
  exclusions. The subtraction reference is omitted. Raw absorbance and the old
  `Abs` and `Abs_subtr` sheets are not saved.

With `--smooth`, the input CD block before the zero separator contains the smoothed
signals. Subtraction formulas reference these smoothed values. The `CD` sheet
retains unsmoothed measurements, with any requested HT exclusions; `HT`
retains its original values. Epsilon uses unsmoothed absorbance after HT exclusions.

Each channel sheet starts with `wavelength_nm`, followed by one column per spectrum
named using the table's `short name`. The input and corrected blocks also use
these short names, as do the merge plot legends. `CD_subtr` contains CD normalized
by concentration in mdeg/mM. `Epsilon` also normalizes absorbance by path length
and residue count. Smoothing applies only to CD; HT filtering applies to both
channels. Missing or excluded sample or reference absorbance values leave the
corresponding epsilon cells blank.

To subtract a reference spectrum, provide its filename prefix:

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14" --subtract "H2O"
```

The prefix must match exactly one spectrum, ignoring case. The reference remains
in the raw columns of `CD_subtr`; after the zero separator, each other spectrum
gets Excel formulas for `(sample - reference) / concentration_in_mM`.
The reference itself is omitted from the corrected block. With a concentration of
`1.0 mM`, a formula might be `=B2-$O2`; with a concentration of `2.0 mM`, it becomes
`=(B2-$O2)/2.0`. The reference column is fixed, and each wavelength uses its own
row. HT exclusions apply to `CD`, `CD_subtr`, `MRE`, and `Epsilon`; `HT` retains
its original signals. Epsilon is computed numerically from the filtered input
absorbance, with the same reference and sample concentrations as CD.

Excel calculates the `CD_subtr` formulas when the workbook opens. Open and save the workbook
in Excel before reading calculated subtraction values with Python or another tool
that relies on stored formula results.

Add `--plot` to display corrected CD and absorbance spectra in two panels sharing
the wavelength axis in nm. The CD panel shows MRE using the settings in
`concentration_table.txt`. Absorbance is shown as epsilon in `M^-1 cm^-1` per residue,
matching the numeric `Epsilon` sheet. Curves and legend entries follow acquisition order, with the same
sample colors in both panels. With `--subtract`, both panels show the subtracted,
converted spectra and omit the reference itself. The plot uses the
CSV values directly, so Excel does not need to calculate the workbook first.
Missing or HT-excluded absorbance values produce gaps; spectra with no valid epsilon
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

Use `--ht-threshold 500` to exclude CD and epsilon readings whose corresponding HT is **above**
500 V. `--ht-threshold` without a value also selects 500 V; omitting the flag
disables HT filtering. Values equal to the threshold are retained. Thresholds must
be finite and non-negative.

Filtering applies separately to each spectrum, including the subtraction reference.
Affected cells are left blank in `CD`, `CD_subtr`, `MRE`, and `Epsilon`; wavelength
rows and other spectra at the same wavelength are retained. If a sample or
reference reading is excluded, the corresponding corrected result is also blank,
with no formula treating the blank as zero. Plots leave gaps at these points.
CSV files and exported HT values are preserved.

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

## Convert corrected CD to MRE during merge

```powershell
.\.venv\Scripts\python.exe merge.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --subtract "H2O"
```

MRE conversion runs on every merge after baseline subtraction, concentration
normalization, optional smoothing, and HT exclusions. All conversion settings come
from `concentration_table.txt`; no conversion flags are needed or accepted.

```text
pathlength_in_mm = 1
number_of_amino_acids = 20

spectrum	concentration_in_mM	pH	short name
peptide	0.1	7.0	Peptide
```

New files default to `1 mm`, `1` residue per molecule, and `1.0 mM` per sample.
Path length and sample concentrations must be finite and greater than zero; the
amino acid count must be a positive integer. Both shared settings are required in
existing files. Missing settings, duplicate or unknown settings, invalid values,
and the old concentration column header cause errors without rewriting the file.

Each concentration is the **peptide/protein molecule concentration in that measured
sample**, in mM. It is not a relative dilution coefficient, metal concentration,
or amino acid residue concentration. All samples share the path length and
amino acid count. Edit these values and rerun merge to recalculate the workbook.

Only the corrected sample columns after the zero separator in `CD_subtr` are
converted. The subtraction reference is omitted. Without `--subtract`, all
normalized spectra selected in the concentration table are converted.

Mean residue ellipticity (MRE), in `deg cm^2 dmol^-1` per residue, is calculated as:

```text
concentration_in_M = concentration_in_mM / 1000
MRE = baseline_corrected_CD_mdeg / (pathlength_in_mm * concentration_in_M * number_of_amino_acids)

# CD_subtr is already divided by concentration_in_mM:
MRE = CD_subtr * 1000 / (pathlength_in_mm * number_of_amino_acids)
```

For example, `-20 mdeg` at `1 mm`, `0.1 mM`, and `20` residues gives MRE of
`-10000 deg cm^2 dmol^-1`. See the
[CD conversion protocol](https://pmc.ncbi.nlm.nih.gov/articles/PMC2728378/).

Merge evaluates its subtraction formulas directly to produce numeric MRE, so
Excel does not need to calculate the workbook first. Excluded measurements stay
blank, and wavelength order and spectrum labels are retained. Conversion errors
leave an existing combined workbook intact.

Alongside the measured and corrected sheets, merge writes:

- `MRE`: wavelength in nm followed by one numeric MRE column per corrected spectrum.
- `MRE_parameters`: path length, residue count, concentration in mM and M, units,
  formulas, source sheets, each sample concentration indexed by short name, and
  epsilon reference and HT cutoff metadata.

MRE conversion does not require SESCA. To convert without fitting, run only merge;
merge's `--plot` always shows CD in MRE (`deg cm^2 dmol^-1` per residue),
matching the saved `MRE` sheet. Absorbance is plotted as epsilon per residue,
matching the saved `Epsilon` sheet.

## Convert absorbance to epsilon per residue during merge

Every merge converts absorbance using the same path length, residue count, sample
concentrations, and optional subtraction reference as CD. No extra flag is needed.
The Beer-Lambert relation gives the per-residue molar absorption coefficient:

```text
A_corrected = sample_abs - reference_abs  # reference_abs = 0 without --subtract
epsilon_per_residue = A_corrected / (concentration_in_M * pathlength_in_cm * number_of_amino_acids)

# For the concentration table's mM and mm units:
epsilon_per_residue = A_corrected * 10000 / (concentration_in_mM * pathlength_in_mm * number_of_amino_acids)
```

For example, corrected absorbance of `0.5` at `0.1 mM`, `1 mm`, and `20` residues
gives epsilon of `2500 M^-1 cm^-1` per residue. Concentrations must be molecule
concentrations; supplying residue concentrations would normalize by residue count
twice. This is ordinary absorption epsilon, not the CD difference `delta epsilon`.
See the [IUPAC Beer-Lambert definition](https://goldbook.iupac.org/terms/view/B00626).

Only corrected numeric values are saved in `Epsilon`, without a raw input block,
zero separator, or subtraction reference column. Excel recalculation is unnecessary
for this sheet. `Abs` and `Abs_subtr` are no longer generated. Conversion settings
and units are recorded alongside the shared parameters in `MRE_parameters`.

`--ht-threshold` masks both sample and reference absorbance before subtraction.
Missing or excluded values produce blank epsilon cells and gaps in the plot;
fully blank curves are omitted, and the epsilon panel is omitted when all its
curves are blank. `--smooth` continues to affect only CD. Negative corrected
absorbance produces negative epsilon without clipping. Invalid numeric values or
nonfinite conversion results leave an existing output workbook intact.

## Estimate secondary structure with SESCA

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd"
```

Deconvolution reads the numeric `MRE` sheet from `<folder-name>-combined.xlsx`
beside the data folder and runs the official `SESCA_deconv.py` engine for each
sample. It does not apply dilution normalization or MRE conversion again.
A missing workbook or `MRE` sheet produces an error suggesting a rerun of merge
using the settings in `concentration_table.txt`. Existing MRE values and conversion
metadata are preserved. `deconvolution.py` requires only `--folder`, with optional
fitting and plotting settings.

To install SESCA in the default local directory `.sesca/SESCA_v097`, download and extract the
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

Deconvolution writes these sheets in the same combined workbook:

- `SESCA_results`: one row per spectrum with structure fractions, intensity scaling
  factor, RMSD in kMRE, normalized RMSD, fitted wavelength range, and point count.
  Fractions are stored as numbers from 0 to 1 and displayed as percentages.
- `SESCA_fit`: one row per fitted wavelength and spectrum, containing the original
  observed MRE, scaled observed MRE, fitted MRE, and residual MRE
  (`scaled_observed_MRE - fitted_MRE`).

SESCA settings, random seed, and engine/basis file hashes are added to
`MRE_parameters`. Reruns replace fitting metadata and retain conversion settings.

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

Add `--shared-scaling` to determine **one automatic scaling factor for the entire
series**. Each eligible spectrum keeps its own structure composition. The
compositions and common factor are optimized jointly; individual scaling factors
are used only to initialize the fit, not averaged to produce the result.

To supply your own common factor, give it as the optional value:

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --shared-scaling 1.33334 --plot
```

Here `1.33334` is held fixed for every fitted spectrum. Structure fractions are
optimized with that factor; the factor itself is not optimized. It multiplies
the observed MRE, so `scaled_observed_MRE = observed_MRE * 1.33334`. The value must
be finite and greater than zero. Omitting the flag keeps independent scaling;
`--shared-scaling` without a value keeps automatic shared scaling.

The joint fit minimizes the sum of squared spectral residuals in the original
unscaled intensities over all valid fitted points. It uses the selected SESCA
basis and nonnegative coefficients, with every spectrum's coefficient sum
constrained to the same value `q`. The shared scaling factor is `1 / q`, and
each spectrum's fractions are its coefficients divided by `q`. This retains
SESCA's least-squares criterion while enforcing a common scale. Every included
wavelength contributes equally, so spectra with more valid fitted points
contribute more to the combined objective. With a supplied factor, `q` is fixed
to its reciprocal while the structure coefficients are fitted.

The same `scaling_factor` is written to every fitted row of `SESCA_results`.
Fractions, fitted curves, RMSD, and NRMSD are recalculated for the shared fit;
`SESCA_fit` retains the same scaled-observation and residual conventions.
`MRE_parameters` records `sesca_scaling_mode`, `sesca_shared_scaling_factor`,
`sesca_shared_scaling_source` (`automatic` or `fixed`), the solver status, and
the combined RMSD in unscaled kMRE. Skipped spectra do not participate and keep
blank fractions and scaling factors. If all spectra are skipped, no automatic
factor is reported; an explicitly supplied factor is still recorded in the
parameters sheet.

Shared scaling is implemented in this adapter with joint constrained least
squares using [SciPy SLSQP](https://docs.scipy.org/doc/scipy/reference/optimize.minimize-slsqp.html).
Without the flag, SESCA continues to determine a separate factor per spectrum.
`--shared-scaling` can be combined with `--plot`.

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --shared-scaling --plot
```

Optional controls:

| Flag                         | Purpose                                                             | Default                                |
| ---------------------------- | ------------------------------------------------------------------- | -------------------------------------- |
| `--plot`                     | Display structure percentages from the saved `SESCA_results` sheet. | No plot window.                        |
| `--shared-scaling [FACTOR]`  | Use one common factor: optimize automatically without a value, or hold FACTOR fixed. | Separate factor per spectrum.       |
| `--sesca-dir PATH`           | SESCA root containing `scripts/` and `libs/`.                       | `SESCA_HOME`, then local installation. |
| `--sesca-basis NAME_OR_PATH` | Standalone basis name or basis file path.                           | `DS-dT`.                               |
| `--sesca-min-nm VALUE`       | Lower fit wavelength bound.                                         | `190`.                                 |
| `--sesca-max-nm VALUE`       | Upper fit wavelength bound.                                         | `250`.                                 |
| `--sesca-repeats INTEGER`    | Number of fit restarts.                                             | `10`.                                  |
| `--sesca-seed INTEGER`       | Random seed between 0 and 4294967295.                               | `0`.                                   |

Add `--plot` to display a horizontal stacked bar chart with one bar per spectrum
and a color for each structure class in the selected basis. Spectra follow their
workbook order; skipped fits are labeled `Skipped` without a composition bar.
MRE headers, SESCA spectrum names, and plot labels inherit the short names from
the corrected workbook headers generated by `merge.py`.
The workbook is saved before the plot opens. Close the plot window to return to
the terminal. If every fit is skipped, the command reports that no plot is
available.

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --plot
```

For example, select the six-class `DS5-4` basis or a different fit range:

```powershell
.\.venv\Scripts\python.exe deconvolution.py --folder "X:\Cations_and_PGA\ECD-Prague\2026-10-14-Cd" --sesca-basis DS5-4 --sesca-min-nm 195 --sesca-max-nm 250
```

Use standalone basis sets for this workflow. Mixed backbone/side-chain sets need
sequence-dependent correction spectra and are not configured automatically here.
These outputs are SESCA deconvolution estimates; Bayesian estimation and
class-specific uncertainty intervals are not part of this command.
See the [SESCA project and publications](https://www.mpinat.mpg.de/sesca).

Rerunning replaces the two SESCA output sheets and fitting metadata. A fitting or
save error leaves the original workbook intact. Close the workbook in Excel before
running. Rerunning `merge.py` recreates the entire workbook, so run
`deconvolution.py` again afterward if fitting is needed.

## Analyze spectral components

```powershell
.\.venv\Scripts\python.exe spectral_component_analysis.py --file "X:\Data\corrected.xlsx" --components 2 --no-show
```

See [the analysis guide](SPECTRAL_COMPONENT_ANALYSIS.md) for input layout,
configuration, outputs, and scientific limitations.

## Check the code

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```
