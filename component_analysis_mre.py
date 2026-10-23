"""Estimate spectral rank with SVD of the MRE sheet prepared by merge.py."""

from __future__ import annotations

import argparse
import os
import sys
import tempfile
from dataclasses import dataclass
from math import isfinite
from pathlib import Path
from zipfile import BadZipFile

import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from deconvolution import read_mre
from merge import MRE_SHEET, MRE_UNITS
from spectral_component_analysis import (
    CUMULATIVE_VARIANCE_TARGET, SVDResult, analyze_singular_values,
)


RESULTS_SHEET = "MRE_SVD_results"
PARAMETERS_SHEET = "MRE_SVD_parameters"
INPUT_SHEET = "MRE_SVD_input"
DEFAULT_MIN_NM = 190.0
DEFAULT_MAX_NM = 250.0


@dataclass(frozen=True)
class MreSvdAnalysis:
    names: list[str]
    wavelengths: np.ndarray
    spectra: np.ndarray  # (spectra, retained wavelength points), in saved MRE units
    svd: SVDResult
    min_nm: float
    max_nm: float
    points_in_range: int
    points_outside_range: int

    @property
    def excluded_points(self) -> int:
        return self.points_in_range - len(self.wavelengths)


def analyze_mre(
    sheet: Worksheet, *, min_nm: float = DEFAULT_MIN_NM, max_nm: float = DEFAULT_MAX_NM,
) -> MreSvdAnalysis:
    """Keep complete shared wavelengths; never interpolate or renormalize MRE."""
    if not isfinite(min_nm) or not isfinite(max_nm) or not 0 < min_nm < max_nm:
        raise ValueError("Wavelength bounds must be finite, positive, and increasing.")

    names, rows = read_mre(sheet)
    if len(names) < 2:
        raise ValueError("MRE component analysis requires at least two spectra.")
    values = np.asarray(rows, dtype=float)
    wavelengths = values[:, 0]
    if len(np.unique(wavelengths)) != len(wavelengths):
        raise ValueError("MRE wavelengths must be unique. Rerun merge.py to regenerate the workbook.")

    in_range = (wavelengths >= min_nm) & (wavelengths <= max_nm)
    points_in_range = int(np.count_nonzero(in_range))
    if not points_in_range:
        raise ValueError(
            f"No MRE wavelengths in {min_nm:g}-{max_nm:g} nm; "
            f"the workbook spans {wavelengths.min():g}-{wavelengths.max():g} nm."
        )
    # read_mre rejects formulas, text, errors and infinities; only blanks become NaN.
    complete = in_range & np.isfinite(values[:, 1:]).all(axis=1)
    retained = values[complete]
    if len(retained) < 5:
        raise ValueError(
            f"Only {len(retained)} complete MRE wavelength rows remain in "
            f"{min_nm:g}-{max_nm:g} nm after excluding "
            f"{points_in_range - len(retained)} rows with missing values; need at least 5. "
            "Check HT exclusions or choose a different wavelength range."
        )
    retained = retained[np.argsort(retained[:, 0])]
    spectra = retained[:, 1:].T
    # Override the older script's editable setting: MRE rank uses uncentered data.
    svd = analyze_singular_values(spectra, mean_center=False)
    return MreSvdAnalysis(
        names=names, wavelengths=retained[:, 0], spectra=spectra, svd=svd,
        min_nm=min_nm, max_nm=max_nm, points_in_range=points_in_range,
        points_outside_range=int(len(values) - points_in_range),
    )


def write_results(workbook: Workbook, analysis: MreSvdAnalysis) -> None:
    """Replace only SVD outputs; retain the source spectra and other analyses."""
    for title in (RESULTS_SHEET, PARAMETERS_SHEET, INPUT_SHEET):
        if title in workbook.sheetnames:
            del workbook[title]

    results = workbook.create_sheet(RESULTS_SHEET)
    results.append([
        "component", "singular_value_MRE", "spectral_energy_fraction",
        "cumulative_energy_fraction", "above_noise_threshold",
    ])
    svd = analysis.svd
    for index, singular_value in enumerate(svd.singular_values):
        results.append([
            index + 1, float(singular_value), float(svd.explained_energy_ratio[index]),
            float(svd.cumulative_energy_ratio[index]),
            bool(singular_value > svd.gavish_donoho_threshold),
        ])
    for row in results.iter_rows(min_row=2):
        row[1].number_format = "0.000000E+00"
        row[2].number_format = row[3].number_format = "0.00%"

    parameters = workbook.create_sheet(PARAMETERS_SHEET)
    parameters.append(["parameter", "value"])
    for key, value in (
        ("analysis", "Singular value decomposition (SVD)"),
        ("source_sheet", MRE_SHEET),
        ("MRE_units", MRE_UNITS),
        ("matrix_orientation", "Rows: spectra; columns: retained wavelengths"),
        ("mean_centered", False),
        ("normalization", "None; use saved numeric MRE without further conversion"),
        ("spectra_count", len(analysis.names)),
        ("requested_min_nm", analysis.min_nm),
        ("requested_max_nm", analysis.max_nm),
        ("retained_min_nm", float(analysis.wavelengths.min())),
        ("retained_max_nm", float(analysis.wavelengths.max())),
        ("wavelength_rows_in_range", analysis.points_in_range),
        ("retained_wavelength_rows", len(analysis.wavelengths)),
        ("excluded_missing_rows", analysis.excluded_points),
        ("outside_range_rows", analysis.points_outside_range),
        ("missing_value_policy", "Exclude wavelength row if any spectrum has a blank; no interpolation"),
        ("gavish_donoho_rank", svd.gavish_donoho_rank),
        ("gavish_donoho_threshold_MRE", svd.gavish_donoho_threshold),
        ("noise_threshold_method", "Gavish-Donoho unknown-noise threshold; polynomial approximation"),
        ("spectral_energy_target", CUMULATIVE_VARIANCE_TARGET),
        ("spectral_energy_rank", svd.variance_rank),
        ("component_limit", "None; report every singular value"),
        ("interpretation", "Estimated independent spectral patterns; not a chemical species count"),
    ):
        parameters.append([key, value])
        if key == "gavish_donoho_threshold_MRE":
            parameters.cell(parameters.max_row, 2).number_format = "0.000000E+00"
        elif key == "spectral_energy_target":
            parameters.cell(parameters.max_row, 2).number_format = "0.00%"

    inputs = workbook.create_sheet(INPUT_SHEET)
    inputs.append(["wavelength_nm", *analysis.names])
    for wavelength, signals in zip(analysis.wavelengths, analysis.spectra.T):
        inputs.append([float(wavelength), *[float(value) for value in signals]])
    for row in inputs.iter_rows(min_row=2, min_col=2):
        for cell in row:
            cell.number_format = "0.000000"

    for sheet in (results, parameters, inputs):
        sheet.freeze_panes = "B2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.row_dimensions[1].height = 36
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for column in range(1, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(column)].width = 28
    inputs.column_dimensions["A"].width = 18
    parameters.column_dimensions["A"].width = 36
    parameters.column_dimensions["B"].width = 90
    for row in parameters.iter_rows(min_row=2, min_col=2):
        row[0].alignment = Alignment(wrap_text=True, vertical="top")
        parameters.row_dimensions[row[0].row].height = 32


def plot_svd(analysis: MreSvdAnalysis) -> None:
    """Display the full singular-value spectrum and explained spectral energy."""
    import matplotlib.pyplot as plt
    from matplotlib.ticker import MaxNLocator

    svd = analysis.svd
    components = np.arange(1, len(svd.singular_values) + 1)
    figure, (values_axis, energy_axis) = plt.subplots(1, 2, figsize=(13, 5))
    try:
        values_axis.semilogy(components, svd.singular_values, "o-")
        if svd.gavish_donoho_threshold > 0:
            values_axis.axhline(
                svd.gavish_donoho_threshold, linestyle="--", label="Noise threshold",
            )
        values_axis.axvline(
            svd.gavish_donoho_rank, linestyle=":", label=f"Noise rank: {svd.gavish_donoho_rank}",
        )
        values_axis.set_title("MRE singular values")
        values_axis.set_ylabel("Singular value (MRE units, log scale)")
        values_axis.legend()

        energy_axis.bar(components, 100 * svd.explained_energy_ratio, label="Individual")
        energy_axis.plot(components, 100 * svd.cumulative_energy_ratio, "o-", label="Cumulative")
        energy_axis.axhline(
            100 * CUMULATIVE_VARIANCE_TARGET, linestyle="--",
            label=f"{100 * CUMULATIVE_VARIANCE_TARGET:g}% energy target",
        )
        energy_axis.axvline(svd.variance_rank, linestyle=":", label=f"Energy rank: {svd.variance_rank}")
        energy_axis.set_title("Explained spectral energy")
        energy_axis.set_ylabel("Spectral energy (%)")
        energy_axis.set_ylim(0, 102)
        energy_axis.legend()
        for axis in (values_axis, energy_axis):
            axis.set_xlabel("Component number")
            axis.xaxis.set_major_locator(MaxNLocator(integer=True))
            axis.grid(alpha=0.25)
        figure.suptitle(
            f"{len(analysis.names)} MRE spectra, "
            f"{analysis.wavelengths.min():g}-{analysis.wavelengths.max():g} nm, "
            f"{len(analysis.wavelengths)} complete wavelength rows"
        )
        figure.tight_layout()
        plt.show(block=True)
    finally:
        plt.close(figure)


def analyze_mre_components(
    folder: Path, *, min_nm: float = DEFAULT_MIN_NM, max_nm: float = DEFAULT_MAX_NM,
    plot: bool = False,
) -> Path:
    """Analyze the sibling combined workbook and atomically save SVD sheets."""
    folder = Path(folder).resolve()
    workbook_file = folder.parent / f"{folder.name}-combined.xlsx"
    merge_command = f'python merge.py --folder "{folder}"'
    if not workbook_file.is_file():
        raise FileNotFoundError(
            f"Combined workbook not found: {workbook_file}\nRun merge.py first: {merge_command}"
        )

    workbook = load_workbook(workbook_file, data_only=False)
    temporary_path = None
    try:
        if MRE_SHEET not in workbook.sheetnames:
            raise ValueError(
                f"{workbook_file} has no {MRE_SHEET} sheet. Run merge.py first: {merge_command}"
            )
        analysis = analyze_mre(workbook[MRE_SHEET], min_nm=min_nm, max_nm=max_nm)
        write_results(workbook, analysis)
        # Preserve the source if analysis, saving, or replacement fails.
        with tempfile.NamedTemporaryFile(
            prefix=f".{workbook_file.stem}-", suffix=".xlsx", dir=workbook_file.parent, delete=False,
        ) as temporary:
            temporary_path = Path(temporary.name)
        workbook.save(temporary_path)
        workbook.close()
        os.replace(temporary_path, workbook_file)
    finally:
        workbook.close()
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)

    print(f"Saved {RESULTS_SHEET}, {PARAMETERS_SHEET}, and {INPUT_SHEET} in {workbook_file}.")
    print(f"Spectra: {len(analysis.names)} (saved MRE; no centering or further normalization).")
    print(
        f"Wavelength rows: {len(analysis.wavelengths)} retained, "
        f"{analysis.excluded_points} excluded for missing values, "
        f"{analysis.points_outside_range} outside the requested range."
    )
    print(f"Retained range: {analysis.wavelengths.min():g}-{analysis.wavelengths.max():g} nm.")
    print(f"Gavish-Donoho rank: {analysis.svd.gavish_donoho_rank}.")
    print(f"Rank at {100 * CUMULATIVE_VARIANCE_TARGET:g}% spectral energy: {analysis.svd.variance_rank}.")
    print("Ranks estimate independent spectral patterns, not the number of chemical species.")
    if plot:
        plot_svd(analysis)
    return workbook_file


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Estimate spectral rank with SVD of the MRE sheet generated by merge.py."
    )
    parser.add_argument("--folder", type=Path, required=True, help="Data folder passed to merge.py.")
    parser.add_argument(
        "--min-wavelength", dest="min_nm", type=float, default=DEFAULT_MIN_NM,
        help="Lower wavelength bound, inclusive (default: 190 nm).",
    )
    parser.add_argument(
        "--max-wavelength", dest="max_nm", type=float, default=DEFAULT_MAX_NM,
        help="Upper wavelength bound, inclusive (default: 250 nm).",
    )
    parser.add_argument("--plot", action="store_true", help="Display SVD diagnostics after saving the workbook.")
    args = parser.parse_args()
    try:
        analyze_mre_components(args.folder, min_nm=args.min_nm, max_nm=args.max_nm, plot=args.plot)
    except (OSError, ValueError, RuntimeError, BadZipFile, np.linalg.LinAlgError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
