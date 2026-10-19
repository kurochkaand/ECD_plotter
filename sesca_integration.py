"""Run the official SESCA_deconv engine on corrected MRE spectra."""

from __future__ import annotations

import importlib.util
import io
import os
import random
import tempfile
import warnings
from contextlib import redirect_stdout
from copy import deepcopy
from dataclasses import dataclass, field
from hashlib import sha256
from pathlib import Path

import numpy as np
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from scipy.optimize import OptimizeWarning


SESCA_URL = "https://www.mpinat.mpg.de/sesca"
DEFAULT_SESCA_DIR = Path(__file__).resolve().parent / ".sesca" / "SESCA_v097"
RESULTS_SHEET = "SESCA_results"
FIT_SHEET = "SESCA_fit"


@dataclass(frozen=True)
class SescaOptions:
    directory: Path | None = None
    basis: str = "DS-dT"
    min_nm: float = 190.0
    max_nm: float = 250.0
    repeats: int = 10
    seed: int = 0


@dataclass
class SpectrumFit:
    name: str
    status: str
    note: str = ""
    fractions: list[float] = field(default_factory=list)
    wavelengths: list[float] = field(default_factory=list)
    observed_mre: list[float] = field(default_factory=list)
    fitted_mre: list[float] = field(default_factory=list)
    scale: float | None = None
    rmsd_kmre: float | None = None
    nrmsd: float | None = None


@dataclass
class SescaAnalysis:
    classes: list[str]
    fits: list[SpectrumFit]
    parameters: list[tuple[str, str | int | float]]


def resolve_sesca(options: SescaOptions) -> tuple[Path, Path]:
    """Find an explicit installation, SESCA_HOME, or the local distribution."""
    if not np.isfinite([options.min_nm, options.max_nm]).all() or not 0 < options.min_nm < options.max_nm:
        raise ValueError("SESCA wavelength bounds must be finite, positive, and increasing.")
    if isinstance(options.repeats, bool) or not isinstance(options.repeats, int) or options.repeats < 1:
        raise ValueError("SESCA repeats must be a positive integer.")
    if isinstance(options.seed, bool) or not isinstance(options.seed, int) or not 0 <= options.seed < 2**32:
        raise ValueError("SESCA seed must be an integer between 0 and 4294967295.")
    directory = Path(options.directory or os.environ.get("SESCA_HOME") or DEFAULT_SESCA_DIR).resolve()
    if not (directory / "scripts" / "SESCA_deconv.py").is_file():
        raise FileNotFoundError(
            f"SESCA_deconv.py not found in {directory / 'scripts'}. "
            f"Download and extract SESCA from {SESCA_URL} to {DEFAULT_SESCA_DIR}, "
            "or pass --sesca-dir PATH to its root directory. "
            "Use --mre-only to convert without fitting."
        )
    basis = Path(options.basis)
    if not basis.is_file():
        basis = directory / "libs" / f"Map_{options.basis}.dat"
    if not basis.is_file():
        raise FileNotFoundError(
            f"SESCA basis set not found: {options.basis!r}. "
            "Use a standalone basis name such as DS-dT, DS5-4, or DSSP-1, "
            "or pass the path to a SESCA basis file with --sesca-basis."
        )
    return directory, basis.resolve()


def read_basis(basis_file: Path, options: SescaOptions) -> tuple[list[str], np.ndarray]:
    """Read SESCA's labeled basis table, retaining its native wavelength grid."""
    classes = []
    rows = []
    for line in basis_file.read_text(encoding="ascii").splitlines():
        if line.startswith("#|"):
            classes = line[2:].split()[1:]
        elif line.strip() and not line.lstrip().startswith("#"):
            try:
                row = [float(value) for value in line.split()]
            except ValueError as exc:
                raise ValueError(f"Invalid numeric basis data in {basis_file}.") from exc
            if row and options.min_nm <= row[0] <= options.max_nm:
                rows.append(row)
    if not classes or not rows or any(len(row) != len(classes) + 1 for row in rows):
        raise ValueError(f"No valid SESCA basis table in the selected wavelength range: {basis_file}")
    values = np.asarray(rows, dtype=float)
    if not np.isfinite(values).all() or len(set(values[:, 0])) != len(values):
        raise ValueError(f"SESCA basis data must be finite with unique wavelengths: {basis_file}")
    return classes, values


def load_engine(directory: Path):
    """Import the official engine without changing the upstream installation."""
    script = directory / "scripts" / "SESCA_deconv.py"
    spec = importlib.util.spec_from_file_location("_ecd_sesca_deconv", script)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load SESCA from {script}.")
    module = importlib.util.module_from_spec(spec)
    try:
        with redirect_stdout(io.StringIO()):
            spec.loader.exec_module(module)
    except (ImportError, SyntaxError) as exc:
        raise RuntimeError(f"Cannot load SESCA: {exc}") from exc
    return module


def fit_spectra(
    names: list[str], mre_rows: list[list[float | None]], options: SescaOptions
) -> SescaAnalysis:
    """Fit valid basis-grid points only; never interpolate over excluded CD data."""
    directory, basis_file = resolve_sesca(options)
    classes, basis = read_basis(basis_file, options)
    engine = load_engine(directory)
    defaults = deepcopy(engine.Pass_Defaults())
    wavelengths = np.asarray([row[0] for row in mre_rows], dtype=float)
    if not np.isfinite(wavelengths).all() or len(set(wavelengths)) != len(wavelengths):
        raise ValueError("SESCA input must have finite, unique wavelengths.")
    matches = []
    for basis_row in basis:
        positions = np.flatnonzero(np.isclose(wavelengths, basis_row[0], rtol=0, atol=1e-6))
        if len(positions) > 1:
            raise ValueError(f"Multiple input wavelengths match SESCA basis wavelength {basis_row[0]:g} nm.")
        if len(positions):
            matches.append((int(positions[0]), basis_row))

    fits = []
    with tempfile.TemporaryDirectory(prefix="ECD SESCA ") as temporary:
        spectrum_file = Path(temporary) / "spectrum.dat"
        for column, name in enumerate(names, start=1):
            selected = [(index, row) for index, row in matches if mre_rows[index][column] is not None]
            minimum_points = max(10, 2 * len(classes))
            if len(selected) < minimum_points:
                fits.append(SpectrumFit(name, "skipped", f"Only {len(selected)} valid basis-grid points; need {minimum_points}."))
                continue
            matrix = np.asarray([row[1:] for _, row in selected])
            fit_waves = np.asarray([row[0] for _, row in selected])
            observed = np.asarray([mre_rows[index][column] for index, _ in selected], dtype=float)
            if not np.isfinite(observed).all():
                raise ValueError(f"Nonfinite MRE input for SESCA spectrum {name!r}.")
            if np.ptp(fit_waves) < 20 or np.linalg.matrix_rank(matrix) < len(classes):
                fits.append(SpectrumFit(name, "skipped", "Need at least 20 nm coverage and independent basis spectra."))
                continue
            if not np.any(observed):
                fits.append(SpectrumFit(name, "skipped", "All selected MRE values are zero."))
                continue
            # SESCA and its libraries use kMRE, rather than the workbook's MRE.
            np.savetxt(spectrum_file, np.column_stack([fit_waves, observed / 1000.0]), fmt="%.12g")
            log = io.StringIO()
            python_state, numpy_state = random.getstate(), np.random.get_state()
            try:
                random.seed(options.seed)
                np.random.seed(options.seed)
                # Read_Args mutates SESCA's defaults. Reset them for each spectrum.
                engine.Def_Args = deepcopy(defaults)
                with redirect_stdout(log), warnings.catch_warnings():
                    # SESCA 0.97 passes a retired SciPy option; modern defaults apply.
                    warnings.filterwarnings("ignore", message="Unknown solver options: ftol", category=OptimizeWarning)
                    arguments = engine.Read_Args([
                        "@spect", str(spectrum_file), "@lib", str(basis_file),
                        "@write", "0", "@range", f"{options.min_nm},{options.max_nm}",
                        "@rep", str(options.repeats), "@iter", "5000", "@err", "0", "@verb", "0",
                    ])
                    result = engine.Deconv_Main(arguments)
            except (Exception, SystemExit) as exc:
                raise RuntimeError(f"SESCA failed for {name!r}: {exc}\n{log.getvalue()[-1500:]}") from exc
            finally:
                random.setstate(python_state)
                np.random.set_state(numpy_state)
            fractions = np.asarray(result[0], dtype=float)
            rmsd, _, nrmsd, scale = map(float, result[1:5])
            if (fractions.shape != (len(classes),) or not np.isfinite(fractions).all()
                    or np.any(fractions < 0) or not np.isclose(fractions.sum(), 1, atol=1e-6)
                    or not np.isfinite([rmsd, nrmsd, scale]).all() or min(rmsd, nrmsd) < 0 or scale <= 0):
                raise RuntimeError(f"SESCA returned invalid fit values for {name!r}.")
            fitted = matrix @ fractions * 1000.0
            expected_rmsd = float(np.sqrt(np.mean(((observed * scale - fitted) / 1000.0) ** 2)))
            if not np.isclose(rmsd, expected_rmsd, rtol=1e-5, atol=1e-7):
                raise RuntimeError(f"SESCA fitted spectrum and reported RMSD disagree for {name!r}.")
            fits.append(SpectrumFit(
                name, "ok", fractions=fractions.tolist(), wavelengths=fit_waves.tolist(),
                observed_mre=observed.tolist(), fitted_mre=fitted.tolist(),
                scale=scale, rmsd_kmre=rmsd, nrmsd=nrmsd,
            ))
    return SescaAnalysis(classes, fits, [
        ("secondary_structure_analysis", "SESCA_deconv"),
        ("sesca_directory", str(directory)),
        ("sesca_source", SESCA_URL),
        ("sesca_basis_file", str(basis_file)),
        ("sesca_basis_sha256", sha256(basis_file.read_bytes()).hexdigest()),
        ("sesca_script_sha256", sha256((directory / "scripts" / "SESCA_deconv.py").read_bytes()).hexdigest()),
        ("sesca_min_nm", options.min_nm), ("sesca_max_nm", options.max_nm),
        ("sesca_repeats", options.repeats), ("sesca_max_iterations", 5000),
        ("sesca_seed", options.seed), ("sesca_input_units", "kMRE = MRE / 1000"),
        ("sesca_rescaling", "Enabled; scaled_observed_MRE = observed_MRE * scaling_factor"),
        ("sesca_fitted_spectra", sum(fit.status == "ok" for fit in fits)),
        ("sesca_skipped_spectra", sum(fit.status == "skipped" for fit in fits)),
    ])


def write_results(workbook, analysis: SescaAnalysis) -> None:
    """Write fractions and spectra with explicit units and scaling semantics."""
    results = workbook.create_sheet(RESULTS_SHEET)
    results.append([
        "spectrum", "status", "note", *[f"{name}_fraction" for name in analysis.classes],
        "scaling_factor", "RMSD_kMRE", "NRMSD", "fitted_points", "min_nm", "max_nm",
    ])
    curves = workbook.create_sheet(FIT_SHEET)
    curves.append([
        "spectrum", "wavelength_nm", "observed_MRE", "scaled_observed_MRE",
        "fitted_MRE", "residual_MRE",
    ])
    for fit in analysis.fits:
        results.append([
            fit.name, fit.status, fit.note, *(fit.fractions or [None] * len(analysis.classes)),
            fit.scale, fit.rmsd_kmre, fit.nrmsd, len(fit.wavelengths),
            min(fit.wavelengths) if fit.wavelengths else None,
            max(fit.wavelengths) if fit.wavelengths else None,
        ])
        for column in range(4, 4 + len(analysis.classes)):
            results.cell(results.max_row, column).number_format = "0.00%"
        for wave, observed, predicted in zip(fit.wavelengths, fit.observed_mre, fit.fitted_mre):
            scaled = observed * fit.scale
            curves.append([fit.name, wave, observed, scaled, predicted, scaled - predicted])
    for sheet in (results, curves):
        sheet.freeze_panes = "B2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.column_dimensions["A"].width = 65
        sheet.row_dimensions[1].height = 32
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for column in range(2, sheet.max_column + 1):
            sheet.column_dimensions[get_column_letter(column)].width = 22
    results.column_dimensions["B"].width = 12
    results.column_dimensions["C"].width = 75
    for row in results.iter_rows(min_row=2, min_col=4 + len(analysis.classes), max_col=6 + len(analysis.classes)):
        for cell in row:
            cell.number_format = "0.000000"
    for row in curves.iter_rows(min_row=2, min_col=3):
        for cell in row:
            cell.number_format = "0.000000"
