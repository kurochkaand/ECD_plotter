"""
Spectral component analysis for an Excel workbook on Windows.

Expected workbook layout
------------------------
Sheet: "Sheet1"
Row 1: headers
Column A: wavelength values (descending or ascending)
Columns B...: one spectrum per column

The program:
1. reads and validates the workbook;
2. restricts the data to a configurable wavelength interval;
3. estimates the number of independent spectral components with SVD;
4. performs PCA and plots scores, loadings, explained variance, and residuals;
5. optionally performs non-negative MCR-ALS to estimate constituent-like
   spectra and their contributions;
6. saves figures and CSV result tables.

Install dependencies:
    python -m pip install numpy pandas matplotlib scipy openpyxl

Run using the path configured below:
    python spectral_component_analysis.py

Optional command-line overrides:
    python spectral_component_analysis.py --file data.xlsx --components 3 --no-show
"""

from __future__ import annotations

import argparse
import sys
import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.optimize import nnls


# =============================================================================
# USER CONFIGURATION
# =============================================================================

# Put the path to the Excel file here. An absolute path is recommended.
EXCEL_PATH = Path("corrected.xlsx")

SHEET_NAME = "Sheet1"

# Wavelength interval used for all calculations and plots, inclusive.
MIN_WAVELENGTH = 190.0
MAX_WAVELENGTH = 260.0

# Automatic selection method used when N_COMPONENTS is None.
# "gavish_donoho" estimates the numerical rank from the singular-value noise
# threshold. "variance" uses CUMULATIVE_VARIANCE_TARGET.
COMPONENT_SELECTION: Literal["gavish_donoho", "variance"] = "gavish_donoho"
N_COMPONENTS: int | None = None
MAX_COMPONENTS = 4
CUMULATIVE_VARIANCE_TARGET = 0.995

# Rank estimation is normally performed on the uncentered spectral matrix,
# because its matrix rank is closer to the number of independent absorbing
# profiles. Set True only when overall offsets dominate the result.
MEAN_CENTER_FOR_RANK = False

# PCA is normally mean-centered across spectra.
MEAN_CENTER_FOR_PCA = True

# Missing values are interpolated along wavelength only when the missing
# fraction in a spectrum does not exceed this limit.
MAX_MISSING_FRACTION = 0.05

# MCR-ALS can estimate non-negative constituent-like spectra and their relative
# contributions. It is skipped if the input contains materially negative data.
RUN_MCR_ALS = True
MCR_MAX_ITERATIONS = 500
MCR_TOLERANCE = 1e-8
MCR_RESTARTS = 5
MCR_NEGATIVE_TOLERANCE = 1e-12

OUTPUT_DIRECTORY = Path("component_analysis_output")
SAVE_FIGURES = True
SHOW_PLOTS = True
FIGURE_DPI = 180


@dataclass(frozen=True)
class SpectralDataset:
    wavelengths: np.ndarray
    spectra: np.ndarray  # shape: (number of spectra, number of wavelengths)
    names: list[str]


@dataclass(frozen=True)
class SVDResult:
    singular_values: np.ndarray
    explained_energy_ratio: np.ndarray
    cumulative_energy_ratio: np.ndarray
    gavish_donoho_threshold: float
    gavish_donoho_rank: int
    variance_rank: int


@dataclass(frozen=True)
class PCAResult:
    mean_spectrum: np.ndarray
    scores: np.ndarray
    loadings: np.ndarray
    singular_values: np.ndarray
    explained_variance_ratio: np.ndarray
    cumulative_variance_ratio: np.ndarray
    reconstruction: np.ndarray
    residuals: np.ndarray


@dataclass(frozen=True)
class MCRResult:
    contributions: np.ndarray
    component_spectra: np.ndarray
    reconstruction: np.ndarray
    residuals: np.ndarray
    relative_error: float
    iterations: int


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Perform SVD/PCA and optional MCR-ALS on spectral columns in Excel."
    )
    parser.add_argument(
        "--file",
        type=Path,
        default=EXCEL_PATH,
        help="Optional override for EXCEL_PATH configured in the script.",
    )
    parser.add_argument(
        "--min-wavelength",
        type=float,
        default=MIN_WAVELENGTH,
        help="Optional override for MIN_WAVELENGTH.",
    )
    parser.add_argument(
        "--max-wavelength",
        type=float,
        default=MAX_WAVELENGTH,
        help="Optional override for MAX_WAVELENGTH.",
    )
    parser.add_argument(
        "--components",
        type=int,
        default=N_COMPONENTS,
        help="Force the number of components instead of automatic selection.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_DIRECTORY,
        help="Optional output-directory override.",
    )
    parser.add_argument(
        "--no-mcr",
        action="store_true",
        help="Skip MCR-ALS even when the data are non-negative.",
    )
    parser.add_argument(
        "--no-show",
        action="store_true",
        help="Save plots without opening interactive windows.",
    )
    return parser.parse_args()


def load_spectral_dataset(
    excel_path: Path,
    sheet_name: str,
    min_wavelength: float,
    max_wavelength: float,
) -> SpectralDataset:
    if not excel_path.exists():
        raise FileNotFoundError(
            f"Excel file not found: {excel_path}\n"
            "Edit EXCEL_PATH near the top of the script or pass --file."
        )

    try:
        frame = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
    except ValueError as exc:
        raise ValueError(
            f"Could not read sheet {sheet_name!r} from {excel_path}. "
            "Check that the sheet exists and is spelled exactly as configured."
        ) from exc

    if frame.shape[1] < 3:
        raise ValueError(
            "The sheet must contain wavelength column A and at least two spectral columns."
        )

    wavelength_name = frame.columns[0]
    frame[wavelength_name] = pd.to_numeric(frame[wavelength_name], errors="coerce")
    frame = frame.dropna(subset=[wavelength_name]).copy()

    if frame.empty:
        raise ValueError("No numeric wavelength values were found in column A.")

    spectral_columns = list(frame.columns[1:])
    for column in spectral_columns:
        frame[column] = pd.to_numeric(frame[column], errors="coerce")

    # Duplicate wavelengths are averaged rather than silently counted twice.
    if frame[wavelength_name].duplicated().any():
        warnings.warn(
            "Duplicate wavelength rows were found. Their spectral values were averaged.",
            RuntimeWarning,
        )
        frame = frame.groupby(wavelength_name, as_index=False)[spectral_columns].mean()

    low, high = sorted((float(min_wavelength), float(max_wavelength)))
    observed_min = float(frame[wavelength_name].min())
    observed_max = float(frame[wavelength_name].max())
    frame = frame.loc[
        frame[wavelength_name].between(low, high, inclusive="both")
    ].copy()

    if frame.empty:
        raise ValueError(
            f"No wavelength rows fall inside {low:g}-{high:g} nm. "
            f"Observed workbook range is approximately {observed_min:g}-{observed_max:g} nm."
        )

    # Analysis and plots use ascending wavelength even when Excel is descending.
    frame = frame.sort_values(wavelength_name, ascending=True).reset_index(drop=True)

    invalid_columns: list[str] = []
    for column in spectral_columns:
        missing_fraction = float(frame[column].isna().mean())
        if missing_fraction == 1.0 or missing_fraction > MAX_MISSING_FRACTION:
            invalid_columns.append(f"{column!s} ({missing_fraction:.1%} missing)")
            continue
        if missing_fraction > 0:
            frame[column] = frame[column].interpolate(
                method="linear", limit_direction="both"
            )

    if invalid_columns:
        raise ValueError(
            "Some spectral columns contain too many missing/non-numeric values: "
            + ", ".join(invalid_columns)
        )

    wavelengths = frame[wavelength_name].to_numpy(dtype=float)
    spectra = frame[spectral_columns].to_numpy(dtype=float).T
    names = [str(name) for name in spectral_columns]

    if wavelengths.size < 5:
        raise ValueError("At least five wavelength points are required after filtering.")
    if not np.isfinite(spectra).all():
        raise ValueError("Non-finite spectral values remain after interpolation.")
    if np.allclose(np.std(spectra, axis=1), 0.0):
        raise ValueError("All spectra are constant across the selected wavelength interval.")

    return SpectralDataset(wavelengths=wavelengths, spectra=spectra, names=names)


def gavish_donoho_rank(singular_values: np.ndarray, matrix_shape: tuple[int, int]) -> tuple[int, float]:
    """Estimate matrix rank using the Gavish-Donoho unknown-noise threshold.

    This uses the commonly used polynomial approximation to the optimal
    singular-value threshold coefficient for rectangular matrices.
    """
    m, n = matrix_shape
    beta = min(m, n) / max(m, n)
    omega = 0.56 * beta**3 - 0.95 * beta**2 + 1.82 * beta + 1.43
    threshold = float(omega * np.median(singular_values))
    rank = int(np.count_nonzero(singular_values > threshold))
    return max(rank, 1), threshold


def analyze_singular_values(
    spectra: np.ndarray, *, mean_center: bool | None = None,
) -> SVDResult:
    """Analyze rank, optionally overriding this script's centering setting."""
    matrix = spectra.copy()
    center = MEAN_CENTER_FOR_RANK if mean_center is None else mean_center
    if center:
        matrix -= matrix.mean(axis=0, keepdims=True)

    singular_values = np.linalg.svd(matrix, full_matrices=False, compute_uv=False)
    squared = singular_values**2
    total = float(squared.sum())
    if total <= np.finfo(float).eps:
        raise ValueError("The selected spectral matrix has effectively zero numerical rank.")

    return summarize_singular_values(singular_values, matrix.shape)


def summarize_singular_values(
    singular_values: np.ndarray, matrix_shape: tuple[int, int], *, allow_zero_rank: bool = False,
) -> SVDResult:
    """Summarize an existing decomposition, including zero centered variation."""
    if not singular_values.size or singular_values[0] == 0:
        if not allow_zero_rank:
            raise ValueError("The selected spectral matrix has effectively zero numerical rank.")
        return SVDResult(
            singular_values=singular_values, explained_energy_ratio=np.zeros_like(singular_values),
            cumulative_energy_ratio=np.zeros_like(singular_values), gavish_donoho_threshold=0.0,
            gavish_donoho_rank=0, variance_rank=0,
        )

    # A common scale cancels in the fractions and avoids squaring very large values.
    squared = (singular_values / singular_values[0])**2
    explained = squared / squared.sum()
    cumulative = np.cumsum(explained)
    gd_rank, threshold = gavish_donoho_rank(singular_values, matrix_shape)
    if allow_zero_rank:
        gd_rank = int(np.count_nonzero(singular_values > threshold))
    variance_rank = int(np.searchsorted(cumulative, CUMULATIVE_VARIANCE_TARGET) + 1)

    return SVDResult(
        singular_values=singular_values,
        explained_energy_ratio=explained,
        cumulative_energy_ratio=cumulative,
        gavish_donoho_threshold=threshold,
        gavish_donoho_rank=gd_rank,
        variance_rank=variance_rank,
    )


def choose_component_count(
    svd_result: SVDResult,
    spectra_shape: tuple[int, int],
    forced_components: int | None,
) -> int:
    maximum = min(MAX_COMPONENTS, min(spectra_shape))
    if maximum < 1:
        raise ValueError("Not enough data to calculate a component model.")

    if forced_components is not None:
        if forced_components < 1 or forced_components > maximum:
            raise ValueError(
                f"Requested {forced_components} components, but the allowed range is 1-{maximum}."
            )
        return forced_components

    if COMPONENT_SELECTION == "gavish_donoho":
        selected = svd_result.gavish_donoho_rank
    elif COMPONENT_SELECTION == "variance":
        selected = svd_result.variance_rank
    else:
        raise ValueError(f"Unknown COMPONENT_SELECTION: {COMPONENT_SELECTION}")

    return max(1, min(selected, maximum))


def perform_pca(spectra: np.ndarray, n_components: int) -> PCAResult:
    if MEAN_CENTER_FOR_PCA:
        mean_spectrum = spectra.mean(axis=0)
        matrix = spectra - mean_spectrum
    else:
        mean_spectrum = np.zeros(spectra.shape[1], dtype=float)
        matrix = spectra.copy()

    u, singular_values, vt = np.linalg.svd(matrix, full_matrices=False)
    scores_all = u * singular_values

    squared = singular_values**2
    total = float(squared.sum())
    explained = squared / total if total > np.finfo(float).eps else np.zeros_like(squared)

    scores = scores_all[:, :n_components]
    loadings = vt[:n_components, :]
    reconstruction = scores @ loadings + mean_spectrum
    residuals = spectra - reconstruction

    return PCAResult(
        mean_spectrum=mean_spectrum,
        scores=scores,
        loadings=loadings,
        singular_values=singular_values,
        explained_variance_ratio=explained,
        cumulative_variance_ratio=np.cumsum(explained),
        reconstruction=reconstruction,
        residuals=residuals,
    )


def nndsvd_initialization(matrix: np.ndarray, n_components: int) -> tuple[np.ndarray, np.ndarray]:
    """Non-negative double-SVD initialization for MCR/NMF-like ALS."""
    u, singular_values, vt = np.linalg.svd(matrix, full_matrices=False)
    n_samples, n_features = matrix.shape
    c = np.zeros((n_samples, n_components), dtype=float)
    s = np.zeros((n_components, n_features), dtype=float)

    c[:, 0] = np.sqrt(singular_values[0]) * np.abs(u[:, 0])
    s[0, :] = np.sqrt(singular_values[0]) * np.abs(vt[0, :])

    for index in range(1, n_components):
        x = u[:, index]
        y = vt[index, :]

        x_positive = np.maximum(x, 0.0)
        x_negative = np.maximum(-x, 0.0)
        y_positive = np.maximum(y, 0.0)
        y_negative = np.maximum(-y, 0.0)

        positive_norm = np.linalg.norm(x_positive) * np.linalg.norm(y_positive)
        negative_norm = np.linalg.norm(x_negative) * np.linalg.norm(y_negative)

        if positive_norm >= negative_norm:
            x_part, y_part, sigma = x_positive, y_positive, positive_norm
        else:
            x_part, y_part, sigma = x_negative, y_negative, negative_norm

        if sigma <= np.finfo(float).eps:
            continue

        x_part /= max(np.linalg.norm(x_part), np.finfo(float).eps)
        y_part /= max(np.linalg.norm(y_part), np.finfo(float).eps)
        scale = np.sqrt(singular_values[index] * sigma)
        c[:, index] = scale * x_part
        s[index, :] = scale * y_part

    fill_value = max(float(matrix.mean()) * 1e-6, np.finfo(float).eps)
    c[c == 0.0] = fill_value
    s[s == 0.0] = fill_value
    return c, s


def normalize_mcr_factors(
    contributions: np.ndarray, component_spectra: np.ndarray
) -> tuple[np.ndarray, np.ndarray]:
    """Normalize each component spectrum to unit maximum and preserve C @ S."""
    scale = np.max(component_spectra, axis=1)
    scale[scale <= np.finfo(float).eps] = 1.0
    component_spectra = component_spectra / scale[:, None]
    contributions = contributions * scale[None, :]
    return contributions, component_spectra


def perform_mcr_als(
    spectra: np.ndarray,
    n_components: int,
    max_iterations: int,
    tolerance: float,
    restarts: int,
) -> MCRResult:
    """Fit X ~= C @ S under non-negativity constraints using alternating NNLS."""
    if np.min(spectra) < -MCR_NEGATIVE_TOLERANCE:
        raise ValueError(
            "MCR-ALS requires non-negative input, but negative values were detected."
        )

    matrix = np.maximum(spectra, 0.0)
    matrix_norm = max(np.linalg.norm(matrix), np.finfo(float).eps)
    base_c, base_s = nndsvd_initialization(matrix, n_components)
    rng = np.random.default_rng(20260728)

    best: MCRResult | None = None

    for restart in range(max(1, restarts)):
        if restart == 0:
            c = base_c.copy()
            s = base_s.copy()
        else:
            c = base_c * rng.lognormal(mean=0.0, sigma=0.15, size=base_c.shape)
            s = base_s * rng.lognormal(mean=0.0, sigma=0.15, size=base_s.shape)

        previous_error = np.inf
        iterations_used = 0

        for iteration in range(1, max_iterations + 1):
            # Solve contribution profile for each measured spectrum.
            for row in range(matrix.shape[0]):
                c[row, :], _ = nnls(s.T, matrix[row, :])

            # Solve each wavelength point of the component spectra.
            for column in range(matrix.shape[1]):
                s[:, column], _ = nnls(c, matrix[:, column])

            c, s = normalize_mcr_factors(c, s)
            reconstruction = c @ s
            relative_error = float(np.linalg.norm(matrix - reconstruction) / matrix_norm)
            iterations_used = iteration

            if np.isfinite(previous_error) and (
                abs(previous_error - relative_error)
                <= tolerance * max(1.0, previous_error)
            ):
                break
            previous_error = relative_error

        reconstruction = c @ s
        residuals = matrix - reconstruction
        result = MCRResult(
            contributions=c.copy(),
            component_spectra=s.copy(),
            reconstruction=reconstruction,
            residuals=residuals,
            relative_error=float(np.linalg.norm(residuals) / matrix_norm),
            iterations=iterations_used,
        )

        if best is None or result.relative_error < best.relative_error:
            best = result

    if best is None:
        raise RuntimeError("MCR-ALS did not produce a solution.")
    return best


def save_figure(
    figure: plt.Figure,
    filename: str,
    output_directory: Path,
    save_figures: bool,
) -> None:
    figure.tight_layout()
    if save_figures:
        figure.savefig(output_directory / filename, dpi=FIGURE_DPI, bbox_inches="tight")


def plot_input_spectra(dataset: SpectralDataset, output_directory: Path) -> None:
    fig, ax = plt.subplots(figsize=(10, 6))
    for spectrum, name in zip(dataset.spectra, dataset.names):
        ax.plot(dataset.wavelengths, spectrum, linewidth=1.3, label=name)
    ax.set_title("Input spectra in selected wavelength range")
    ax.set_xlabel("Wavelength (nm)")
    ax.set_ylabel("Signal")
    ax.grid(alpha=0.25)
    if len(dataset.names) <= 15:
        ax.legend(fontsize=8, ncol=2)
    save_figure(fig, "01_input_spectra.png", output_directory, SAVE_FIGURES)


def plot_svd(svd_result: SVDResult, selected_components: int, output_directory: Path) -> None:
    count = min(len(svd_result.singular_values), MAX_COMPONENTS)
    component_axis = np.arange(1, count + 1)

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.semilogy(component_axis, svd_result.singular_values[:count], "o-")
    ax.axhline(
        svd_result.gavish_donoho_threshold,
        linestyle="--",
        label="Gavish-Donoho threshold",
    )
    ax.axvline(selected_components, linestyle=":", label=f"Selected: {selected_components}")
    ax.set_title("Singular-value spectrum")
    ax.set_xlabel("Component number")
    ax.set_ylabel("Singular value (log scale)")
    ax.set_xticks(component_axis)
    ax.grid(alpha=0.25)
    ax.legend()
    save_figure(fig, "02_singular_values.png", output_directory, SAVE_FIGURES)

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.bar(component_axis, 100.0 * svd_result.explained_energy_ratio[:count], label="Individual")
    ax.plot(
        component_axis,
        100.0 * svd_result.cumulative_energy_ratio[:count],
        "o-",
        label="Cumulative",
    )
    ax.axhline(100.0 * CUMULATIVE_VARIANCE_TARGET, linestyle="--")
    ax.axvline(selected_components, linestyle=":")
    ax.set_title("SVD explained spectral energy")
    ax.set_xlabel("Component number")
    ax.set_ylabel("Explained energy (%)")
    ax.set_xticks(component_axis)
    ax.set_ylim(0, 102)
    ax.grid(alpha=0.25)
    ax.legend()
    save_figure(fig, "03_explained_energy.png", output_directory, SAVE_FIGURES)


def plot_pca(
    dataset: SpectralDataset,
    pca: PCAResult,
    output_directory: Path,
) -> None:
    n_components = pca.loadings.shape[0]

    fig, ax = plt.subplots(figsize=(10, 6))
    for index in range(n_components):
        ax.plot(
            dataset.wavelengths,
            pca.loadings[index],
            label=f"PC{index + 1} ({100 * pca.explained_variance_ratio[index]:.2f}%)",
        )
    ax.set_title("PCA loading spectra")
    ax.set_xlabel("Wavelength (nm)")
    ax.set_ylabel("Loading")
    ax.axhline(0.0, linewidth=0.8)
    ax.grid(alpha=0.25)
    ax.legend()
    save_figure(fig, "04_pca_loadings.png", output_directory, SAVE_FIGURES)

    if n_components >= 2:
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.scatter(pca.scores[:, 0], pca.scores[:, 1], s=55)
        for index, name in enumerate(dataset.names):
            ax.annotate(name, (pca.scores[index, 0], pca.scores[index, 1]), fontsize=8)
        ax.axhline(0.0, linewidth=0.8)
        ax.axvline(0.0, linewidth=0.8)
        ax.set_title("PCA score plot")
        ax.set_xlabel(f"PC1 score ({100 * pca.explained_variance_ratio[0]:.2f}%)")
        ax.set_ylabel(f"PC2 score ({100 * pca.explained_variance_ratio[1]:.2f}%)")
        ax.grid(alpha=0.25)
        save_figure(fig, "05_pca_scores.png", output_directory, SAVE_FIGURES)
    else:
        fig, ax = plt.subplots(figsize=(9, 5))
        ax.plot(np.arange(1, len(dataset.names) + 1), pca.scores[:, 0], "o-")
        ax.set_xticks(np.arange(1, len(dataset.names) + 1))
        ax.set_xticklabels(dataset.names, rotation=45, ha="right", fontsize=8)
        ax.set_title("PC1 scores")
        ax.set_xlabel("Spectrum")
        ax.set_ylabel("PC1 score")
        ax.grid(alpha=0.25)
        save_figure(fig, "05_pca_scores.png", output_directory, SAVE_FIGURES)

    residual_rms = np.sqrt(np.mean(pca.residuals**2, axis=1))
    fig, ax = plt.subplots(figsize=(9, 5))
    x = np.arange(1, len(dataset.names) + 1)
    ax.bar(x, residual_rms)
    ax.set_xticks(x)
    ax.set_xticklabels(dataset.names, rotation=45, ha="right", fontsize=8)
    ax.set_title("PCA reconstruction residual by spectrum")
    ax.set_xlabel("Spectrum")
    ax.set_ylabel("Residual RMS")
    ax.grid(axis="y", alpha=0.25)
    save_figure(fig, "06_pca_residuals.png", output_directory, SAVE_FIGURES)


def plot_mcr(
    dataset: SpectralDataset,
    mcr: MCRResult,
    output_directory: Path,
) -> None:
    n_components = mcr.component_spectra.shape[0]

    fig, ax = plt.subplots(figsize=(10, 6))
    for index in range(n_components):
        ax.plot(
            dataset.wavelengths,
            mcr.component_spectra[index],
            linewidth=1.8,
            label=f"MCR component {index + 1}",
        )
    ax.set_title("MCR-ALS estimated constituent-like spectra")
    ax.set_xlabel("Wavelength (nm)")
    ax.set_ylabel("Relative component spectrum")
    ax.grid(alpha=0.25)
    ax.legend()
    save_figure(fig, "07_mcr_component_spectra.png", output_directory, SAVE_FIGURES)

    fig, ax = plt.subplots(figsize=(10, 6))
    x = np.arange(1, len(dataset.names) + 1)
    for index in range(n_components):
        ax.plot(
            x,
            mcr.contributions[:, index],
            "o-",
            label=f"Component {index + 1}",
        )
    ax.set_xticks(x)
    ax.set_xticklabels(dataset.names, rotation=45, ha="right", fontsize=8)
    ax.set_title("MCR-ALS component contributions")
    ax.set_xlabel("Measured spectrum")
    ax.set_ylabel("Relative contribution")
    ax.grid(alpha=0.25)
    ax.legend()
    save_figure(fig, "08_mcr_contributions.png", output_directory, SAVE_FIGURES)

    # Overlay measured and reconstructed spectra. Solid = measured, dashed = fit.
    fig, ax = plt.subplots(figsize=(10, 6))
    for measured, reconstructed, name in zip(
        dataset.spectra, mcr.reconstruction, dataset.names
    ):
        line = ax.plot(dataset.wavelengths, measured, linewidth=1.1, label=name)[0]
        ax.plot(
            dataset.wavelengths,
            reconstructed,
            linestyle="--",
            linewidth=1.0,
            color=line.get_color(),
        )
    ax.set_title("MCR-ALS reconstruction: solid measured, dashed fitted")
    ax.set_xlabel("Wavelength (nm)")
    ax.set_ylabel("Signal")
    ax.grid(alpha=0.25)
    if len(dataset.names) <= 15:
        ax.legend(fontsize=8, ncol=2)
    save_figure(fig, "09_mcr_reconstruction.png", output_directory, SAVE_FIGURES)


def export_results(
    dataset: SpectralDataset,
    svd_result: SVDResult,
    pca: PCAResult,
    output_directory: Path,
    mcr: MCRResult | None,
) -> None:
    input_frame = pd.DataFrame(
        dataset.spectra.T,
        columns=dataset.names,
    )
    input_frame.insert(0, "wavelength_nm", dataset.wavelengths)
    input_frame.to_csv(output_directory / "filtered_input_spectra.csv", index=False)

    svd_frame = pd.DataFrame(
        {
            "component": np.arange(1, len(svd_result.singular_values) + 1),
            "singular_value": svd_result.singular_values,
            "explained_energy_fraction": svd_result.explained_energy_ratio,
            "cumulative_energy_fraction": svd_result.cumulative_energy_ratio,
        }
    )
    svd_frame.to_csv(output_directory / "svd_diagnostics.csv", index=False)

    loading_frame = pd.DataFrame(
        pca.loadings.T,
        columns=[f"PC{index + 1}" for index in range(pca.loadings.shape[0])],
    )
    loading_frame.insert(0, "wavelength_nm", dataset.wavelengths)
    loading_frame.to_csv(output_directory / "pca_loadings.csv", index=False)

    scores_frame = pd.DataFrame(
        pca.scores,
        columns=[f"PC{index + 1}" for index in range(pca.scores.shape[1])],
    )
    scores_frame.insert(0, "spectrum", dataset.names)
    scores_frame.to_csv(output_directory / "pca_scores.csv", index=False)

    pca_reconstruction = pd.DataFrame(pca.reconstruction.T, columns=dataset.names)
    pca_reconstruction.insert(0, "wavelength_nm", dataset.wavelengths)
    pca_reconstruction.to_csv(output_directory / "pca_reconstructed_spectra.csv", index=False)

    if mcr is not None:
        component_frame = pd.DataFrame(
            mcr.component_spectra.T,
            columns=[f"component_{index + 1}" for index in range(mcr.component_spectra.shape[0])],
        )
        component_frame.insert(0, "wavelength_nm", dataset.wavelengths)
        component_frame.to_csv(output_directory / "mcr_component_spectra.csv", index=False)

        contribution_frame = pd.DataFrame(
            mcr.contributions,
            columns=[f"component_{index + 1}" for index in range(mcr.contributions.shape[1])],
        )
        contribution_frame.insert(0, "spectrum", dataset.names)
        contribution_frame.to_csv(output_directory / "mcr_contributions.csv", index=False)

        mcr_reconstruction = pd.DataFrame(mcr.reconstruction.T, columns=dataset.names)
        mcr_reconstruction.insert(0, "wavelength_nm", dataset.wavelengths)
        mcr_reconstruction.to_csv(
            output_directory / "mcr_reconstructed_spectra.csv", index=False
        )


def print_summary(
    dataset: SpectralDataset,
    svd_result: SVDResult,
    selected_components: int,
    pca: PCAResult,
    mcr: MCRResult | None,
    output_directory: Path,
) -> None:
    print("\n=== Spectral component analysis summary ===")
    print(f"Spectra:                {dataset.spectra.shape[0]}")
    print(f"Wavelength points:      {dataset.spectra.shape[1]}")
    print(
        f"Analyzed range:         {dataset.wavelengths.min():g}-"
        f"{dataset.wavelengths.max():g} nm"
    )
    print(f"Gavish-Donoho rank:     {svd_result.gavish_donoho_rank}")
    print(
        f"Rank at {100 * CUMULATIVE_VARIANCE_TARGET:.2f}% energy: "
        f"{svd_result.variance_rank}"
    )
    print(f"Selected components:    {selected_components}")
    print(
        f"PCA captured variance:  "
        f"{100 * pca.cumulative_variance_ratio[selected_components - 1]:.4f}%"
    )
    pca_relative_error = np.linalg.norm(pca.residuals) / max(
        np.linalg.norm(dataset.spectra), np.finfo(float).eps
    )
    print(f"PCA relative residual:  {pca_relative_error:.6g}")
    if mcr is not None:
        print(f"MCR-ALS relative error: {mcr.relative_error:.6g}")
        print(f"MCR-ALS iterations:     {mcr.iterations}")
    else:
        print("MCR-ALS:                not run")
    print(f"Results directory:      {output_directory.resolve()}")
    print(
        "\nInterpretation warning: the selected number is an estimate of numerical "
        "spectral rank. Co-linear species, invariant concentration ratios, baseline "
        "drift, scattering, and noise can respectively hide or create apparent "
        "components. MCR-ALS solutions are non-negative but are not chemically unique "
        "without additional constraints or reference spectra."
    )


def main() -> int:
    args = parse_arguments()

    output_directory = args.output
    run_mcr = RUN_MCR_ALS and not args.no_mcr
    show_plots = SHOW_PLOTS and not args.no_show

    try:
        output_directory.mkdir(parents=True, exist_ok=True)
        dataset = load_spectral_dataset(
            excel_path=args.file,
            sheet_name=SHEET_NAME,
            min_wavelength=args.min_wavelength,
            max_wavelength=args.max_wavelength,
        )

        svd_result = analyze_singular_values(dataset.spectra)
        selected_components = choose_component_count(
            svd_result=svd_result,
            spectra_shape=dataset.spectra.shape,
            forced_components=args.components,
        )
        pca = perform_pca(dataset.spectra, selected_components)

        mcr: MCRResult | None = None
        if run_mcr:
            minimum = float(dataset.spectra.min())
            if minimum < -MCR_NEGATIVE_TOLERANCE:
                warnings.warn(
                    f"MCR-ALS was skipped because the data contain negative values "
                    f"(minimum {minimum:.6g}). PCA results are still valid. Do not "
                    "shift ECD or difference spectra merely to force non-negativity.",
                    RuntimeWarning,
                )
            else:
                mcr = perform_mcr_als(
                    spectra=dataset.spectra,
                    n_components=selected_components,
                    max_iterations=MCR_MAX_ITERATIONS,
                    tolerance=MCR_TOLERANCE,
                    restarts=MCR_RESTARTS,
                )

        plot_input_spectra(dataset, output_directory)
        plot_svd(svd_result, selected_components, output_directory)
        plot_pca(dataset, pca, output_directory)
        if mcr is not None:
            plot_mcr(dataset, mcr, output_directory)

        export_results(
            dataset=dataset,
            svd_result=svd_result,
            pca=pca,
            output_directory=output_directory,
            mcr=mcr,
        )
        print_summary(
            dataset=dataset,
            svd_result=svd_result,
            selected_components=selected_components,
            pca=pca,
            mcr=mcr,
            output_directory=output_directory,
        )

        if show_plots:
            plt.show()
        else:
            plt.close("all")
        return 0

    except (FileNotFoundError, ValueError, OSError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # Defensive: preserve a clear terminal error.
        print(f"UNEXPECTED ERROR: {type(exc).__name__}: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
