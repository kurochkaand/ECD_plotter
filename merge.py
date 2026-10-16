"""Merge Windows instrument CSV exports and normalize CD spectra."""

from __future__ import annotations

import argparse
import sys
from io import StringIO
from math import isfinite
from pathlib import Path

import pandas as pd


DEFAULT_FOLDER = Path(
    r"X:\Cations_and_PGA\ECD-Szeged\K.Andrii\260730-PLGA-titration-Zn"
)
SPECTRAL_COLUMNS = ["wavelength_nm", "cd_mdeg", "ht_v"]
CONCENTRATION_COLUMNS = ["spectrum", "concentration_coefficient", "Ph"]


def find_spectral_files(folder: Path) -> tuple[list[Path], Path]:
    """Find CSV files in Windows creation order and identify the water reference."""
    if not folder.is_dir():
        raise NotADirectoryError(f"Data folder not found: {folder}")

    files = sorted(
        (file for file in folder.glob("*.csv") if file.is_file()),
        key=lambda file: (file.stat().st_birthtime, file.name.lower()),
    )
    if not files:
        raise FileNotFoundError(f"No CSV files found directly inside {folder}")

    water_files = [file for file in files if file.stem.lower().startswith("water")]
    if len(water_files) != 1:
        raise ValueError(
            f"Expected exactly one CSV with a name starting with 'water' in {folder}, "
            f"found {len(water_files)}"
        )
    return files, water_files[0]


def read_spectrum(file: Path) -> pd.DataFrame:
    """Read the block between XYDATA and Extended Information."""
    lines = file.read_text(encoding="latin-1").splitlines(keepends=True)
    data_start = None
    data_end = None
    for row_number, line in enumerate(lines):
        if line.strip() == "XYDATA":
            data_start = row_number + 1
        elif line.startswith("##### Extended Information"):
            data_end = row_number
            break

    if data_start is None:
        raise ValueError(f"XYDATA marker not found in {file}")
    if data_end is None:
        raise ValueError(f"Extended Information marker not found in {file}")

    data_block = "".join(lines[data_start:data_end])
    if not data_block.strip():
        raise ValueError(f"No spectral data found in {file}")

    return pd.read_csv(
        StringIO(data_block),
        header=None,
        names=SPECTRAL_COLUMNS,
        skip_blank_lines=True,
    )


def load_concentration_coefficients(folder: Path, sample_names: list[str]) -> pd.Series:
    """Create or validate the concentration table and return coefficients by name."""
    concentration_file = folder / "concentration_table.txt"
    if not concentration_file.exists():
        concentration_table = pd.DataFrame(
            {
                "spectrum": sample_names,
                "concentration_coefficient": 1.0,
                "Ph": 7.0,
            }
        )
        concentration_table.to_csv(concentration_file, sep="\t", index=False)
        print(f"Created {concentration_file}")
    else:
        concentration_table = pd.read_csv(concentration_file, sep="\t")
        if concentration_table.shape[1] not in (2, 3):
            raise ValueError(
                f"{concentration_file} must have two or three tab-separated columns: "
                "spectrum name, concentration coefficient, and optional Ph"
            )
        concentration_table.columns = CONCENTRATION_COLUMNS[
            : concentration_table.shape[1]
        ]
        concentration_table["spectrum"] = (
            concentration_table["spectrum"].astype(str).str.strip()
        )
        if concentration_table["spectrum"].duplicated().any():
            duplicates = concentration_table.loc[
                concentration_table["spectrum"].duplicated(), "spectrum"
            ].tolist()
            raise ValueError(
                f"Duplicate spectra in {concentration_file}: {', '.join(duplicates)}"
            )

    concentration_table["concentration_coefficient"] = pd.to_numeric(
        concentration_table["concentration_coefficient"], errors="raise"
    )
    coefficients = concentration_table.set_index("spectrum")[
        "concentration_coefficient"
    ]
    missing_spectra = [name for name in sample_names if name not in coefficients.index]
    extra_spectra = [name for name in coefficients.index if name not in sample_names]
    if missing_spectra or extra_spectra:
        problems = []
        if missing_spectra:
            problems.append(f"missing: {', '.join(missing_spectra)}")
        if extra_spectra:
            problems.append(f"not present in data folder: {', '.join(extra_spectra)}")
        raise ValueError(
            f"{concentration_file} does not match the CSV files "
            f"({'; '.join(problems)})"
        )

    invalid = coefficients[
        ~coefficients.map(isfinite) | (coefficients <= 0)
    ].index.tolist()
    if invalid:
        raise ValueError(
            "Concentration coefficients must be finite and greater than zero for: "
            + ", ".join(invalid)
        )
    return coefficients


def merge_spectra(folder: Path) -> Path:
    """Write RAW and SUMMARY sheets, preserving the instrument's wavelength axis."""
    files, water_file = find_spectral_files(folder)
    raw_columns: list[pd.Series] = []
    cd_data: dict[str, pd.Series] = {}

    for file in files:
        frame = read_spectrum(file)
        if not raw_columns:
            raw_columns.append(frame["wavelength_nm"])
        elif not raw_columns[0].equals(frame["wavelength_nm"]):
            raise ValueError(f"X axis in {file} does not match the other files")

        raw_columns.extend(
            [
                frame["cd_mdeg"].rename(f"{file.stem}_cd_mdeg"),
                frame["ht_v"].rename(f"{file.stem}_ht_v"),
            ]
        )
        cd_data[file.stem] = frame["cd_mdeg"]

    raw = pd.concat(raw_columns, axis=1)
    water_name = water_file.stem
    water = cd_data[water_name]
    sample_names = [name for name in cd_data if name != water_name]
    coefficients = load_concentration_coefficients(folder, sample_names)

    # Keep raw CD values before the zero separator and corrected values after it.
    summary = pd.concat(
        [
            raw["wavelength_nm"],
            water.rename(water_name),
            *(cd_data[name].rename(name) for name in sample_names),
            pd.Series(0, index=raw.index, name="zero"),
            *(
                ((cd_data[name] - water) / coefficients[name]).rename(name)
                for name in sample_names
            ),
        ],
        axis=1,
    )

    output_file = folder.parent / f"{folder.name}-combined.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        raw.to_excel(writer, sheet_name="RAW", index=False)
        summary.to_excel(writer, sheet_name="SUMMARY", index=False)
    return output_file


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Merge instrument CSV files and normalize CD spectra on Windows."
    )
    parser.add_argument(
        "--folder",
        type=Path,
        default=DEFAULT_FOLDER,
        help="Data-folder override for DEFAULT_FOLDER configured in the script.",
    )
    args = parser.parse_args()

    try:
        output_file = merge_spectra(args.folder.resolve())
    except (OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Saved {output_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
