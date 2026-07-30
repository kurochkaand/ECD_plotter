import pandas as pd
from io import StringIO
from pathlib import Path

folder = Path(
    # "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260720_PLGA-preliminary"
    # "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260721_PLGA-first-tiration/"
    # "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260721_PLGA-metals/"
    # "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260722-PLGA-pH3-methals/"
    # "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260722-PLGA-pH-titrat/"
    # "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260722-PDGA-pH-titrat/"
    # "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260727-PLGA-pH8-methals/"
    "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260729-PLGA-pH8-methals/"
)

def file_creation_time(file: Path) -> float:
    """Return the file creation time, with a fallback for non-macOS systems."""
    file_stat = file.stat()
    return getattr(file_stat, "st_birthtime", file_stat.st_ctime)


files = sorted(
    folder.glob("*.csv"),
    key=lambda file: (file_creation_time(file), file.name.lower()),
)
water_files = [file for file in files if file.stem.lower() == "water"]

if not files:
    raise FileNotFoundError(f"No CSV files found directly inside {folder}")
if len(water_files) != 1:
    raise FileNotFoundError(
        f"Expected exactly one water.csv in {folder}, found {len(water_files)}"
    )

water_file = water_files[0]

result = None
cd_data = {}

for file in files:
    # Read only the block between XYDATA and Extended Information.
    data_start = None
    data_end = None
    lines = file.read_text(encoding="latin-1").splitlines(keepends=True)
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

    df = pd.read_csv(
        StringIO(data_block),
        header=None,
        names=["wavelength_nm", "cd_mdeg", "ht_v"],
        skip_blank_lines=True,
    )

    if result is None:
        result = df[["wavelength_nm"]].copy()
    elif not result["wavelength_nm"].equals(df["wavelength_nm"]):
        raise ValueError(f"X axis in {file} does not match the other files")

    result[f"{file.stem}_cd_mdeg"] = df["cd_mdeg"]
    result[f"{file.stem}_ht_v"] = df["ht_v"]
    cd_data[file.stem] = df["cd_mdeg"]

water_name = water_file.stem
sample_names = [name for name in cd_data if name != water_name]

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
    concentration_table.columns = [
        "spectrum",
        "concentration_coefficient",
        *(["Ph"] if concentration_table.shape[1] == 3 else []),
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
    concentration_table["concentration_coefficient"],
    errors="raise",
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
    raise ValueError(f"{concentration_file} does not match the CSV files ({'; '.join(problems)})")
if (coefficients <= 0).any():
    invalid = coefficients[coefficients <= 0].index.tolist()
    raise ValueError(
        f"Concentration coefficients must be greater than zero for: "
        f"{', '.join(invalid)}"
    )

# Raw CD data is preserved; corrected data is water-subtracted and
# concentration-normalized.
summary_columns = [
    result["wavelength_nm"].rename("wavelength_nm"),
    cd_data[water_name].rename(water_name),
    *(cd_data[name].rename(name) for name in sample_names),
    pd.Series(0, index=result.index, name="zero"),
    *(
        (
            (cd_data[name] - cd_data[water_name])
            / coefficients[name]
        ).rename(name)
        for name in sample_names
    ),
]
summary = pd.concat(summary_columns, axis=1)

output_file = folder.parent / f"{folder.name}-combined.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    result.to_excel(writer, sheet_name="RAW", index=False)
    summary.to_excel(writer, sheet_name="SUMMARY", index=False)


print(f"Saved {output_file}")