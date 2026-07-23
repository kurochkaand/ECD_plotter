import pandas as pd
from io import StringIO
from pathlib import Path

folder = Path(
    "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260720_PLGA-preliminary"
)

result = None

for file in sorted(folder.glob("*.csv")):
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

if result is None:
    raise FileNotFoundError(f"No CSV files found directly inside {folder}")

output_file = folder.parent / f"{folder.name}-combined.xlsx"
result.to_excel(output_file, index=False)
print(f"Saved {output_file}")
