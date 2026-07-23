import pandas as pd
from io import StringIO
from openpyxl.chart import Reference, ScatterChart, Series
from pathlib import Path

folder = Path(
    "/Users/andriikurochka/Documents/PGA_methals/K.Andrii/260720_PLGA-preliminary"
)

files = sorted(folder.glob("*.csv"))
water_files = [file for file in files if file.stem.lower() == "water"]

if not files:
    raise FileNotFoundError(f"No CSV files found directly inside {folder}")
if len(water_files) != 1:
    raise FileNotFoundError(
        f"Expected exactly one water.csv in {folder}, found {len(water_files)}"
    )

water_file = water_files[0]
files = [water_file, *(file for file in files if file != water_file)]

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

# Match the example layout: raw CD data, zero, then water-corrected CD data.
summary_columns = [
    result["wavelength_nm"].rename("wavelength_nm"),
    cd_data[water_name].rename(water_name),
    *(cd_data[name].rename(name) for name in sample_names),
    pd.Series(0, index=result.index, name="zero"),
    *(
        (cd_data[name] - cd_data[water_name]).rename(name)
        for name in sample_names
    ),
]
summary = pd.concat(summary_columns, axis=1)

output_file = folder.parent / f"{folder.name}-combined.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    result.to_excel(writer, sheet_name="RAW", index=False)
    summary.to_excel(writer, sheet_name="SUMMARY", index=False)

    summary_sheet = writer.sheets["SUMMARY"]
    chart = ScatterChart()
    chart.title = "Water-corrected CD spectra"
    chart.style = 13
    chart.scatterStyle = "smoothMarker"
    chart.x_axis.title = "Wavelength (nm)"
    chart.y_axis.title = "CD (mdeg)"
    chart.x_axis.scaling.min = float(result["wavelength_nm"].min())
    chart.x_axis.scaling.max = float(result["wavelength_nm"].max())
    chart.legend.position = "b"
    chart.height = 7.5
    chart.width = 15

    x_values = Reference(
        summary_sheet,
        min_col=1,
        min_row=2,
        max_row=len(summary) + 1,
    )
    zero_column = len(cd_data) + 2
    for column in range(zero_column, summary.shape[1] + 1):
        y_values = Reference(
            summary_sheet,
            min_col=column,
            min_row=2,
            max_row=len(summary) + 1,
        )
        series = Series(
            y_values,
            x_values,
            title=summary_sheet.cell(row=1, column=column).value,
        )
        series.smooth = True
        series.marker.symbol = "none"
        chart.series.append(series)

    summary_sheet.add_chart(chart, "A10")

print(f"Saved {output_file}")
