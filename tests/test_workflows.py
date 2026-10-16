"""Regression checks using temporary instrument exports and Excel workbooks."""

import os
import subprocess
import sys
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import numpy as np
import pandas as pd

import merge
import spectral_component_analysis as analysis


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class MergeTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="ECD cleanup ")
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name) / "instrument data"
        self.folder.mkdir()
        self.write_spectrum("sample", [5, 8, 11])

    def write_spectrum(
        self, name, cd, wavelengths=(200, 201, 202), absorbance=0.25,
        date="26/10/14", time="18:41:36", ht=500,
    ):
        rows = []
        for wavelength, signal in zip(wavelengths, cd):
            values = [wavelength, signal, ht]
            if absorbance is not None:
                values.append(absorbance)
            rows.append(",".join(map(str, values)) + "\n")
        (self.folder / f"{name}.csv").write_text(
            "TITLE,Instrument export\nDATA TYPE,\nORIGIN,JASCO\nOWNER,\n"
            f"DATE,{date}\nTIME,{time}\nXYDATA\n" + "".join(rows)
            + "##### Extended Information\n",
            encoding="latin-1",
        )

    def write_coefficients(self, rows):
        (self.folder / "concentration_table.txt").write_text(
            "spectrum\tconcentration_coefficient\n" + rows,
            encoding="utf-8",
        )

    def test_workbook_preserves_raw_values_and_normalizes_without_subtraction(self):
        self.write_coefficients("sample\t2\n")
        output = merge.merge_spectra(self.folder)
        self.assertEqual(output, self.folder.parent / "instrument data-combined.xlsx")
        with pd.ExcelFile(output) as workbook:
            self.assertEqual(workbook.sheet_names, ["CD", "HT", "Abs", "SUMMARY"])
            cd = pd.read_excel(workbook, sheet_name="CD")
            ht = pd.read_excel(workbook, sheet_name="HT")
            absorbance = pd.read_excel(workbook, sheet_name="Abs")
            summary = pd.read_excel(workbook, sheet_name="SUMMARY", header=None)
        for channel in (cd, ht, absorbance):
            self.assertEqual(channel.columns.tolist(), ["wavelength_nm", "sample"])
            self.assertEqual(channel["wavelength_nm"].tolist(), [200, 201, 202])
        self.assertEqual(cd["sample"].tolist(), [5, 8, 11])
        self.assertEqual(ht["sample"].tolist(), [500, 500, 500])
        self.assertEqual(absorbance["sample"].tolist(), [0.25, 0.25, 0.25])
        self.assertEqual(
            summary.iloc[0].tolist(),
            ["wavelength_nm", "sample", "zero", "sample"],
        )
        np.testing.assert_allclose(
            summary.iloc[1:].to_numpy(dtype=float),
            [[200, 5, 0, 2.5], [201, 8, 0, 4], [202, 11, 0, 5.5]],
        )

    def test_first_run_creates_default_concentration_table(self):
        merge.merge_spectra(self.folder)
        table = pd.read_csv(self.folder / "concentration_table.txt", sep="\t")
        self.assertEqual(table.to_dict("list"), {
            "spectrum": ["sample"], "concentration_coefficient": [1.0], "Ph": [7.0],
        })

    def test_ph_is_extracted_from_filename_tokens(self):
        for name, expected in (
            ("Cd-peptide_0.2ekv_pH7.43_1mm_185-300nm_4sc_20nm-1", 7.43),
            ("Cd-peptide_1.6ekv_pH5.77_1mm_185-300nm_4sc_20nm-1", 5.77),
            ("peptide_pH2.18_201uL_1mm_185-300nm_4sc_20nm-1", 2.18),
            ("peptide_pH7.9_0uL_1mm_185-300nm_4sc_20nm-1", 7.9),
            ("peptide_pH 7.01_21uL_1mm_185-300nm_4sc_20nm-1", 7.01),
            ("peptide_PH8_1mm", 8.0),
            ("pH0", 0.0),
            ("peptide_pH-0.5_1mm", -0.5),
            ("H2O_1mm_185-300nm_4sc_20nm-1", None),
            ("phosphate_7.43_1mm", None),
        ):
            with self.subTest(name=name):
                self.assertEqual(merge.extract_ph_from_name(name), expected)

    def test_new_table_uses_filename_ph_and_default_for_missing_tokens(self):
        names = [
            "Cd-peptide_0.2ekv_pH7.43_1mm_185-300nm_4sc_20nm-1",
            "peptide_pH7.9_0uL_1mm_185-300nm_4sc_20nm-1",
            "H2O_1mm_185-300nm_4sc_20nm-1",
        ]
        merge.load_concentration_coefficients(self.folder, names)
        table = pd.read_csv(self.folder / "concentration_table.txt", sep="\t")
        self.assertEqual(table["spectrum"].tolist(), names)
        self.assertEqual(table["Ph"].tolist(), [7.43, 7.9, 7.0])
        self.assertEqual(table["concentration_coefficient"].tolist(), [1.0, 1.0, 1.0])

    def test_missing_ph_in_existing_tables_is_inferred(self):
        for contents in (
            "spectrum\tconcentration_coefficient\nsample_pH5.77\t2\n",
            "spectrum\tconcentration_coefficient\tPh\nsample_pH5.77\t2\t\n",
        ):
            with self.subTest(contents=contents):
                table_file = self.folder / "concentration_table.txt"
                table_file.write_text(contents, encoding="utf-8")
                merge.load_concentration_coefficients(self.folder, ["sample_pH5.77"])
                table = pd.read_csv(table_file, sep="\t")
                self.assertEqual(table["Ph"].tolist(), [5.77])
                self.assertEqual(table["concentration_coefficient"].tolist(), [2])

    def test_existing_ph_is_preserved_unless_refresh_is_requested(self):
        table_file = self.folder / "concentration_table.txt"
        table_file.write_text(
            "spectrum\tconcentration_coefficient\tPh\n"
            "sample_pH7.43\t2\t6.8\nH2O_1mm\t3\t6.5\n",
            encoding="utf-8",
        )
        names = ["sample_pH7.43", "H2O_1mm"]
        merge.load_concentration_coefficients(self.folder, names)
        table = pd.read_csv(table_file, sep="\t")
        self.assertEqual(table["Ph"].tolist(), [6.8, 6.5])

        merge.load_concentration_coefficients(self.folder, names, refresh_ph=True)
        table = pd.read_csv(table_file, sep="\t")
        self.assertEqual(table["Ph"].tolist(), [7.43, 6.5])
        self.assertEqual(table["concentration_coefficient"].tolist(), [2, 3])

    def test_cli_refreshes_ph_and_preserves_concentration_coefficients(self):
        self.write_spectrum("peptide_pH5.77", [4, 8, 12], time="19:00:00")
        table_file = self.folder / "concentration_table.txt"
        table_file.write_text(
            "spectrum\tconcentration_coefficient\tPh\n"
            "sample\t2\t6.5\npeptide_pH5.77\t4\t7.0\n",
            encoding="utf-8",
        )
        result = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "merge.py"),
             "--folder", str(self.folder), "--refresh-ph"],
            capture_output=True, text=True, check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        table = pd.read_csv(table_file, sep="\t")
        self.assertEqual(table["spectrum"].tolist(), ["sample", "peptide_pH5.77"])
        self.assertEqual(table["Ph"].tolist(), [6.5, 5.77])
        self.assertEqual(table["concentration_coefficient"].tolist(), [2, 4])

    def test_workbook_and_existing_table_follow_acquisition_datetime(self):
        self.write_spectrum("a_later", [4, 8, 12], date="26/10/15", time="00:00:00")
        self.write_spectrum("z_earlier", [5, 10, 15], date="26/10/13", time="11:59:59 PM")
        self.write_spectrum("morning", [3, 6, 9], time="6:41:36 AM")
        table_file = self.folder / "concentration_table.txt"
        table_file.write_text(
            "spectrum\t\t\t\tconcentration_coefficient\tPh\n"
            "a_later\t4\t6.5\nsample\t2\t7.2\n"
            "z_earlier\t5\t8.1\nmorning\t3\t6.8\n",
            encoding="utf-8",
        )
        expected_names = ["z_earlier", "morning", "sample", "a_later"]

        output = merge.merge_spectra(self.folder)
        with pd.ExcelFile(output) as workbook:
            for sheet_name in ("CD", "HT", "Abs"):
                channel = pd.read_excel(workbook, sheet_name=sheet_name)
                self.assertEqual(channel.columns.tolist(), ["wavelength_nm", *expected_names])
        summary = pd.read_excel(output, sheet_name="SUMMARY", header=None)
        table = pd.read_csv(table_file, sep="\t")
        self.assertEqual(
            summary.iloc[0].tolist(),
            ["wavelength_nm", *expected_names, "zero", *expected_names],
        )
        self.assertEqual(table["spectrum"].tolist(), expected_names)
        self.assertEqual(table["concentration_coefficient"].tolist(), [5, 3, 2, 4])
        self.assertEqual(table["Ph"].tolist(), [8.1, 6.8, 7.2, 6.5])
        np.testing.assert_allclose(
            summary.iloc[1:, -4:].to_numpy(dtype=float),
            [[1, 1, 2.5, 1], [2, 2, 4, 2], [3, 3, 5.5, 3]],
        )

    def test_new_concentration_table_follows_acquisition_datetime(self):
        self.write_spectrum("z_earlier", [1, 2, 3], time="06:00:00")
        merge.merge_spectra(self.folder)
        table = pd.read_csv(self.folder / "concentration_table.txt", sep="\t")
        self.assertEqual(table["spectrum"].tolist(), ["z_earlier", "sample"])
        self.assertEqual(table["concentration_coefficient"].tolist(), [1.0, 1.0])

    def test_equal_timestamps_use_filename_order(self):
        self.write_spectrum("zzz", [1, 2, 3])
        self.write_spectrum("Alpha", [1, 2, 3], time="6:41:36 PM")
        files = merge.find_spectral_files(self.folder)
        self.assertEqual([file.stem for file in files], ["Alpha", "sample", "zzz"])

    def test_acquisition_datetime_supports_24_hour_and_am_pm_times(self):
        for date, time, hour in (
            ("26/10/14", "18:41:36", 18),
            ("2026/10/14", "6:41:36 PM", 18),
            ("26/10/14", "12:41:36 AM", 0),
            ("26/10/14", "12:41:36 PM", 12),
        ):
            with self.subTest(date=date, time=time):
                self.write_spectrum("sample", [1, 2, 3], date=date, time=time)
                self.assertEqual(
                    merge.read_acquisition_datetime(self.folder / "sample.csv"),
                    datetime(2026, 10, 14, hour, 41, 36),
                )

    def test_invalid_acquisition_datetime_reports_file_and_cells(self):
        file = self.folder / "sample.csv"
        for date, time in (("", "18:41:36"), ("26/13/14", "18:41:36"), ("26/10/14", "25:00:00")):
            with self.subTest(date=date, time=time):
                self.write_spectrum("sample", [1, 2, 3], date=date, time=time)
                with self.assertRaisesRegex(ValueError, "Invalid acquisition date/time") as raised:
                    merge.find_spectral_files(self.folder)
                message = str(raised.exception)
                self.assertIn(str(file), message)
                self.assertIn(f"B5 (CSV line 5): {date!r}", message)
                self.assertIn(f"B6 (CSV line 6): {time!r}", message)

    def test_mismatched_wavelengths_do_not_write_workbook(self):
        self.write_spectrum("other sample", [5, 8, 11], wavelengths=(200, 201, 203))
        reference_file = self.folder / "sample.csv"
        actual_file = self.folder / "other sample.csv"
        with patch.object(merge, "find_spectral_files", return_value=[reference_file, actual_file]):
            with self.assertRaisesRegex(ValueError, "X axis") as raised:
                merge.merge_spectra(self.folder)
        message = str(raised.exception)
        self.assertIn(f"X axis mismatch in {actual_file}", message)
        self.assertIn(f"Reference file: {reference_file}", message)
        self.assertIn("Spectral data row: 3", message)
        self.assertIn("Expected: 202 nm (CSV line 10)", message)
        self.assertIn("Actual:   203 nm (CSV line 10)", message)
        self.assertIn("Data rows: expected 3, actual 3", message)
        self.assertFalse((self.folder.parent / "instrument data-combined.xlsx").exists())

    def test_mismatch_line_numbers_include_blank_lines(self):
        reference_file = self.folder / "sample.csv"
        actual_file = self.folder / "other sample.csv"
        actual_file.write_text(
            "Instrument export\nXYDATA\n\n200,5,500\n \n201,8,500\n"
            "203,11,500\n##### Extended Information\n",
            encoding="latin-1",
        )
        with self.assertRaises(ValueError) as raised:
            merge.validate_wavelength_axis(
                merge.read_spectrum(reference_file), merge.read_spectrum(actual_file),
                reference_file, actual_file,
            )
        self.assertIn("Expected: 202 nm (CSV line 10)", str(raised.exception))
        self.assertIn("Actual:   203 nm (CSV line 7)", str(raised.exception))

    def test_different_axis_lengths_report_the_first_missing_or_extra_row(self):
        reference_file = self.folder / "sample.csv"
        actual_file = self.folder / "other sample.csv"
        for wavelengths, row, expected, actual in (
            ((200, 201), 3, "202 nm (CSV line 10)",
             "<end of spectral data> (CSV line 10)"),
            ((200, 201, 202, 203), 4, "<end of spectral data> (CSV line 11)",
             "203 nm (CSV line 11)"),
        ):
            with self.subTest(wavelengths=wavelengths):
                self.write_spectrum("other sample", range(len(wavelengths)), wavelengths)
                with self.assertRaises(ValueError) as raised:
                    merge.validate_wavelength_axis(
                        merge.read_spectrum(reference_file), merge.read_spectrum(actual_file),
                        reference_file, actual_file,
                    )
                message = str(raised.exception)
                self.assertIn(f"Spectral data row: {row}", message)
                self.assertIn(f"Expected: {expected}", message)
                self.assertIn(f"Actual:   {actual}", message)
                self.assertIn(f"Data rows: expected 3, actual {len(wavelengths)}", message)

    def test_equivalent_integer_and_decimal_wavelengths_are_accepted(self):
        self.write_spectrum("other sample", [1, 2, 3], wavelengths=(200.0, 201.0, 202.0))
        output = merge.merge_spectra(self.folder)
        self.assertTrue(output.is_file())

    def test_four_column_exports_are_written_to_separate_channel_sheets(self):
        self.write_spectrum("sample", [5, 8, 11], absorbance=0.25)
        self.write_spectrum("other sample", [1, 2, 3], absorbance=0.5, ht=650)
        frame = merge.read_spectrum(self.folder / "sample.csv")
        self.assertEqual(frame["wavelength_nm"].tolist(), [200, 201, 202])
        self.assertEqual(frame["cd_mdeg"].tolist(), [5, 8, 11])
        self.assertEqual(frame["ht_v"].tolist(), [500, 500, 500])
        self.assertEqual(frame["absorbance"].tolist(), [0.25, 0.25, 0.25])
        self.assertEqual(frame.index.tolist(), [0, 1, 2])
        output = merge.merge_spectra(self.folder)
        with pd.ExcelFile(output) as workbook:
            self.assertEqual(workbook.sheet_names, ["CD", "HT", "Abs", "SUMMARY"])
            for sheet_name, expected_sample, expected_other in (
                ("CD", [5, 8, 11], [1, 2, 3]),
                ("HT", [500, 500, 500], [650, 650, 650]),
                ("Abs", [0.25, 0.25, 0.25], [0.5, 0.5, 0.5]),
            ):
                with self.subTest(sheet_name=sheet_name):
                    channel = pd.read_excel(workbook, sheet_name=sheet_name)
                    self.assertEqual(channel["wavelength_nm"].tolist(), [200, 201, 202])
                    self.assertEqual(channel["sample"].tolist(), expected_sample)
                    self.assertEqual(channel["other sample"].tolist(), expected_other)

    def test_three_column_exports_leave_absorbance_blank(self):
        self.write_spectrum("sample", [5, 8, 11], absorbance=None)
        output = merge.merge_spectra(self.folder)
        absorbance = pd.read_excel(output, sheet_name="Abs")
        self.assertEqual(absorbance["wavelength_nm"].tolist(), [200, 201, 202])
        self.assertTrue(absorbance["sample"].isna().all())

    def test_all_filenames_are_treated_as_samples(self):
        self.write_spectrum("water", [1, 2, 3])
        self.write_spectrum("Water blank", [2, 4, 6])
        self.write_coefficients("sample\t2\nwater\t4\nWater blank\t2\n")
        self.assertEqual(len(merge.find_spectral_files(self.folder)), 3)

        output = merge.merge_spectra(self.folder)
        summary = pd.read_excel(output, sheet_name="SUMMARY", header=None)
        headers = summary.iloc[0].tolist()
        for name, expected in (
            ("sample", [2.5, 4, 5.5]),
            ("water", [0.25, 0.5, 0.75]),
            ("Water blank", [1, 2, 3]),
        ):
            with self.subTest(name=name):
                columns = [index for index, header in enumerate(headers) if header == name]
                self.assertEqual(len(columns), 2)
                np.testing.assert_allclose(
                    summary.iloc[1:, columns[-1]].to_numpy(dtype=float), expected
                )

    def test_invalid_concentration_coefficients_are_rejected(self):
        for coefficient in ("0", "-1", "NaN", "inf", "-inf"):
            with self.subTest(coefficient=coefficient):
                self.write_coefficients(f"sample\t{coefficient}\n")
                with self.assertRaisesRegex(ValueError, "finite and greater than zero"):
                    merge.load_concentration_coefficients(self.folder, ["sample"])

    def test_concentration_table_must_match_samples(self):
        for rows, error in (
            ("sample\t1\nsample\t2\n", "Duplicate spectra"),
            ("other\t1\n", "missing: sample; not present in data folder: other"),
        ):
            with self.subTest(rows=rows):
                self.write_coefficients(rows)
                with self.assertRaisesRegex(ValueError, error):
                    merge.load_concentration_coefficients(self.folder, ["sample"])

    def test_missing_markers_and_empty_blocks_are_rejected(self):
        file = self.folder / "malformed.csv"
        for content, error in (
            ("200,1,500\n##### Extended Information\n", "XYDATA marker"),
            ("XYDATA\n200,1,500\n", "Extended Information marker"),
            ("XYDATA\n\n##### Extended Information\n", "No spectral data"),
        ):
            with self.subTest(content=content):
                file.write_text(content, encoding="latin-1")
                with self.assertRaisesRegex(ValueError, error):
                    merge.read_spectrum(file)

    def test_cli_accepts_windows_folder_with_spaces_and_trailing_separator(self):
        result = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "merge.py"),
             "--folder", str(self.folder) + "\\"],
            capture_output=True, text=True, check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Saved ", result.stdout)

    def test_cli_reports_missing_folder_without_traceback(self):
        result = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "merge.py"),
             "--folder", str(self.folder / "missing")],
            capture_output=True, text=True, check=False,
        )
        self.assertEqual(result.returncode, 1)
        self.assertIn("ERROR: Data folder not found", result.stderr)
        self.assertNotIn("Traceback", result.stderr)


class AnalysisTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="ECD analysis ")
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name)
        wavelengths = np.linspace(190, 260, 20)
        components = np.vstack([
            np.exp(-((wavelengths - 210) / 12) ** 2),
            np.exp(-((wavelengths - 240) / 10) ** 2),
        ])
        spectra = np.array([[1.0, 0.2], [0.4, 1.0], [0.8, 0.8]]) @ components
        self.frame = pd.DataFrame(spectra.T, columns=["sample 1", "sample 2", "sample 3"])
        self.frame.insert(0, "wavelength_nm", wavelengths)
        self.workbook = self.folder / "test spectra.xlsx"
        self.frame.to_excel(self.workbook, index=False, sheet_name="Sheet1")

    def test_outside_range_reports_observed_bounds_with_one_workbook_read(self):
        with patch.object(analysis.pd, "read_excel", wraps=pd.read_excel) as read_excel:
            with self.assertRaisesRegex(ValueError, "190-260 nm"):
                analysis.load_spectral_dataset(self.workbook, "Sheet1", 300, 400)
        self.assertEqual(read_excel.call_count, 1)

    def test_cli_exports_headless_pca_and_mcr_results(self):
        self.run_analysis("positive", expected_csv=8, expected_png=9)

    def test_cli_skips_mcr_for_signed_ecd_data(self):
        self.frame.iloc[:, 1:] -= 0.5
        self.frame.to_excel(self.workbook, index=False, sheet_name="Sheet1")
        result = self.run_analysis("signed", expected_csv=5, expected_png=6)
        self.assertIn("MCR-ALS was skipped", result.stderr)

    def run_analysis(self, name, expected_csv, expected_png):
        output = self.folder / name
        result = subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "spectral_component_analysis.py"),
             "--file", str(self.workbook), "--components", "2",
             "--output", str(output), "--no-show"],
            env={**os.environ, "MPLBACKEND": "Agg"},
            capture_output=True, text=True, check=False,
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(len(list(output.glob("*.csv"))), expected_csv)
        self.assertEqual(len(list(output.glob("*.png"))), expected_png)
        reconstructed = pd.read_csv(output / "pca_reconstructed_spectra.csv")
        np.testing.assert_allclose(reconstructed, self.frame, atol=1e-12)
        return result


if __name__ == "__main__":
    unittest.main()
