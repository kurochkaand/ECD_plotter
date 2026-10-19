"""Check CD unit conversion, merge integration, and safe workbook updates."""

import io
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import load_workbook

import deconvolution
import merge


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class DeconvolutionTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="ECD MRE ")
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name).resolve() / "sample data"
        self.folder.mkdir()
        self.write_spectrum("peptide", [-20, 0, 20])
        self.parameters = dict(path_in_mm=1.0, num_of_aa=20, concentration_in_M=0.0001)

    def write_spectrum(self, name, values, ht=(400, 400, 400)):
        rows = "".join(
            f"{wavelength},{signal},{voltage},0.1\n"
            for wavelength, signal, voltage in zip((200, 201, 202), values, ht)
        )
        (self.folder / f"{name}.csv").write_text(
            "TITLE,Instrument export\nDATA TYPE,\nORIGIN,JASCO\nOWNER,\n"
            "DATE,26/10/14\nTIME,18:41:36\nXYDATA\n" + rows
            + "##### Extended Information\n", encoding="latin-1",
        )

    def open_workbook(self, path, **kwargs):
        workbook = load_workbook(path, **kwargs)
        self.addCleanup(workbook.close)
        return workbook

    def test_normalization_units_and_existing_sheets_are_preserved(self):
        (self.folder / "concentration_table.txt").write_text(
            "spectrum\tconcentration_coefficient\npeptide\t2\n", encoding="utf-8"
        )
        output = merge.merge_spectra(self.folder)
        original = self.open_workbook(output)
        notes = original.create_sheet("Notes")
        notes["A1"] = "Keep this"
        notes["B1"] = "=1+2"
        notes["A1"].number_format = "@"
        original.save(output)
        before = {sheet.title: list(sheet.values) for sheet in original}

        self.assertEqual(deconvolution.convert_to_mre(self.folder, **self.parameters), output)
        workbook = self.open_workbook(output)
        for name, rows in before.items():
            self.assertEqual(list(workbook[name].values), rows)
        self.assertEqual(workbook["Notes"]["A1"].number_format, "@")
        # -20 mdeg / dilution coefficient 2 / (1 mm * 0.0001 M * 20) = -5000.
        self.assertEqual(list(workbook["MRE"].values), [
            ("wavelength_nm", "peptide"), (200, -5000), (201, 0), (202, 5000),
        ])
        metadata = dict(list(workbook["MRE_parameters"].values)[1:])
        self.assertEqual(metadata["concentration_in_M"], 0.0001)
        self.assertEqual(metadata["path_in_mm"], 1)
        self.assertEqual(metadata["num_of_aa"], 20)
        self.assertIn("dmol^-1", metadata["MRE_units"])

    def test_subtraction_works_without_excel_cached_results(self):
        self.write_spectrum("H2O", [2, 4, 6])
        self.write_spectrum("second", [10, 12, 14])
        (self.folder / "concentration_table.txt").write_text(
            "spectrum\tconcentration_coefficient\nH2O\t1\npeptide\t0.5\nsecond\t1\n",
            encoding="utf-8",
        )
        output = merge.merge_spectra(self.folder, subtract="H2O")
        cached = self.open_workbook(output, data_only=True)
        self.assertIsNone(cached["CD_subtr"]["F2"].value)
        original = self.open_workbook(output)
        formulas = list(original["CD_subtr"].values)

        deconvolution.convert_to_mre(self.folder, **self.parameters)
        workbook = self.open_workbook(output)
        self.assertEqual(list(workbook["CD_subtr"].values), formulas)
        self.assertEqual(list(workbook["MRE"].values), [
            ("wavelength_nm", "peptide", "second"),
            (200, -22000, 4000), (201, -4000, 4000), (202, 14000, 4000),
        ])

    def test_ht_exclusions_remain_blank_with_subtraction_and_smoothing(self):
        self.write_spectrum("peptide", [-20, 0, 20], ht=(600, 400, 400))
        self.write_spectrum("H2O", [2, 4, 6], ht=(400, 600, 400))
        output = merge.merge_spectra(self.folder, subtract="H2O", ht_threshold=500, smooth=3)
        deconvolution.convert_to_mre(self.folder, **self.parameters)
        workbook = self.open_workbook(output)
        self.assertEqual(list(workbook["MRE"].values), [
            ("wavelength_nm", "peptide"), (200, None), (201, None), (202, 7000),
        ])

    def test_rerun_replaces_outputs_with_updated_parameters(self):
        output = merge.merge_spectra(self.folder)
        deconvolution.convert_to_mre(self.folder, **self.parameters)
        self.parameters["path_in_mm"] = 2
        deconvolution.convert_to_mre(self.folder, **self.parameters)
        workbook = self.open_workbook(output)
        self.assertEqual(workbook.sheetnames.count("MRE"), 1)
        self.assertEqual(workbook.sheetnames.count("MRE_parameters"), 1)
        self.assertEqual(workbook["MRE"]["B2"].value, -5000)
        self.assertEqual(dict(list(workbook["MRE_parameters"].values)[1:])["path_in_mm"], 2)

    def test_fully_excluded_spectra_keep_blank_columns(self):
        self.write_spectrum("peptide", [-20, 0, 20], ht=(600, 600, 600))
        output = merge.merge_spectra(self.folder, ht_threshold=500)
        deconvolution.convert_to_mre(self.folder, **self.parameters)
        workbook = self.open_workbook(output)
        self.assertEqual(list(workbook["MRE"].values), [
            ("wavelength_nm", "peptide"), (200, None), (201, None), (202, None),
        ])

    def test_missing_corrected_sheet_or_block_suggests_merge(self):
        for problem in ("sheet", "separator", "spectra"):
            with self.subTest(problem=problem):
                output = merge.merge_spectra(self.folder)
                workbook = self.open_workbook(output)
                if problem == "sheet":
                    del workbook["CD_subtr"]
                elif problem == "separator":
                    workbook["CD_subtr"]["C1"] = "other"
                else:
                    workbook["CD_subtr"].delete_cols(4)
                workbook.save(output)
                original_bytes = output.read_bytes()
                with self.assertRaises(ValueError):
                    deconvolution.convert_to_mre(self.folder, **self.parameters)
                self.assertEqual(output.read_bytes(), original_bytes)

    def test_invalid_parameters_leave_workbook_unchanged(self):
        output = merge.merge_spectra(self.folder)
        original_bytes = output.read_bytes()
        for parameter, values in (
            ("path_in_mm", [0, -1, float("nan"), float("inf")]),
            ("concentration_in_M", [0, -1, float("nan"), float("inf")]),
            ("num_of_aa", [0, -1, 2.5, True]),
        ):
            for value in values:
                with self.subTest(parameter=parameter, value=value):
                    parameters = {**self.parameters, parameter: value}
                    with self.assertRaises(ValueError):
                        deconvolution.convert_to_mre(self.folder, **parameters)
                    self.assertEqual(output.read_bytes(), original_bytes)

    def test_invalid_data_and_unsupported_formulas_leave_workbook_unchanged(self):
        for value in ("bad CD", "#DIV/0!", "=SUM(B2)", "=B2-$B3", "=(B2-$B2)/0.0"):
            with self.subTest(value=value):
                output = merge.merge_spectra(self.folder)
                workbook = self.open_workbook(output)
                workbook["CD_subtr"]["D2"] = value
                workbook.save(output)
                original_bytes = output.read_bytes()
                with self.assertRaisesRegex(ValueError, "CD_subtr!D2"):
                    deconvolution.convert_to_mre(self.folder, **self.parameters)
                self.assertEqual(output.read_bytes(), original_bytes)

    def test_failed_replace_preserves_source_and_cleans_temporary_workbook(self):
        output = merge.merge_spectra(self.folder)
        original_bytes = output.read_bytes()
        with patch.object(deconvolution.os, "replace", side_effect=PermissionError("Workbook is open")):
            with self.assertRaises(PermissionError):
                deconvolution.convert_to_mre(self.folder, **self.parameters)
        self.assertEqual(output.read_bytes(), original_bytes)
        self.assertEqual(list(output.parent.glob(".*-combined-*.xlsx")), [])

    def run_cli(self, folder=None):
        return subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "deconvolution.py"),
             "--folder", str(folder or self.folder), "--path-in-mm", "1",
             "--num-of-aa", "20", "--concentration-in-mM", "0.1", "--mre-only"],
            capture_output=True, text=True, check=False,
        )

    def test_cli_accepts_millimolar_concentration_and_folder_with_spaces(self):
        output = merge.merge_spectra(self.folder)
        result = self.run_cli()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("Saved MRE", result.stdout)
        self.assertEqual(self.open_workbook(output, data_only=True)["MRE"]["B2"].value, -10000)

    def test_cli_missing_workbook_suggests_merge(self):
        result = self.run_cli()
        self.assertEqual(result.returncode, 1)
        self.assertIn(str(self.folder.parent / "sample data-combined.xlsx"), result.stderr)
        self.assertIn("Run merge.py first", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_all_cli_flags_are_required(self):
        arguments = [
            "--folder", str(self.folder), "--path-in-mm", "1",
            "--num-of-aa", "20", "--concentration-in-mM", "0.1",
        ]
        for position in range(0, len(arguments), 2):
            with self.subTest(flag=arguments[position]):
                stderr = io.StringIO()
                with patch.object(sys, "argv", ["deconvolution.py", *arguments[:position], *arguments[position + 2:]]):
                    with patch("sys.stderr", stderr), self.assertRaises(SystemExit) as error:
                        deconvolution.main()
                self.assertEqual(error.exception.code, 2)
                self.assertIn(arguments[position], stderr.getvalue())


if __name__ == "__main__":
    unittest.main()
