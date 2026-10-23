"""Check MRE SVD, common-wavelength filtering, and safe workbook updates."""

import io
import os
import subprocess
import sys
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest.mock import patch

import matplotlib
import numpy as np
from openpyxl import Workbook, load_workbook

matplotlib.use("Agg")
import matplotlib.pyplot as plt

import component_analysis_mre as analysis
import spectral_component_analysis as original_analysis


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class MreSvdTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="ECD MRE SVD ")
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name).resolve() / "raw data"
        self.folder.mkdir()
        self.workbook_file = self.folder.parent / "raw data-combined.xlsx"
        self.wavelengths = np.arange(185, 256, dtype=float)
        components = np.vstack([
            -10000 * np.exp(-((self.wavelengths - 208) / 10) ** 2),
            8000 * np.exp(-((self.wavelengths - 235) / 9) ** 2),
        ])
        fraction = np.linspace(0.05, 0.95, 8)
        self.spectra = np.column_stack([fraction, 1 - fraction]) @ components
        # A noise estimator needs a noise floor above floating-point roundoff.
        self.spectra += np.random.default_rng(17).normal(0, 0.1, self.spectra.shape)
        self.names = [f"sample {index + 1}" for index in range(len(fraction))]
        self.write_workbook()

    def write_workbook(self, *, wavelengths=None, spectra=None, names=None):
        wavelengths = self.wavelengths if wavelengths is None else wavelengths
        spectra = self.spectra if spectra is None else spectra
        names = self.names if names is None else names
        workbook = Workbook()
        raw = workbook.active
        raw.title = "CD_subtr"
        raw.append(["measured", "reference", "difference"])
        raw.append([5, 2, "=A2-B2"])
        source = workbook.create_sheet("MRE")
        source.append(["wavelength_nm", *names])
        for wavelength, signals in zip(wavelengths, spectra.T):
            source.append([float(wavelength), *signals.tolist()])
        for title in ("MRE_parameters", "SESCA_results", "SESCA_fit", "Epsilon"):
            workbook.create_sheet(title).append(["keep", title])
        workbook.save(self.workbook_file)
        workbook.close()

    def open_workbook(self):
        workbook = load_workbook(self.workbook_file, data_only=False)
        self.addCleanup(workbook.close)
        return workbook

    def run_analysis(self, **kwargs):
        with redirect_stdout(io.StringIO()) as output:
            result = analysis.analyze_mre_components(self.folder, **kwargs)
        return result, output.getvalue()

    def parameters(self, workbook):
        return dict(list(workbook[analysis.PARAMETERS_SHEET].values)[1:])

    def run_cli(self, *arguments):
        return subprocess.run(
            [sys.executable, str(PROJECT_ROOT / "component_analysis_mre.py"),
             "--folder", str(self.folder) + "\\", *arguments],
            env={**os.environ, "MPLBACKEND": "Agg"},
            capture_output=True, text=True, check=False,
        )

    def test_known_two_component_signed_mre_uses_defaults_without_raw_files_or_table(self):
        self.assertEqual(list(self.folder.iterdir()), [])
        with patch.object(analysis, "plot_svd") as plot:
            result, output = self.run_analysis()
        plot.assert_not_called()
        self.assertEqual(result, self.workbook_file)
        workbook = self.open_workbook()
        parameters = self.parameters(workbook)
        self.assertEqual(parameters["gavish_donoho_rank"], 2)
        self.assertEqual(parameters["spectral_energy_rank"], 2)
        self.assertEqual(parameters["requested_min_nm"], 190)
        self.assertEqual(parameters["requested_max_nm"], 250)
        self.assertEqual(parameters["retained_wavelength_rows"], 61)
        self.assertEqual(parameters["excluded_missing_rows"], 0)
        self.assertEqual(parameters["outside_range_rows"], 10)
        self.assertFalse(parameters["mean_centered"])
        self.assertIn("Gavish-Donoho rank: 2", output)
        fractions = [row[2] for row in list(workbook[analysis.RESULTS_SHEET].values)[1:]]
        self.assertAlmostEqual(sum(fractions), 1)

    def test_saved_input_preserves_mre_units_and_original_sheets_and_formulas(self):
        source = self.open_workbook()
        original = {sheet.title: list(sheet.values) for sheet in source}
        source.close()
        self.run_analysis()
        workbook = self.open_workbook()
        for title, values in original.items():
            self.assertEqual(list(workbook[title].values), values)
        inputs = list(workbook[analysis.INPUT_SHEET].values)
        self.assertEqual(inputs[0], ("wavelength_nm", *self.names))
        expected = (self.wavelengths >= 190) & (self.wavelengths <= 250)
        np.testing.assert_allclose(np.asarray(inputs[1:])[:, 1:], self.spectra[:, expected].T)
        self.assertEqual(workbook["CD_subtr"]["C2"].value, "=A2-B2")

    def test_blanks_exclude_entire_rows_and_descending_input_is_sorted(self):
        spectra = self.spectra.astype(object)
        for wavelength, spectrum in ((190, 0), (201, 1), (250, 2), (185, 3)):
            spectra[spectrum, np.flatnonzero(self.wavelengths == wavelength)[0]] = None
        self.write_workbook(wavelengths=self.wavelengths[::-1], spectra=spectra[:, ::-1])
        self.run_analysis()
        workbook = self.open_workbook()
        parameters = self.parameters(workbook)
        self.assertEqual(parameters["excluded_missing_rows"], 3)
        self.assertEqual(parameters["retained_wavelength_rows"], 58)
        self.assertEqual(parameters["retained_min_nm"], 191)
        self.assertEqual(parameters["retained_max_nm"], 249)
        inputs = np.asarray(list(workbook[analysis.INPUT_SHEET].values)[1:], dtype=float)
        self.assertTrue(np.isfinite(inputs).all())
        self.assertTrue(np.all(np.diff(inputs[:, 0]) > 0))
        self.assertNotIn(201, inputs[:, 0])

    def test_rank_is_uncentered_even_if_old_script_centering_is_enabled(self):
        workbook = self.open_workbook()
        with patch.object(original_analysis, "MEAN_CENTER_FOR_RANK", True):
            result = analysis.analyze_mre(workbook["MRE"])
            centered = original_analysis.analyze_singular_values(result.spectra)
        self.assertEqual(result.svd.variance_rank, 2)
        self.assertEqual(centered.variance_rank, 1)

    def test_all_singular_values_and_energy_rank_are_reported_beyond_four(self):
        wavelengths = np.arange(190, 251, dtype=float)
        spectra = np.vstack([
            10000 * np.sin(np.pi * index * (wavelengths - 190) / 60)
            for index in range(1, 7)
        ])
        self.write_workbook(wavelengths=wavelengths, spectra=spectra, names=self.names[:6])
        self.run_analysis()
        workbook = self.open_workbook()
        self.assertEqual(workbook[analysis.RESULTS_SHEET].max_row, 7)
        self.assertEqual(self.parameters(workbook)["spectral_energy_rank"], 6)

    def test_rerun_replaces_only_svd_outputs_and_updates_range(self):
        self.run_analysis()
        self.run_analysis(min_nm=200, max_nm=220)
        workbook = self.open_workbook()
        for title in (analysis.RESULTS_SHEET, analysis.PARAMETERS_SHEET, analysis.INPUT_SHEET):
            self.assertEqual(workbook.sheetnames.count(title), 1)
        parameters = self.parameters(workbook)
        self.assertEqual(parameters["retained_wavelength_rows"], 21)
        self.assertEqual(parameters["requested_min_nm"], 200)
        self.assertEqual(parameters["requested_max_nm"], 220)
        self.assertIn("SESCA_results", workbook.sheetnames)
        self.assertEqual(workbook["MRE"].max_row, 72)

    def test_plot_uses_full_diagnostics_and_runs_after_successful_save(self):
        def inspect_plot(*args, **kwargs):
            saved = self.open_workbook()
            self.assertIn(analysis.RESULTS_SHEET, saved.sheetnames)
            figures = [plt.figure(number) for number in plt.get_fignums()]
            singular_figures = [
                figure for figure in figures
                if any(axis.get_title() == "MRE singular values" for axis in figure.axes)
            ]
            self.assertEqual(len(singular_figures), 1)
            figure = singular_figures[0]
            singular_axis = next(
                axis for axis in figure.axes if axis.get_title() == "MRE singular values"
            )
            np.testing.assert_array_equal(singular_axis.lines[0].get_xdata(), np.arange(1, 9))
            self.assertIn("190-250 nm", figure._suptitle.get_text())
            self.assertGreater(len(figures), 1)

        with patch.object(plt, "show", side_effect=inspect_plot) as show:
            self.run_analysis(plot=True, reconstruction_ranks=[2, 5])
        show.assert_called_once_with(block=True)
        self.assertEqual(plt.get_fignums(), [])

    def test_plot_preserves_figures_created_by_the_caller(self):
        figure = plt.figure()
        self.addCleanup(plt.close, figure)
        with patch.object(plt, "show") as show:
            self.run_analysis(plot=True)
        show.assert_called_once_with(block=True)
        self.assertEqual(plt.get_fignums(), [figure.number])

    def test_diagnostic_sheets_are_replaced_without_removing_other_analyses(self):
        self.run_analysis(reconstruction_ranks=[2, 5])
        titles = (
            analysis.PATTERNS_SHEET, analysis.SCORES_SHEET, analysis.CENTERED_SHEET,
            analysis.RECONSTRUCTION_SHEET, analysis.COMPARISON_SHEET,
            analysis.NOISE_SHEET, analysis.NOISE_WAVELENGTH_SHEET,
        )
        workbook = self.open_workbook()
        for title in titles:
            self.assertIn(title, workbook.sheetnames)
            workbook[title]["A10000"] = "stale output"
        workbook.save(self.workbook_file)
        workbook.close()

        self.run_analysis(reconstruction_ranks=[1, 2], diagnostic_components=3)
        workbook = self.open_workbook()
        for title in titles:
            self.assertEqual(workbook.sheetnames.count(title), 1)
            self.assertLess(workbook[title].max_row, 10000)
        self.assertEqual(workbook["CD_subtr"]["C2"].value, "=A2-B2")
        self.assertEqual(list(workbook["SESCA_results"].values), [("keep", "SESCA_results")])

    def test_default_run_skips_all_optional_calculations_and_sheets(self):
        with patch.object(analysis, "decompose_spectra") as decompose, \
                patch.object(analysis, "reconstruct_spectra") as reconstruct:
            _, output = self.run_analysis(diagnostic_components=3)
        decompose.assert_not_called()
        reconstruct.assert_not_called()
        workbook = self.open_workbook()
        self.assertEqual(
            set(workbook.sheetnames) & set(analysis.OUTPUT_SHEETS), set(analysis.BASE_SHEETS),
        )
        self.assertFalse(self.parameters(workbook)["diagnostics_enabled"])
        self.assertNotIn("centered_variance_rank", self.parameters(workbook))
        self.assertNotIn("Mean-centered ranks", output)
        self.assertNotIn("reconstruction:", output)
        self.assertNotIn("Residual noise checks", output)

    def test_default_plot_shows_only_svd_rank_figure(self):
        def inspect_plot(*args, **kwargs):
            self.assertEqual(len(plt.get_fignums()), 1)
            self.assertEqual(len(plt.gcf().axes), 2)
            self.assertEqual(plt.gcf().axes[0].get_title(), "MRE singular values")
            self.assertEqual(plt.gcf().axes[1].get_title(), "Explained spectral energy")

        with patch.object(plt, "show", side_effect=inspect_plot):
            self.run_analysis(plot=True)
        self.assertEqual(plt.get_fignums(), [])

    def test_basic_rerun_clears_previous_optional_outputs(self):
        self.run_analysis(reconstruction_ranks=[2, 5])
        self.run_analysis(min_nm=200, max_nm=220)
        workbook = self.open_workbook()
        self.assertTrue(set(analysis.BASE_SHEETS).issubset(workbook.sheetnames))
        self.assertTrue(set(analysis.DIAGNOSTIC_SHEETS).isdisjoint(workbook.sheetnames))
        self.assertFalse(self.parameters(workbook)["diagnostics_enabled"])
        self.assertEqual(self.parameters(workbook)["retained_wavelength_rows"], 21)
        self.assertEqual(list(workbook["SESCA_results"].values), [("keep", "SESCA_results")])
        self.assertEqual(workbook["CD_subtr"]["C2"].value, "=A2-B2")

    def test_cli_without_reconstruction_flag_runs_only_svd(self):
        result = self.run_cli("--plot")
        self.assertEqual(result.returncode, 0, result.stderr)
        workbook = self.open_workbook()
        self.assertFalse(self.parameters(workbook)["diagnostics_enabled"])
        self.assertTrue(set(analysis.DIAGNOSTIC_SHEETS).isdisjoint(workbook.sheetnames))
        self.assertIn("uncentered SVD only", result.stdout)
        self.assertNotIn("Mean-centered ranks", result.stdout)

    def test_cli_supports_reconstruction_ranks_and_diagnostic_component_count(self):
        result = self.run_cli(
            "--diagnostic-components", "3", "--reconstruction-ranks", "1", "2", "2",
        )
        self.assertEqual(result.returncode, 0, result.stderr)
        workbook = self.open_workbook()
        self.assertIn(analysis.COMPARISON_SHEET, workbook.sheetnames)
        self.assertIn(analysis.PATTERNS_SHEET, workbook.sheetnames)
        self.assertTrue(self.parameters(workbook)["diagnostics_enabled"])

    def test_exported_scores_patterns_and_reconstructions_match_saved_inputs(self):
        self.run_analysis(reconstruction_ranks=[2, 5], diagnostic_components=1)
        workbook = self.open_workbook()
        inputs = list(workbook[analysis.INPUT_SHEET].values)
        wavelengths = np.asarray([row[0] for row in inputs[1:]])
        spectra = np.asarray([row[1:] for row in inputs[1:]]).T
        patterns = list(workbook[analysis.PATTERNS_SHEET].values)
        scores = list(workbook[analysis.SCORES_SHEET].values)
        self.assertEqual(patterns[0], ("mode", "wavelength_nm", "offset_MRE", *[
            f"SV{i + 1}" for i in range(len(self.names))
        ]))
        self.assertEqual(scores[0], ("mode", "spectrum", *[
            f"SV{i + 1}" for i in range(len(self.names))
        ]))
        for mode, component_count in (("uncentered", 8), ("centered", 7)):
            with self.subTest(mode=mode):
                mode_patterns = [row for row in patterns[1:] if row[0] == mode]
                mode_scores = [row for row in scores[1:] if row[0] == mode]
                self.assertEqual([row[1] for row in mode_scores], self.names)
                np.testing.assert_array_equal([row[1] for row in mode_patterns], wavelengths)
                loading_matrix = np.asarray([row[3:3 + component_count] for row in mode_patterns]).T
                score_matrix = np.asarray([row[2:2 + component_count] for row in mode_scores])
                offsets = np.asarray([row[2] for row in mode_patterns])
                np.testing.assert_allclose(score_matrix @ loading_matrix + offsets, spectra, atol=1e-9)
                if mode == "uncentered":
                    uncentered_scores, uncentered_patterns = score_matrix, loading_matrix
                else:
                    self.assertTrue(all(row[-1] is None for row in mode_patterns))
                    self.assertTrue(all(row[-1] is None for row in mode_scores))

        reconstructions = list(workbook[analysis.RECONSTRUCTION_SHEET].values)
        self.assertEqual(reconstructions[0], (
            "rank", "spectrum", "wavelength_nm", "observed_MRE", "reconstructed_MRE", "residual_MRE",
        ))
        for rank in (2, 5):
            model_rows = [row for row in reconstructions[1:] if row[0] == rank]
            self.assertEqual(len(model_rows), spectra.size)
            expected = uncentered_scores[:, :rank] @ uncentered_patterns[:rank]
            for spectrum_index, name in enumerate(self.names):
                rows = [row for row in model_rows if row[1] == name]
                np.testing.assert_array_equal([row[2] for row in rows], wavelengths)
                np.testing.assert_allclose([row[3] for row in rows], spectra[spectrum_index], atol=1e-9)
                np.testing.assert_allclose([row[4] for row in rows], expected[spectrum_index], atol=1e-9)
                np.testing.assert_allclose(
                    [row[5] for row in rows], spectra[spectrum_index] - expected[spectrum_index], atol=1e-9,
                )
        self.assertEqual(list(workbook[analysis.CENTERED_SHEET].values)[0], (
            "component", "singular_value_MRE", "variance_fraction",
            "cumulative_variance_fraction", "above_noise_threshold",
        ))

    def test_exported_noise_correlations_are_blank_and_explained_when_undefined(self):
        spectra = np.zeros((2, 6))
        spectra[0, 0] = spectra[1, 1] = 1
        cases = (
            (np.arange(190, 196), spectra, "negligible or constant"),
            (np.array([190, 191, 193, 194, 197, 199]), self.spectra[:, :6], "irregular"),
        )
        for wavelengths, values, reason in cases:
            with self.subTest(reason=reason):
                names = self.names[:len(values)]
                self.write_workbook(wavelengths=wavelengths, spectra=values, names=names)
                self.run_analysis(reconstruction_ranks=[2])
                rows = list(self.open_workbook()[analysis.NOISE_SHEET].values)
                self.assertEqual(rows[0], (
                    "rank", "spectrum", "residual_mean_MRE", "residual_rms_MRE",
                    "lag1_correlation", "lag1_pairs", "lag1_status",
                ))
                self.assertEqual(len(rows) - 1, len(names))
                for row in rows[1:]:
                    self.assertIsNone(row[4])
                    self.assertIn(reason, row[6])

    def test_roundoff_residuals_do_not_receive_noise_variation_coefficients(self):
        self.run_analysis(reconstruction_ranks=[len(self.names)])
        rows = list(self.open_workbook()[analysis.COMPARISON_SHEET].values)
        headers, values = rows
        self.assertLess(values[headers.index("relative_Frobenius_error")], 1e-13)
        self.assertIsNone(values[headers.index("spectrum_residual_rms_CV")])
        self.assertIsNone(values[headers.index("wavelength_residual_rms_CV")])

    def test_invalid_diagnostic_options_preserve_the_source_workbook(self):
        original = self.workbook_file.read_bytes()
        for arguments in (
            ["--reconstruction-ranks", "0"],
            ["--reconstruction-ranks", "9"],
            ["--diagnostic-components", "0"],
        ):
            with self.subTest(arguments=arguments):
                result = self.run_cli(*arguments)
                self.assertNotEqual(result.returncode, 0)
                self.assertNotIn("Traceback", result.stderr)
                self.assertEqual(self.workbook_file.read_bytes(), original)

    def test_cli_supports_folder_with_spaces_trailing_separator_and_range_overrides(self):
        result = self.run_cli("--min-wavelength", "200", "--max-wavelength", "220")
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertIn("200-220 nm", result.stdout)
        self.assertIn("MRE_SVD_results", result.stdout)
        self.assertEqual(self.parameters(self.open_workbook())["retained_wavelength_rows"], 21)

    def test_invalid_ranges_preserve_existing_outputs(self):
        self.run_analysis()
        original = self.workbook_file.read_bytes()
        for minimum, maximum in ((250, 190), (190, 190), (0, 250), (np.nan, 250), (190, np.inf)):
            with self.subTest(minimum=minimum, maximum=maximum):
                with self.assertRaisesRegex(ValueError, "bounds"):
                    self.run_analysis(min_nm=minimum, max_nm=maximum)
                self.assertEqual(self.workbook_file.read_bytes(), original)

    def test_outside_range_and_insufficient_complete_rows_preserve_workbook(self):
        original = self.workbook_file.read_bytes()
        with self.assertRaisesRegex(ValueError, "workbook spans 185-255"):
            self.run_analysis(min_nm=300, max_nm=400)
        self.assertEqual(self.workbook_file.read_bytes(), original)

        spectra = self.spectra.astype(object)
        spectra[0, :] = None
        self.write_workbook(spectra=spectra)
        original = self.workbook_file.read_bytes()
        with self.assertRaisesRegex(ValueError, "Only 0 complete.*excluding 61"):
            self.run_analysis()
        self.assertEqual(self.workbook_file.read_bytes(), original)

    def test_zero_data_and_single_spectrum_are_rejected(self):
        for spectra, names, message in (
            (np.zeros_like(self.spectra), self.names, "zero numerical rank"),
            (self.spectra[:1], self.names[:1], "at least two spectra"),
        ):
            with self.subTest(message=message):
                self.write_workbook(spectra=spectra, names=names)
                original = self.workbook_file.read_bytes()
                with self.assertRaisesRegex(ValueError, message):
                    self.run_analysis()
                self.assertEqual(self.workbook_file.read_bytes(), original)

    def test_invalid_mre_values_layout_and_duplicate_wavelengths_are_rejected(self):
        for cell, value, message in (
            ("B7", "=B6", "MRE!B7"),
            ("B7", "#DIV/0!", "MRE!B7"),
            ("B7", "bad MRE", "MRE!B7"),
            ("A1", "wrong", "wavelength_nm"),
            ("C1", self.names[0], "Duplicate spectrum names"),
            ("B1", " ", "nonempty names"),
            ("A8", 190, "wavelengths must be unique"),
        ):
            with self.subTest(cell=cell, value=value):
                self.write_workbook()
                workbook = self.open_workbook()
                workbook["MRE"][cell] = value
                workbook.save(self.workbook_file)
                workbook.close()
                original = self.workbook_file.read_bytes()
                with self.assertRaisesRegex(ValueError, message):
                    self.run_analysis()
                self.assertEqual(self.workbook_file.read_bytes(), original)

    def test_missing_mre_or_workbook_reports_merge_command_without_traceback(self):
        workbook = self.open_workbook()
        del workbook["MRE"]
        workbook.save(self.workbook_file)
        workbook.close()
        original = self.workbook_file.read_bytes()
        result = self.run_cli()
        self.assertEqual(result.returncode, 1)
        self.assertIn("no MRE sheet", result.stderr)
        self.assertIn("Run merge.py first", result.stderr)
        self.assertNotIn("Traceback", result.stderr)
        self.assertEqual(self.workbook_file.read_bytes(), original)
        self.workbook_file.unlink()
        result = self.run_cli()
        self.assertEqual(result.returncode, 1)
        self.assertIn(str(self.workbook_file), result.stderr)
        self.assertIn("Run merge.py first", result.stderr)
        self.assertNotIn("Traceback", result.stderr)

    def test_failed_save_or_replacement_preserves_source_and_cleans_temporary_files(self):
        for target in ("save", "replace"):
            with self.subTest(target=target):
                original = self.workbook_file.read_bytes()
                owner = analysis.os if target == "replace" else Workbook
                with patch.object(owner, target, side_effect=PermissionError("Workbook is open")):
                    with self.assertRaises(PermissionError):
                        self.run_analysis()
                self.assertEqual(self.workbook_file.read_bytes(), original)
                self.assertEqual(list(self.workbook_file.parent.glob(".*-combined-*.xlsx")), [])


if __name__ == "__main__":
    unittest.main()
