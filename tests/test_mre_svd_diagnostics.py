"""Numerical guarantees for signed MRE SVD and its diagnostic comparisons."""

import unittest

import numpy as np
from openpyxl import Workbook

import component_analysis_mre as analysis


class MreSvdDiagnosticsTests(unittest.TestCase):
    def sheet(self, spectra, *, wavelengths=None):
        spectra = np.asarray(spectra)
        if wavelengths is None:
            wavelengths = np.arange(190, 190 + spectra.shape[1], dtype=float)
        workbook = Workbook()
        self.addCleanup(workbook.close)
        source = workbook.active
        source.title = "MRE"
        source.append(["wavelength_nm", *[f"sample {i + 1}" for i in range(len(spectra))]])
        for wavelength, values in zip(wavelengths, spectra.T):
            source.append([float(wavelength), *[
                None if value is None else float(value) for value in values
            ]])
        return source

    def dataset(self, *, spectra_count=9, wavelength_count=61):
        wavelengths = np.linspace(190, 250, wavelength_count)
        patterns = np.vstack([
            -10000 * np.exp(-((wavelengths - 207) / 9) ** 2),
            8000 * np.exp(-((wavelengths - 235) / 8) ** 2),
            2000 * np.sin((wavelengths - 190) / 15),
        ])
        rng = np.random.default_rng(801)
        spectra = rng.uniform(0.1, 1, (spectra_count, len(patterns))) @ patterns
        spectra += rng.normal(0, 0.5, spectra.shape)
        return wavelengths, spectra

    def test_both_decompositions_preserve_orientation_units_and_full_reconstruction(self):
        wavelengths, spectra = self.dataset()
        result = analysis.analyze_mre(self.sheet(spectra, wavelengths=wavelengths), reconstruction_ranks=[1])
        self.assertIs(result.svd, result.uncentered.stats)
        np.testing.assert_allclose(result.spectra, spectra)
        for mode, expected_count, expected_shape in (
            (result.uncentered, 9, (9, 61)),
            (result.centered, 8, (8, 61)),
        ):
            with self.subTest(effective_shape=expected_shape):
                self.assertEqual(mode.effective_shape, expected_shape)
                self.assertEqual(mode.scores.shape, (9, expected_count))
                self.assertEqual(mode.patterns.shape, (expected_count, 61))
                self.assertEqual(mode.offset.shape, (61,))
                np.testing.assert_allclose(
                    mode.patterns @ mode.patterns.T, np.eye(expected_count), atol=2e-14,
                )
                np.testing.assert_allclose(
                    mode.scores @ mode.patterns + mode.offset, spectra, atol=2e-10,
                )
        np.testing.assert_allclose(result.uncentered.offset, 0, atol=0)
        np.testing.assert_allclose(result.centered.offset, spectra.mean(axis=0))
        np.testing.assert_allclose(result.centered.scores.mean(axis=0), 0, atol=2e-11)
        self.assertEqual(result.diagnostic_components, 5)

    def test_truncated_reconstruction_error_matches_discarded_singular_values(self):
        wavelengths, spectra = self.dataset()
        result = analysis.analyze_mre(
            self.sheet(spectra, wavelengths=wavelengths), reconstruction_ranks=[1, 2, 5, 9],
        )
        self.assertEqual([item.rank for item in result.reconstructions], [1, 2, 5, 9])
        squared = result.svd.singular_values ** 2
        errors = []
        for item in result.reconstructions:
            with self.subTest(rank=item.rank):
                expected = (
                    result.uncentered.scores[:, :item.rank]
                    @ result.uncentered.patterns[:item.rank]
                )
                np.testing.assert_allclose(item.reconstructed, expected, atol=1e-10)
                np.testing.assert_allclose(item.residuals, spectra - expected, atol=1e-10)
                residual_sse = np.sum(item.residuals ** 2)
                np.testing.assert_allclose(residual_sse, squared[item.rank:].sum(), atol=1e-12)
                self.assertAlmostEqual(item.energy_fraction, squared[:item.rank].sum() / squared.sum())
                self.assertAlmostEqual(item.relative_error, np.linalg.norm(item.residuals) / np.linalg.norm(spectra))
                np.testing.assert_allclose(item.rms_by_spectrum, np.sqrt(np.mean(item.residuals ** 2, axis=1)))
                np.testing.assert_allclose(item.rms_by_wavelength, np.sqrt(np.mean(item.residuals ** 2, axis=0)))
                errors.append(item.relative_error)
        self.assertTrue(np.all(np.diff(errors) <= 0))
        self.assertLess(errors[-1], 1e-13)

    def test_constant_sum_mixture_has_one_centered_direction_and_two_uncentered(self):
        wavelengths = np.arange(190, 251, dtype=float)
        profiles = np.vstack([
            -10000 * np.exp(-((wavelengths - 208) / 10) ** 2),
            8000 * np.exp(-((wavelengths - 235) / 9) ** 2),
        ])
        fraction = np.linspace(0.05, 0.95, 8)
        spectra = np.column_stack([fraction, 1 - fraction]) @ profiles
        result = analysis.analyze_mre(self.sheet(spectra, wavelengths=wavelengths), reconstruction_ranks=[1])
        self.assertEqual(result.uncentered.stats.variance_rank, 2)
        self.assertEqual(result.centered.stats.variance_rank, 1)
        reconstructed = (
            result.centered.scores[:, :1] @ result.centered.patterns[:1]
            + result.centered.offset
        )
        np.testing.assert_allclose(reconstructed, spectra, atol=2e-10)

    def test_centered_threshold_uses_independent_sample_dimension(self):
        for spectra_count in (2, 9):
            with self.subTest(spectra_count=spectra_count):
                wavelengths, spectra = self.dataset(spectra_count=spectra_count)
                result = analysis.analyze_mre(self.sheet(spectra, wavelengths=wavelengths), reconstruction_ranks=[1])
                effective_shape = (spectra_count - 1, 61)
                self.assertEqual(result.centered.effective_shape, effective_shape)
                expected_singular_values = np.linalg.svd(
                    spectra - spectra.mean(axis=0), full_matrices=False, compute_uv=False,
                )[:spectra_count - 1]
                np.testing.assert_allclose(
                    result.centered.stats.singular_values, expected_singular_values, atol=1e-10,
                )
                beta = min(effective_shape) / max(effective_shape)
                omega = 0.56 * beta ** 3 - 0.95 * beta ** 2 + 1.82 * beta + 1.43
                self.assertAlmostEqual(
                    result.centered.stats.gavish_donoho_threshold,
                    omega * np.median(expected_singular_values), places=8,
                )

    def test_centering_keeps_full_wavelength_rank_when_samples_outnumber_wavelengths(self):
        wavelengths, spectra = self.dataset(spectra_count=12, wavelength_count=5)
        result = analysis.analyze_mre(self.sheet(spectra, wavelengths=wavelengths), reconstruction_ranks=[1])
        self.assertEqual(result.centered.effective_shape, (11, 5))
        self.assertEqual(result.centered.scores.shape, (12, 5))
        self.assertEqual(result.centered.patterns.shape, (5, 5))
        self.assertEqual(len(result.centered.stats.singular_values), 5)
        np.testing.assert_allclose(
            result.centered.scores @ result.centered.patterns + result.centered.offset,
            spectra, atol=2e-10,
        )

    def test_identical_signed_spectra_report_no_centered_variation_without_failure(self):
        spectrum = np.linspace(-16384, 8192, 61)
        spectra = np.repeat(spectrum[None, :], 8, axis=0)
        result = analysis.analyze_mre(self.sheet(spectra), reconstruction_ranks=[1])
        self.assertEqual(result.svd.variance_rank, 1)
        self.assertEqual(result.centered.stats.variance_rank, 0)
        self.assertEqual(result.centered.stats.gavish_donoho_rank, 0)
        np.testing.assert_array_equal(result.centered.stats.explained_energy_ratio, 0)
        np.testing.assert_array_equal(result.centered.stats.cumulative_energy_ratio, 0)
        np.testing.assert_allclose(
            result.centered.scores @ result.centered.patterns + result.centered.offset,
            spectra, atol=1e-10,
        )

    def test_reconstruction_ranks_are_deduplicated_and_optional(self):
        wavelengths, spectra = self.dataset()
        sheet = self.sheet(spectra, wavelengths=wavelengths)
        result = analysis.analyze_mre(sheet, reconstruction_ranks=[5, 2, 5, 1], diagnostic_components=99)
        self.assertEqual([item.rank for item in result.reconstructions], [1, 2, 5])
        self.assertEqual(result.diagnostic_components, min(spectra.shape))
        default = analysis.analyze_mre(sheet)
        self.assertEqual(default.reconstructions, ())
        self.assertIsNone(default.uncentered)
        self.assertIsNone(default.centered)
        self.assertFalse(default.diagnostics_enabled)
        self.assertEqual(default.diagnostic_components, 0)
        np.testing.assert_allclose(default.svd.singular_values, result.svd.singular_values, atol=1e-9)
        self.assertEqual(default.svd.variance_rank, result.svd.variance_rank)
        self.assertEqual(default.svd.gavish_donoho_rank, result.svd.gavish_donoho_rank)

    def test_reconstruction_rank_and_display_count_validation(self):
        wavelengths, spectra = self.dataset()
        sheet = self.sheet(spectra, wavelengths=wavelengths)
        for ranks in ([], [0], [-1], [10]):
            with self.subTest(ranks=ranks):
                with self.assertRaises(ValueError):
                    analysis.analyze_mre(sheet, reconstruction_ranks=ranks)
        for count in (0, -1):
            with self.subTest(diagnostic_components=count):
                with self.assertRaises(ValueError):
                    analysis.analyze_mre(sheet, diagnostic_components=count)

    def test_residual_correlations_skip_pairs_across_excluded_wavelength_rows(self):
        wavelengths = np.arange(190, 205, dtype=float)
        rng = np.random.default_rng(95)
        spectra = rng.normal(0, 2, (5, len(wavelengths))).astype(object)
        spectra[0, [3, 8, 12]] = None
        result = analysis.analyze_mre(
            self.sheet(spectra, wavelengths=wavelengths), reconstruction_ranks=[1],
        )
        self.assertEqual(result.wavelength_step_nm, 1)
        pair_mask = np.isclose(np.diff(result.wavelengths), 1)
        self.assertTrue(np.any(~pair_mask))
        residuals = result.reconstructions[0].residuals
        expected = [
            np.corrcoef(row[:-1][pair_mask], row[1:][pair_mask])[0, 1]
            for row in residuals
        ]
        np.testing.assert_allclose(result.reconstructions[0].lag1_correlations, expected)

    def test_irregular_source_axis_has_no_fixed_lag_correlations(self):
        rng = np.random.default_rng(18)
        wavelengths = np.array([190, 191, 193, 194, 197, 198, 202, 206], dtype=float)
        result = analysis.analyze_mre(
            self.sheet(rng.normal(size=(4, len(wavelengths))), wavelengths=wavelengths),
            reconstruction_ranks=[1],
        )
        self.assertIsNone(result.wavelength_step_nm)
        self.assertTrue(np.isnan(result.reconstructions[0].lag1_correlations).all())

    def test_too_few_adjacent_pairs_and_zero_residuals_have_undefined_correlations(self):
        wavelengths = np.arange(190, 200, dtype=float)
        rng = np.random.default_rng(24)
        spectra = rng.normal(size=(3, len(wavelengths))).astype(object)
        spectra[0, [1, 3, 5, 7, 9]] = None
        result = analysis.analyze_mre(
            self.sheet(spectra, wavelengths=wavelengths), reconstruction_ranks=[1],
        )
        self.assertEqual(result.wavelength_step_nm, 1)
        self.assertTrue(np.isnan(result.reconstructions[0].lag1_correlations).all())

        spectra = np.zeros((2, 6))
        spectra[0, 0] = spectra[1, 1] = 1
        exact = analysis.analyze_mre(self.sheet(spectra), reconstruction_ranks=[2])
        np.testing.assert_array_equal(exact.reconstructions[0].residuals, 0)
        self.assertTrue(np.isnan(exact.reconstructions[0].lag1_correlations).all())


if __name__ == "__main__":
    unittest.main()
