"""
Unit & Integration Tests for Advance Settlement Reconciliation
Built using Python standard unittest.
"""
import unittest
from pathlib import Path
from src.matcher import ReconciliationMatcher
from src.config import GL_FILE_PATH, WP_FILE_PATH


class TestReconciliation(unittest.TestCase):
    def setUp(self):
        self.matcher = ReconciliationMatcher(GL_FILE_PATH, WP_FILE_PATH)
        self.matcher.load_data()
        self.results = self.matcher.match_all()
        self.metrics = self.matcher.get_summary_metrics()

    def test_data_loading(self):
        self.assertEqual(len(self.matcher.gl_kredit_df), 39, "GL should contain 39 credit transactions")
        self.assertEqual(len(self.matcher.wp_rows), 15, "Working Paper should contain 15 advance items")

    def test_matching_metrics(self):
        self.assertEqual(self.metrics['total_items'], 15)
        self.assertEqual(self.metrics['count_settled'], 14)
        self.assertEqual(self.metrics['count_partial'], 1)
        self.assertEqual(self.metrics['count_unsettled'], 0)
        self.assertEqual(self.metrics['total_advance'], 570406879.0)
        self.assertEqual(self.metrics['total_realization'], 463302663.0)
        self.assertEqual(self.metrics['total_saldo'], 107104216.0)
        self.assertEqual(self.metrics['settlement_rate_pct'], 81.22)

    def test_styling_multi_voucher_accuracy(self):
        # Styling 2BR
        styling_2br = next(r for r in self.results if 'SM 1132' in r['desc'] and 'STYLING' in r['desc'])
        self.assertEqual(styling_2br['vouchers_count'], 8)
        self.assertEqual(styling_2br['realization_amount'], 25695900.0)
        self.assertEqual(styling_2br['saldo'], 0.0)

        # Styling Studio
        styling_studio = next(r for r in self.results if 'SM 1127' in r['desc'] and 'STYLING' in r['desc'])
        self.assertEqual(styling_studio['vouchers_count'], 3)
        self.assertEqual(styling_studio['realization_amount'], 1583700.0)
        self.assertEqual(styling_studio['saldo'], 0.0)

    def test_pbb_partial_settlement(self):
        pbb = next(r for r in self.results if 'PBB' in r['desc'])
        self.assertEqual(pbb['amount'], 422974179.0)
        self.assertEqual(pbb['realization_amount'], 315869963.0)
        self.assertEqual(pbb['saldo'], 107104216.0)
        self.assertEqual(pbb['status'], 'PARTIAL')


if __name__ == '__main__':
    unittest.main()
