"""Finance matching by TYPE and Mode (decision 2026-10-07), using cases from the live data."""

import unittest

from loa_engine import FinanceLookupEngine, finance_mode_class


def row(type_, area, tcp, ded=None, title=''):
    return {'TYPE': type_, 'Area Index': area, 'Title Mod': title,
            'Total Contract Price': str(tcp), 'Total Payment Deductible to TCP': str(tcp if ded is None else ded)}


class TestModeClass(unittest.TestCase):
    def test_modes_in_land_master(self):
        for mode in ('SECURED SALE', 'SS - STP', 'SS-STP', 'SS - FROM ROW', 'SS - FROM LAND STP',
                     'SS - RESCINDED', 'SOLD TO PARTNER', 'SECURED SALE - NON SP'):
            self.assertEqual(finance_mode_class(mode), 'SALE', mode)
        for mode in ('SECURED LEASE', 'SL - STP', 'SL - FROM LAND STP', 'SECURED LEASE - NON SP'):
            self.assertEqual(finance_mode_class(mode), 'LEASE', mode)
        for mode in ('RESCINDED', 'TO BE IDENTIFIED', '', None):
            self.assertEqual(finance_mode_class(mode), '', mode)


class TestTypeAwareMatching(unittest.TestCase):
    def test_pen_299_sale_row_wins_over_cwa(self):
        engine = FinanceLookupEngine([row('SALE', 'PEN-299', 6281628), row('CWA', 'PEN-299', 39396)])
        res = engine.lookup('LAND', '', 'PEN-299', '', 'SECURED SALE')
        self.assertTrue(res['matched'])
        self.assertEqual(res['total_contract_price'], 6281628.0)

    def test_srs_793_row_account_ignores_cwa_rows(self):
        engine = FinanceLookupEngine([
            row('ROW', 'SRS-793', 1000000), row('CWA', 'SRS-793', 34502), row('CWA', 'SRS-793', 74028),
            row('CWA', 'SRS-793', 4820), row('CWA', 'SRS-793', 34502),
        ])
        res = engine.lookup('ROW', '', 'SRS-793', '', 'SECURED SALE')
        self.assertTrue(res['matched'])
        self.assertEqual(res['match_type'], 'ROW (Area Index)')
        self.assertEqual(res['total_contract_price'], 1000000.0)

    def test_lease_mode_takes_lease_row(self):
        engine = FinanceLookupEngine([row('SALE', 'A-1', 500), row('LEASE', 'A-1', 90)])
        self.assertEqual(engine.lookup('LAND', '', 'A-1', '', 'SL - STP')['total_contract_price'], 90.0)
        self.assertEqual(engine.lookup('LAND', '', 'A-1', '', 'SS - STP')['total_contract_price'], 500.0)

    def test_unclear_mode_with_sale_and_lease_is_ambiguous(self):
        engine = FinanceLookupEngine([row('SALE', 'A-1', 500), row('LEASE', 'A-1', 90)])
        res = engine.lookup('LAND', '', 'A-1', '', 'RESCINDED')
        self.assertFalse(res['matched'])
        self.assertTrue(res['ambiguous'])

    def test_same_type_conflict_is_a_duplicate(self):
        engine = FinanceLookupEngine([row('SALE', 'A-1', 500), row('SALE', 'A-1', 700), row('LEASE', 'A-1', 90)])
        res = engine.lookup('LAND', '', 'A-1', '', 'SECURED SALE')
        self.assertFalse(res['matched'])
        self.assertTrue(res['ambiguous'])
        self.assertEqual(res['ambiguous_count'], 2)
        # A lease account on the same Area-Index is unaffected by the SALE duplicates.
        self.assertTrue(engine.lookup('LAND', '', 'A-1', '', 'SECURED LEASE')['matched'])

    def test_different_types_are_not_duplicates(self):
        engine = FinanceLookupEngine([row('SALE', 'A-1', 500), row('CWA', 'A-1', 10), row('CWA', 'A-1', 20)])
        self.assertTrue(engine.lookup('LAND', '', 'A-1', '', 'SECURED SALE')['matched'])

    def test_sa_supplies_only_land_special_and_cwa_never_supplies_land(self):
        # Decision 2026-10-12: SA = Special Acquisition. A LAND - SPECIAL account with
        # no SALE/LEASE row uses its SA row; plain LAND accounts never do; CWA never.
        engine = FinanceLookupEngine([row('SA', 'TAA-1', 14000000), row('CWA', 'X-1', 100)])
        special = engine.lookup('LAND - SPECIAL', '', 'TAA-1', '', 'SECURED SALE')
        self.assertTrue(special['matched'])
        self.assertEqual(special['total_contract_price'], 14000000)
        self.assertIn('Special Acquisition', special['match_type'])
        self.assertFalse(engine.lookup('LAND', '', 'TAA-1', '', 'SECURED SALE')['matched'])
        self.assertFalse(engine.lookup('LAND', '', 'X-1', '', 'SECURED SALE')['matched'])
        self.assertFalse(engine.lookup('LAND - SPECIAL', '', 'X-1', '', 'SECURED SALE')['matched'])

    def test_sale_row_still_wins_over_sa_for_land_special(self):
        engine = FinanceLookupEngine([row('SALE', 'TAA-2', 500), row('SA', 'TAA-2', 900)])
        res = engine.lookup('LAND - SPECIAL', '', 'TAA-2', '', 'SECURED SALE')
        self.assertEqual(res['total_contract_price'], 500)
        self.assertNotIn('Special Acquisition', res['match_type'])

    def test_conflicting_sa_rows_are_reported(self):
        engine = FinanceLookupEngine([row('SA', 'TAA-3', 1), row('SA', 'TAA-3', 2)])
        res = engine.lookup('LAND - SPECIAL', '', 'TAA-3', '', '')
        self.assertFalse(res['matched'])
        self.assertTrue(res['ambiguous'])

    def test_row_account_never_takes_sale_row_on_shared_area(self):
        engine = FinanceLookupEngine([row('SALE', 'PEN-269', 800)])
        self.assertFalse(engine.lookup('ROW', '', 'PEN-269', '', 'SECURED SALE')['matched'])
        self.assertTrue(engine.lookup('LAND', '', 'PEN-269', '', 'SECURED SALE')['matched'])

    def test_title_fallback_respects_type(self):
        engine = FinanceLookupEngine([row('CWA', 'Z', 10, title='T-5555'), row('SALE', 'Z2', 99, title='T-5555')])
        res = engine.lookup('LAND', '', 'NO-MATCH', 'T-5555', 'SECURED SALE')
        self.assertTrue(res['matched'])
        self.assertEqual(res['total_contract_price'], 99.0)

    def test_untyped_excel_rows_still_match_land_accounts(self):
        engine = FinanceLookupEngine([{'Area Index': 'A-9', 'Total Contract Price': '100',
                                       'Total Payment Deductible to TCP': '50'}])
        self.assertTrue(engine.lookup('LAND', '', 'A-9', '', 'SECURED SALE')['matched'])


if __name__ == '__main__':
    unittest.main()
