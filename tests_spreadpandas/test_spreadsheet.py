"""Test the correct functioning of the spreadsheet class"""
from __future__ import annotations

from unittest import TestCase

import pandas as pd
from spreadpandas import Spreadsheet


class TestSpreadsheet(TestCase):
    def setUp(self) -> None:
        df1 = pd.DataFrame(
            {
                "A": [0, 1, 2],
                "B": [0, 1, 2],
                "C": [0, 1, 2],
            }
        )
        self.t1 = Spreadsheet(df1)
        self.t1_index = Spreadsheet(df1, keep_index=True)
        self.t1_skiprow = Spreadsheet(df1, skip_rows=1)
        self.t1_skipcol = Spreadsheet(df1, skip_columns=1)
        self.t1_noheader = Spreadsheet(df1, keep_header=False)

    def test_header(self):
        """Test that header cells are correctly identified"""
        self.assertEqual(self.t1.header.cells, ("A1", "B1", "C1"))
        self.assertEqual(self.t1_index.header.cells, ("B1", "C1", "D1"))
        self.assertEqual(self.t1_skiprow.header.cells, ("A2", "B2", "C2"))
        self.assertEqual(self.t1_skipcol.header.cells, ("B1", "C1", "D1"))
        self.assertIsNone(self.t1_noheader.header)

    def test_body(self):
        """Test that body cells are correctly identified"""
        self.assertEqual(
            self.t1.body.cells, ("A2", "B2", "C2", "A3", "B3", "C3", "A4", "B4", "C4")
        )
        self.assertEqual(
            self.t1_index.body.cells,
            ("B2", "C2", "D2", "B3", "C3", "D3", "B4", "C4", "D4"),
        )
        self.assertEqual(
            self.t1_skiprow.body.cells,
            ("A3", "B3", "C3", "A4", "B4", "C4", "A5", "B5", "C5"),
        )
        self.assertEqual(
            self.t1_skipcol.body.cells,
            ("B2", "C2", "D2", "B3", "C3", "D3", "B4", "C4", "D4"),
        )
        self.assertEqual(
            self.t1_noheader.body.cells,
            ("A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"),
        )

    def test_index(self):
        """Test that index is correctly identified"""
        self.assertIsNone(self.t1.index)
        self.assertEqual(
            self.t1_index.index.cells,
            ("A2", "A3", "A4"),
        )

    def test_table(self):
        """Test that table is correctly identified"""
        self.assertEqual(self.t1.table.cells_range, "A1:C4")
        self.assertEqual(self.t1_index.table.cells_range, "A1:D4")
        self.assertEqual(self.t1_skiprow.table.cells_range, "A2:C5")
        self.assertEqual(self.t1_skipcol.table.cells_range, "B1:D4")
        self.assertEqual(self.t1_noheader.table.cells_range, "A1:C3")
