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

    def test_header(self):
        """Test that header cells are correctly identified"""
        self.assertEqual(self.t1.header.cells, ("A1", "B1", "C1"))
        self.assertEqual(self.t1_index.header.cells, ("B1", "C1", "D1"))
        self.assertEqual(self.t1_skiprow.header.cells, ("A2", "B2", "C2"))
        self.assertEqual(self.t1_skipcol.header.cells, ("B1", "C1", "D1"))

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

    def test_index(self):
        """Test that index is correctly identified"""
        self.assertIsNone(self.t1.index)
        self.assertEqual(
            self.t1_index.index.cells,
            ("A2", "A3", "A4"),
        )

    def test_first_column(self):
        """Test that first column is correctly identified"""
        self.assertEqual(self.t1.first_column.cells, ("A2", "A3", "A4"))
        self.assertEqual(self.t1_index.first_column.cells, ("B2", "B3", "B4"))
        self.assertEqual(self.t1_skiprow.first_column.cells, ("A3", "A4", "A5"))
        self.assertEqual(self.t1_skipcol.first_column.cells, ("B2", "B3", "B4"))
