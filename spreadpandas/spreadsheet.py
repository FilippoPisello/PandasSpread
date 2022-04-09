"""Main class to represent a Spreadsheet object"""
from __future__ import annotations

import pandas as pd

from .spreadsheet_element import SpreadsheetElement


class SpreadMap:
    """
    Class to represent a pandas dataframe to be loaded into a spreadsheet.

    The class intakes as main argument a pandas dataframe. Based on the dimensions,
    it finds a correspondence between the dataframe and its representation on a
    spreadsheet. The attributes are then the cells that the various parts of
    the dataframe would occupy on a spreadsheet.

    The spreadsheet elements are the following.
    - Header: cells which contain the column names. Can be of depth greater than
      one in case of multicolumns.
    - Index: cells which contain the row names. Can be of depth greater than
      one in case of multiindex.
    - Body: cells which contain the actual data.
    - Table: the union of all the cells above.

    If no rows or columns are skipped, it is assumed that the content is loaded
    into the spreadsheet starting from the top left cell, A1.
    """

    def __init__(
        self,
        dataframe: pd.DataFrame,
        keep_index: bool = False,
        keep_header: bool = True,
        skip_rows: int = 0,
        skip_columns: int = 0,
    ):
        """Create a SpreadMap object to get the dimensionality of a pandas
        data frame when exported to a spreadsheet.

        Parameters
        ----------
        dataframe : pd.DataFrame
            Pandas data frame to be represented as spreadsheet.
        keep_index : bool, optional
            If True, index is considered as an element to be mapped into the
            spreadsheet. If False, it is ignored, by default False.
        keep_header : bool, optional
            If True, header is considered as an element to be mapped into the
            spreadsheet. If False, it is ignored, by default True.
        skip_rows : int, optional
            The number of rows to skip - left empty - at the top of the
            spreadsheet. Data will be placed starting from row skip_rows + 1.
            By default 0.
        skip_columns : int, optional
            The number of columns to skip - left empty- at the left of the
            spreadsheet. Data will be placed starting from column
            skip_columns + 1. By default 0.
        """
        self.keep_index = keep_index
        self.keep_header = keep_header
        self.df = dataframe
        self.skip_rows = skip_rows
        self.skip_cols = skip_columns

    # --------------------------------
    # Data frame derived properties
    # --------------------------------
    @property
    def depth_index(self):
        return self.df.index.nlevels * self.keep_index

    @property
    def depth_columns(self):
        return self.df.columns.nlevels * self.keep_header

    # --------------------------------
    # SpreadsheetElements
    # --------------------------------
    @property
    def body(self) -> SpreadsheetElement:
        start_col = self.depth_index + self.skip_cols
        start_row = self.depth_columns + self.skip_rows
        end_col = start_col - 1 + self.df.shape[1]
        end_row = start_row - 1 + self.df.shape[0]
        return SpreadsheetElement(((start_col, start_row), (end_col, end_row)))

    @property
    def header(self) -> SpreadsheetElement:
        if not self.keep_header:
            return None

        start_row = self.skip_rows
        end_row = start_row + self.depth_columns - 1
        return SpreadsheetElement(
            (
                (self.body.coordinates[0][0], start_row),
                (self.body.coordinates[1][0], end_row),
            )
        )

    @property
    def index(self) -> SpreadsheetElement | None:
        if not self.keep_index:
            return None

        start_col = self.skip_cols
        end_col = start_col + self.depth_index - 1
        return SpreadsheetElement(
            (
                (start_col, self.body.coordinates[0][1]),
                (end_col, self.body.coordinates[1][1]),
            )
        )

    @property
    def table(self) -> SpreadsheetElement:
        if self.keep_index:
            start_col = self.index.coordinates[0][0]
        else:
            start_col = self.body.coordinates[0][0]

        if self.keep_header:
            start_row = self.header.coordinates[0][1]
        else:
            start_row = self.body.coordinates[0][1]

        return SpreadsheetElement(
            (
                (start_col, start_row),
                (self.body.coordinates[1][0], self.body.coordinates[1][1]),
            )
        )
