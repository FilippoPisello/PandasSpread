"""Main class to represent a Spreadsheet object"""
from __future__ import annotations

import pandas as pd

from .spreadsheet_element import SpreadsheetElement


class Spreadsheet:
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
        """Create a Spreadsheet object to get the dimensionality of a pandas
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

    def column(
        self,
        key: str | int | list | tuple,
        include_header: bool = False,
    ):
        """
        Return the set of cells contained in the column whose key is provided.

        The function is designed to obtain the set of cells making up a column.
        The column is identified through the the key argument. It can be chosen
        whether to heave the corresponding header cell(s) included or not.
        The key can either contain the column label, the column latter or its
        index. Further info is provided in the arguments description.
        When evaluating a bit of a str key, the match with the spreadsheet
        letters has priority over the match with column names. This implies
        that if "A" is passed and a column named "A" exists, then the program will
        match the key with that column, regardless of its spreadsheet letter. In
        a similar case index or column label should be used as key.

        Arguments
        ----------------
        key: str/int/List/Tuple (mandatory)
            The parameter used to identify the column. It is extremely flexible.
            - Single column index can be passed as int (ex: 1)
            - Multiple column index can be passed as list/tuple of int (ex: [1, 2])
            - Single spreadsheet column letter can be passed as str (ex: "A")
            - Multiple spreadsheet column letters can be passed as list/tuple of
            str (ex: ["A", "B"]). Alternatively, they can be passed as str
            being comma separated "A, B".
            - A range of spreadsheet columns can be passed as str with the first
            and last letter divided by a colon ("A:C"). Inclusive on both sides.
            - The last three options apply in the same way to column labels.
            (ex: "Foo")(ex: ["Foo", "Bar"])(ex: "Foo, Bar")(ex: "Foo:Bar")
        include_header: Bool, default=False
            It specifies whether the cells making up the column's header should
            be included or not. If False, it only returns the cells containing
            data for that column.
        """
        key = self._input_as_list(key, unwanted_type=dict)

        # From key to col index in int form
        int_index = []
        for element in key:
            if isinstance(element, str):
                int_index.extend(self._str_key_to_int(element))
            elif isinstance(element, int):
                int_index.append(element)

        # From col index in int form to corresponding cells
        cells = []
        # The starting and ending row do not depend on the columns chosen
        top_row = self.indexes_depth[1] * (not include_header) + 1 + self.skip_rows
        end_row = self.body_coordinates[1][1]
        # Loop to populate the cells list
        for col_index in int_index:
            if col_index > self.body_coordinates[1][0]:
                raise KeyError("A column you are trying to access is out of index")

            coordinates_pair = ([col_index, top_row], [col_index, end_row])
            cells.extend(operations.cells_rectangle(coordinates_pair))
        return cells

    def row(self, key: str | int | list | tuple, include_index: bool = False):
        """
        Returns the set of cells contained in the row whose key is provided
        ---------------
        The function is designed to obtain the set of cells making up a row.
        The row is identified through the key argument. It can be chosen whether
        to heave the corresponding index cell(s) included or not, if it was chosen
        to keep the data frame index when the spreadsheet object was constructed.
        The key can either contain the spreadsheet row number or its absolute
        index. The row number "1" corresponds to absolute index 0. Further info
        is provided in the arguments description.
        Arguments
        ----------------
        key: str/int/List/Tuple (mandatory)
            The parameter used to identify the column. It is extremely flexible.
            - Single column index can be passed as int (ex: 1)
            - Multiple column index can be passed as list/tuple of int (ex: [1, 2])
            - Single spreadsheet row index can be passed as str (ex: "1")
            - Multiple spreadsheet row index can be passed as list/tuple of
            str (ex: ["1", "2"]). Alternatively, they can be passed as str
            being comma separated "1, 2".
            - A range of spreadsheet rows can be passed as str with the first
            and last index divided by a colon ("1:3"). Inclusive on both sides.
        include_index: Bool, default=False
            It specifies whether the cells making up the row's index should
            be included or not. If False, it only returns the cells containing
            data for that row.
        """
        key = self._input_as_list(key, unwanted_type=dict)

        # From key to row index in int form
        int_index = []
        for element in key:
            if isinstance(element, str):
                int_index.extend(self._str_key_to_int(element, columns=False))
            elif isinstance(element, int):
                int_index.append(element + 1)

        # From row index in int for to corresponding cells
        cells = []
        # The starting and ending col do not depend on the rows chosen
        left_col = (
            self.skip_cols
            + (self.indexes_depth[0] * (not include_index)) * self.keep_index
        )
        right_col = self.body_coordinates[1][0]
        # Loop to populate the cells list
        for row_index in int_index:
            if row_index > self.body_coordinates[1][1]:
                raise KeyError("A row you are trying to access is out of index")

            coordinates_pair = ([left_col, row_index], [right_col, row_index])
            cells.extend(operations.cells_rectangle(coordinates_pair))
        return cells

    # --------------------------------------------------------------------------
    # 2 - Worker methods
    # --------------------------------------------------------------------------
    # 2.1 - Chain of methods used for the row/column methods in 1.2
    # --------------------------------
    @staticmethod
    def _input_as_list(input_, unwanted_type=None) -> list | tuple:
        """
        Returns a list containing input, if input's type is not list or tuple.
        Otherwise returns input.
        ------------------
        The function returns the input element enclosed in a list.
        It allows to raise an error in case the input matches one or more types.
        In case multiple types want to be excluded, then a tuple should be
        provided ex: (dict, float).
        """
        # Raising error in case of unwanted type
        if unwanted_type is not None:
            if isinstance(input_, unwanted_type):
                raise TypeError(f"The type of the this input {input_} is not accepted")
        # Converting non list/tuple to list
        if not isinstance(input_, (list, tuple)):
            return [input_]
        return input_

    def _str_key_to_int(self, str_key: str, columns=True) -> list[int]:
        """
        Returns a list of integers given a key of type str to individuate column(s)
        or row(s).

        Intakes a str which should individuate one or more spreadsheet columns/rows.
        For columns, the parameter columns should be True. For rows False.
        Multiple items should be chained either with commas, list of distinct
        elements, or colons, range inclusive on both sides.
        The columns can be identified either through their label ("Foo") or
        through their spreadsheet letter ("A"). The rows can be identified through
        the spreadsheet row number ("1").

        Examples
        -------
        Consider the 2x3 dataframe {"Foo": [1, 2], "Bar": [3, 4], "Fez": [5, 6]}.
        This is what the following keys return given columns=True:
        - "A" -> [0]
        - "Foo" -> [0]
        - "Foo, Bar" -> [0, 1]
        - "Foo:Fez" -> [0, 1, 2]
        - "Foo:C" -> [0, 1, 2]
        This is what the following keys return given columns=False:
        - "1" -> [0]
        - "1,3" -> [0, 2]
        - "1:3" -> [0, 1, 2]
        """
        # First the commas are handled. The str is dived in subelements.
        str_key = str_key.split(",")
        output = []
        for bit in str_key:
            # Colons are now addressed
            if bit.find(":") > -1:
                bit = bit.split(":")

                # Inputs of type "X:X:X" are not accepted
                if len(bit) > 2:
                    raise ValueError(
                        "The column key cannot have two colons without a comma in between"
                    )

                output.extend(
                    range(
                        self._str_index_to_int(bit[0], columns),
                        self._str_index_to_int(bit[1], columns) + 1,
                    )
                )
            else:
                output.append(self._str_index_to_int(bit, columns))
        return output

    def _str_index_to_int(self, single_item_str: str, columns=True) -> int:
        """
        Returns an int given a str identifier for a spreadsheet column/row.
        ------------------
        The function intakes a str which is meant to identify a single
        spreadsheet column/row and returns the corresponding int index.
        Column=True
        ----------
        The str can either be (1) the column label of the pandas dataframe or (2)
        the spreadsheet column letter. The script prioritizes (1). Consider the
        second set of examples to understand the implication of this.
        Example given the 2x2 dataframe {"Foo" : [1, 2], "Bar" : [3, 4]}:
        - "Foo" -> 0
        - "Bar" -> 1
        - "A" -> 0
        - "B" -> 1
        Example given the 2x2 dataframe {"Foo" : [1, 2], "A" : [3, 4]}:
        - "A" -> 1
        - "B" -> 1
        Column=False
        ----------
        The str represents the spreadsheet row number which is just transformed
        to int.
        Example:
        - "1" -> 1
        """
        # Remove any unwanted spaces
        single_item_str = single_item_str.strip()

        # Case for index to be interpreted as column
        if columns:
            # Tries first to see if the string appears in the column labels
            try:
                return self.df.columns.to_list().index(single_item_str)
            # If fails it runs the conversion using col str
            except ValueError:
                return operations.index_from_letter(single_item_str)

        # Case for index to be interpreted as row
        try:
            return int(single_item_str)
        except ValueError as e:
            msg = "The row key provided is not convertible to int: "
            raise ValueError(msg + single_item_str) from e
