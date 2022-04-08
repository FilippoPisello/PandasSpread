"""Main class to represent a Spreadsheet object"""
from __future__ import annotations

import pandas as pd

from . import operations
from .custom_types import CoordinatesPair
from .spreadsheet_element import SpreadsheetElement


class Spreadsheet:
    """
    Class to represent a pandas dataframe to be loaded into a spreadsheet.

    ---------------
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

    Arguments
    ----------------
    dataframe: pandas dataframe object (mandatory)
        Dataframe to be considered
    index: Bool, default=False
        If True, it is taken into account that the first column of the spreadsheet
        will be occupied by the index. All the dimensions will be adjusted as a
        consequence.
    skip_rows: int, default=0
        The number of rows to skip - left empty - at the top of the spreadsheet.
        Data will be placed starting from row skip_rows + 1.
    skip_columns: int, default=0
        The number of columns to skip - left empty- at the left of the spreadsheet.
        Data will be placed starting from column skip_columns + 1.
    """

    def __init__(
        self,
        dataframe: pd.DataFrame,
        keep_index: bool = False,
        skip_rows: int = 0,
        skip_columns: int = 0,
    ):
        self.keep_index = keep_index
        self.df = dataframe
        self.skip_rows = skip_rows
        self.skip_cols = skip_columns

    # --------------------------------------------------------------------------
    # 1.1 - Properties
    # --------------------------------------------------------------------------
    @property
    def indexes_depth(self) -> tuple[int, int]:
        """
        Return a tuple of length two, made of the number of multilevels for
        index and columns of the input data frame.

        ---------------
        Takes a pandas dataframe and returns a list of len two containing the
        number of levels of the multiindex and of the multicolumns. If the index
        is simple the associated dimension is 1. Same is true for columns.
        """
        indexes = [self.df.index, self.df.columns]
        return tuple(len(x[0]) if isinstance(x, pd.MultiIndex) else 1 for x in indexes)

    # --------------------------------
    # 1.1.1 - Coordinates Properties
    # The following are coordinates of the four parts of a spreadsheet.
    # Coordinates are in this form: [[int, int], [int, int]].
    # The two pairs of int univocally identify a cell: they are respectively the
    # row index and the column number. The two cells identified are the top left
    # and the bottom right one.
    # --------------------------------
    @property
    def header_coordinates(self) -> CoordinatesPair:
        starting_letter_pos = self.indexes_depth[0] * self.keep_index + self.skip_cols
        starting_number = self.skip_rows + 1
        ending_letter_pos = starting_letter_pos - 1 + self.df.shape[1]
        ending_number = starting_number - 1 + self.indexes_depth[1]
        return [
            (starting_letter_pos, starting_number),
            (ending_letter_pos, ending_number),
        ]

    @property
    def index_coordinates(self) -> CoordinatesPair | None:
        if not self.keep_index:
            return None
        starting_letter_pos = self.skip_cols
        starting_number = self.indexes_depth[1] + 1 + self.skip_rows
        ending_letter_pos = starting_letter_pos - 1 + self.indexes_depth[0]
        ending_number = starting_number - 1 + self.df.shape[0]
        return [
            (starting_letter_pos, starting_number),
            (ending_letter_pos, ending_number),
        ]

    @property
    def body_coordinates(self) -> CoordinatesPair:
        return [
            (self.header_coordinates[0][0], self.first_column_coordinates[0][1]),
            (self.header_coordinates[1][0], self.first_column_coordinates[1][1]),
        ]

    @property
    def table_coordinates(self) -> CoordinatesPair:
        return [
            (self.index_coordinates[0][0], self.header_coordinates[0][1]),
            (self.header_coordinates[1][0], self.index_coordinates[1][1]),
        ]

    @property
    def first_column_coordinates(self) -> CoordinatesPair:
        column_letter = self.skip_cols + self.indexes_depth[0] * self.keep_index
        starting_number = self.indexes_depth[1] + 1 + self.skip_rows
        ending_number = starting_number - 1 + self.df.shape[0]
        return [
            (column_letter, starting_number),
            (column_letter, ending_number),
        ]

    # --------------------------------
    # 1.1.2 - Spreadsheet Elements Objects Properties
    # The following are objects which describe the four parts of a spreadsheet
    # and allow to access their properties. Look at SpreadsheetElement doc for more.
    # --------------------------------
    @property
    def header(self) -> SpreadsheetElement:
        return SpreadsheetElement(self.header_coordinates)

    @property
    def index(self) -> SpreadsheetElement | None:
        if not self.keep_index:
            return None
        return SpreadsheetElement(self.index_coordinates)

    @property
    def body(self) -> SpreadsheetElement:
        return SpreadsheetElement(self.body_coordinates)

    @property
    def table(self) -> SpreadsheetElement:
        return SpreadsheetElement(self.table_coordinates)

    @property
    def first_column(self) -> SpreadsheetElement:
        return SpreadsheetElement(self.first_column_coordinates)

    # --------------------------------------------------------------------------
    # 1.2 - Main methods
    # Methods which are designed and meant to be accessed by the user.
    # --------------------------------------------------------------------------
    def column(self, key: str | int | list | tuple, include_header: bool = False):
        """
        Returns the set of cells contained in the column whose key is provided

        ---------------
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

        ------------------
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
        else:
            try:
                return int(single_item_str)
            except ValueError as e:
                msg = "The row key provided is not convertible to int: "
                raise ValueError(msg + single_item_str) from e
