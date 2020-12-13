# Created by Filippo Pisello
import string
from typing import List, Union, Tuple

import pandas as pd
import numpy as np

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
    dataframe : pandas dataframe object (mandatory)
        Dataframe to be considered
    keep_index : Bool, default=False
        If True, it is taken into account that the first column of the spreadsheet
        will be occupied by the index. All the dimensions will be adjusted as a
        consequence.
    skip_rows: int, default=0
        The number of rows which should be left empty at the top of the spreadsheet.
        Referring to excel row numbering, the table content starts at skip_rows + 1.
        If 0 content start at row 1.
    skip_cols: int, default=0
        The number of columns which should be left empty at the left of the spreadsheet.
        Referring to excel column labelling, the table content starts at letter with
        index skip_cols. If 0, content starts at column "A".
    correct_lists: Bool, default=False
        If True, the lists stored as the dataframe entries are modified to be more
        readable in the traditional spreadsheet softwares. More details on dedicated
        docstring.
    """

    def __init__(self, dataframe: pd.DataFrame, keep_index: bool=False, skip_rows: int=0,
                 skip_columns: int=0, correct_lists: bool=False):
        self.keep_index = keep_index
        self.df = dataframe
        self.skip_rows = skip_rows
        self.skip_cols = skip_columns

        if correct_lists:
            for column in self.df:
                self.df[column] = self.df[column].apply(self.correct_lists_for_export)

    # --------------------------------------------------------------------------
    # 1.1 - Properties
    # These capture object features which are likely to be accessed by the user.
    # --------------------------------------------------------------------------
    @property
    def indexes_depth(self):
        """
        Returns a list containing the number of multilevels for index and columns
        of the pandas dataframe used as input.

        ---------------
        Takes a pandas dataframe and returns a list of len two containing the
        number of levels of the multiindex and of the multicolumns. If the index
        is simple the associated dimension is 1. Same is true for columns.
        """
        indexes = [self.df.index, self.df.columns]
        return [len(x[0]) if isinstance(x, pd.MultiIndex) else 1 for x in indexes]

    @property
    def header_coordinates(self):
        """
        Finds the top left cell of the header and the bottom right one.

        ---------------
        Returns a list containing two lists of length two, each containing two numbers
        which univocally identify a cell. The first number refers to the position
        of the column to which the cell belongs, while the second one is the number of the row.
        Examples:
        - "A1" -> [0, 1]
        - "Z3" -> [25, 3]
        - "AA4" -> [26, 4]
        """
        starting_letter_pos = self.indexes_depth[0] * self.keep_index + self.skip_cols
        starting_number = self.skip_rows + 1
        ending_letter_pos = starting_letter_pos - 1 + self.df.shape[1]
        ending_number = starting_number - 1 + self.indexes_depth[1]
        return [[starting_letter_pos, starting_number], [ending_letter_pos, ending_number]]

    @property
    def index_coordinates(self):
        """
        Finds the top left cell of the index and the bottom right one.

        ---------------
        Returns a list containing two lists of length two, each containing two numbers
        which univocally identify a cell. The first number refers to the position
        of the column to which the cell belongs, while the second one is the number of the row.
        Examples:
        - "A1" -> [0, 1]
        - "Z3" -> [25, 3]
        - "AA4" -> [26, 4]
        """
        starting_letter_pos = self.skip_cols
        starting_number = self.indexes_depth[1] + 1 + self.skip_rows
        ending_letter_pos = starting_letter_pos - 1 + self.indexes_depth[0]
        ending_number = starting_number - 1 + self.df.shape[0]
        return [[starting_letter_pos, starting_number], [ending_letter_pos, ending_number]]

    @property
    def body_coordinates(self):
        """
        Finds the top left cell of the body and the bottom right one.

        ---------------
        The body is defined as the content of the table which belongs neither to
        the index nor to the header.

        It returns a list containing two lists of length two,each containing two numbers
        which univocally identify a cell. The first number refers to the position
        of the column to which the cell belongs, while the second one is the number of the row.
        Examples:
        - "A1" -> [0, 1]
        - "Z3" -> [25, 3]
        - "AA4" -> [26, 4]
        """
        return [[self.header_coordinates[0][0], self.index_coordinates[0][1]],
                [self.header_coordinates[1][0], self.index_coordinates[1][1]]]

    @property
    def table_coordinates(self):
        """
        Finds the top left cell of the table and the bottom right one.

        ---------------
        The table is defined as the union of header, index and body.

        It returns a list containing two lists of length two,each containing two numbers
        which univocally identify a cell. The first number refers to the position
        of the column to which the cell belongs, while the second one is the number of the row.
        Examples:
        - "A1" -> [0, 1]
        - "Z3" -> [25, 3]
        - "AA4" -> [26, 4]
        """
        return [[self.header_coordinates[0][0], self.header_coordinates[0][1]],
                [self.header_coordinates[1][0], self.index_coordinates[1][1]]]

    @property
    def header_cells(self):
        """
        Returns a list containing the names of the cells which constitute the header.

        ---------------
        The form of the output is the following ["A1", "A2", "A3"]. The cells are
        inserted by row.
        """
        return self.rectangle_of_cells(self.header_coordinates)

    @property
    def index_cells(self):
        """
        Returns a list containing the names of the cells which constitute the index.

        ---------------
        The form of the output is the following ["A1", "A2", "A3"]. The cells are
        inserted by row.
        """
        return self.rectangle_of_cells(self.index_coordinates)

    @property
    def body_cells(self):
        """
        Returns a list containing the names of the cells which constitute the body.

        ---------------
        The form of the output is the following ["A1", "A2", "A3"]. The cells are
        inserted by row.
        """
        return self.rectangle_of_cells(self.body_coordinates)

    @property
    def table_cells(self):
        """
        Returns a list containing the names of the cells which constitute the table.

        ---------------
        The form of the output is the following ["A1", "A2", "A3"]. The cells are
        inserted by row.
        """
        return self.rectangle_of_cells(self.table_coordinates)

    # --------------------------------------------------------------------------
    # 1.2 - Main methods
    # Methods which are designed and meant to be accessed by the user.
    # --------------------------------------------------------------------------
    def column(self, key: Union[str, int, List, Tuple], include_header: bool=False):
        """
        Returns the set of cells contained in the column whose key is provided

        ---------------
        The function is designed to obtain the set of cells making up a column.
        The column is identified through the information provided in the key
        argument. It can be chosen whether to heave the corresponding header cell(s)
        included or not.

        The key can either contain the column label, the column latter or its
        index. Further info is provided in the arguments description.

        When evaluating a bit of a str key, the match with the spreadsheet
        letters has priority over the match with column names. This implies
        that if "A" is passed and a column named "A" exists, then the program will
        match the key with that column, regardless of its spreadsheet letter. In
        cases where one wants to refer to columns named as spreadsheet letters,
        index or column label should be used as key.

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
        key = self.input_as_list(key, unwanted_type=dict)

        # From key to col index in int form
        int_index = []
        for element in key:
            if isinstance(element, str):
                int_index.extend(self.str_key_to_int(element))
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
            cells.extend(self.rectangle_of_cells([[col_index, top_row],
                                                  [col_index, end_row]]))
        return cells

    def row(self, key: Union[str, int, List, Tuple], include_index: bool=False):
        """
        Returns the set of cells contained in the row whose key is provided

        ---------------
        The function is designed to obtain the set of cells making up a row.
        The row is identified through the information provided in the key
        argument. It can be chosen whether to heave the corresponding index cell(s)
        included or not, if it was chosen to keep the data frame index when the
        spreadsheet object was constructed.

        The key can either contain the spreadsheet row number or its absolute
        index. Further info is provided in the arguments description. The row
        number "1" corresponds to absolute index 0.

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
        key = self.input_as_list(key, unwanted_type=dict)

        # From key to row index in int form
        int_index = []
        for element in key:
            if isinstance(element, str):
                int_index.extend(self.str_key_to_int(element, columns=False))
            elif isinstance(element, int):
                int_index.append(element + 1)

        # From row index in int for to corresponding cells
        cells = []
        # The starting and ending col do not depend on the rows chosen
        left_col = self.skip_cols + (self.indexes_depth[0] * (not include_index)) * self.keep_index
        right_col = self.body_coordinates[1][0]
        # Loop to populate the cells list
        for row_index in int_index:
            if row_index > self.body_coordinates[1][1]:
                raise KeyError("A row you are trying to access is out of index")
            cells.extend(self.rectangle_of_cells([[left_col, row_index],
                                                  [right_col, row_index]]))
        return cells

    # --------------------------------------------------------------------------
    # 2 - Worker methods
    # --------------------------------------------------------------------------
    # 2.1 - Methods used in attributes
    # --------------------------------
    @staticmethod
    def correct_lists_for_export(element):
        """
        Makes the lists contained in the table more adapt to be viewed in excel.

        -------------------------
        This function must be passed to the columns through the apply method. Lists
        are "corrected" in four ways:
        - If they contain missing values, these get removed since they would be
        exported as the string 'nan'.
        - If the list appearing as entry is empty, it is substituted by a missing
        value.
        - If the list contains only one element, the list is subsituted by that
        element.
        - If the list has multiple elements, they will appear as strings separated
        by a comma.
        """
        if isinstance(element, list):
            element = [i for i in element if i is not np.nan]
            if not element:
                element = np.nan
            elif len(element) == 1:
                element = element[0]
            else:
                element = str(element)
                for character in ["[", "]", "'"]:
                    element = element.replace(character, "")
        return element

    # --------------------------------
    # 2.2 - Chain of methods used for the row/column methods in 1.2
    # --------------------------------
    @staticmethod
    def input_as_list(input_, unwanted_type=None) -> Union[List, Tuple]:
        """
        Returns a list containing input, if input's type is not list or tuple.
        Otherwise returns input.

        ------------------
        The function purpose is to return the user input enclosed in a list so
        that one can loop through it. It is useful to handle later in a single
        operation the cases where the user is allowed to provide one or more
        items.

        The function allows to raise an error in case the input matches one or
        more types. In case multiple types want to be addressed, then a tuple
        should be provided ex: (dict, float).
        """
        # Raising error in case of unwanted type
        if unwanted_type is not None:
            if isinstance(input_, unwanted_type):
                raise TypeError(f"The type of the this input {input_} is not accepted")
        # Converting non list/tuple to list
        if not isinstance(input_, (list, tuple)):
            return [input_]
        return input_

    def str_key_to_int(self, str_key: str, columns=True) -> List[int]:
        """
        Returns a list of integers given a key of type str to individuate column(s)
        or row(s).

        ------------------
        The function intakes a str which is meant to individuate one or more
        spreadsheet columns/rows. If the key should identify columns,
        then the columns parameter should be left on True. For rows False.

        In case multiple items are passed, their references should be combined
        either with commas or a colons. The comma should be used to separate
        distinct items as it happens in a list. A colon implies a range inclusive
        on both sides.

        The columns can be identified either through their label ("Foo") or
        through their spreadsheet letter ("A").

        The rows can be identified through the spreadsheet row number ("1").

        Examples
        -------
        Consider the 2x3 dataframe {"Foo": [1, 2], "Bar": [3, 4], "Fez": [5, 6]}.
        This is what the following keys return given columns=True:
        - "A" -> [0]
        - "Foo" -> [0]
        - "Foo, Bar" -> [0, 1]
        - "Foo, B" -> [0, 1]
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
                    raise ValueError("The column key cannot have two colons without a comma in between")

                output.extend(range(self.str_index_to_int(bit[0], columns),
                                    self.str_index_to_int(bit[1], columns) + 1))
            else:
                output.append(self.str_index_to_int(bit, columns))
        return output

    def str_index_to_int(self, single_item_str: str, columns=True) -> int:
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
        The str represents the spreadsheet row number. It should then just be
        transformed to int.

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
                return self.index_from_letter(single_item_str)

        # Case for index to be interpreted as row
        else:
            try:
                return int(single_item_str)
            except ValueError as e:
                msg = "The row key provided is not convertible to int: "
                raise ValueError(msg + single_item_str) from e

    # --------------------------------
    # 2.3 - Methods used as building blocks in multiple parts of the class
    # --------------------------------
    def rectangle_of_cells(self, coordinates_list: List[List]):
        """
        Returns the cells belonging to a rectangular portion of a spreadsheet.

        ---------------
        Returns a list containing pairs in the form "A1". These correspond to the
        cells contained in a rectangular portion of spreadsheet delimited by a top
        left corner cell and a bottom right corner cell.

        The coordinates provided must be in the form
        [[TL_letter_position, TL_row],[BR_letter_position, BR_row]]
        where TL stands for top left and BR for bottom right.

        Example:
        - If [[0,1],[1,2]] is provided, the output will be [A1, A2, B1, B2]
        """
        starting_letter_pos, starting_number = coordinates_list[0]
        ending_letter_pos, ending_number = coordinates_list[1]

        output_list = []
        increasing_number = starting_number
        while starting_letter_pos <= ending_letter_pos:
            while increasing_number <= ending_number:
                output_list.append(self.letter_from_index(starting_letter_pos) + str(increasing_number))
                increasing_number = increasing_number + 1
            increasing_number = starting_number
            starting_letter_pos = starting_letter_pos + 1
        return output_list

    @staticmethod
    def letter_from_index(letter_position: int) -> str:
        """
        Returns the spreadsheet column's letter given index; ex: 0 -> "A", 26 -> "AA"

        ------------------
        The index should be interpreted as the place where the column is
        counting from left to right. The count starts from 0, which corresponds
        to "A", 1 to "B" and so on. The current program handles the indexes up to
        18 277, corresponding to column "ZZZ".
        """
        units_letter = string.ascii_uppercase[letter_position % 26]
        if letter_position <= 25:
            return units_letter

        hundreds_letter = string.ascii_uppercase[((letter_position - 26) % (26 ** 2)) // 26]
        if letter_position <= (26**2 + 25):
            return hundreds_letter + units_letter

        thousands_letter = string.ascii_uppercase[(letter_position - 26**2 - 26) // 26**2]
        if letter_position <= (26**3 + 26**2 + 25):
            return thousands_letter + hundreds_letter + units_letter

        raise ValueError("The program does not handle indexes past 18 277 yet")

    @staticmethod
    def index_from_letter(spreadsheet_letter: str) -> int:
        """
        Returns the col index given spreadsheet letter; ex: "A" -> 0, "AA" -> 26

        ------------------
        The spreadsheet letter is the column label used in spreadsheets to identify
        the column. The first one is "A" which corresponds to 0. The current program
        handles the columns up to "ZZZ", corresponding to number 18 277.
        """
        units_letter = string.ascii_uppercase.index(spreadsheet_letter[-1])
        if len(spreadsheet_letter) == 1:
            return units_letter

        hundreds_letter = string.ascii_uppercase.index(spreadsheet_letter[-2])
        if len(spreadsheet_letter) == 2:
            return 26 * hundreds_letter + units_letter

        thousands_letter = string.ascii_uppercase.index(spreadsheet_letter[-3])
        if len(spreadsheet_letter) == 3:
            return 26 ** 2 * thousands_letter + 26 * hundreds_letter + units_letter

        raise ValueError("The program does not handle column letters past 'ZZZ' yet")

    # -------------------------------------------------------------------------
    # D - Methods useful for debugging
    # -------------------------------------------------------------------------
    def print_dimensions(self):
        print("Header coordinates:", self.header_coordinates)
        print("Header cells:", self.header_cells)
        print()
        print("Index coordinates:", self.index_coordinates)
        print("Index cells:", self.index_cells)
        print()
        print("Body coordinates:", self.body_coordinates)
        print("Body cells:", self.body_cells)
        print()
        print("Table coordinates:", self.table_coordinates)
        print("Table cells:", self.table_cells)
        print()
        print("Indexes depth:", self.indexes_depth)
