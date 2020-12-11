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

    def __init__(self, dataframe, keep_index=False, skip_rows=0, skip_columns=0,
                 correct_lists=False):
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
    def column(self, key: Union[str, int, List, Tuple], include_header=False):
        """
        Returns the set of cells contained in the column whose key is provided

        ---------------
        The function is designed to obtain the set of cells making up a column.
        The column is identified through the information provided in the key
        argument. It can be chosen whether to heave the corresponding header cell(s)
        included or not.

        Note that when evaluating a bit of a str key, the match with the spreadsheet
        letters is carried out if no match is found with column names. This implies
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
        if not isinstance(key, (list,tuple)):
            key = [key]
        int_index = []
        for element in key:
            if isinstance(element, str):
                int_index.extend(self.str_key_to_int(element))
            elif isinstance(element, int):
                int_index.append(element)
        cells = []
        for col_index in int_index:
            if col_index > self.header_coordinates[1][0]:
                raise KeyError("The column you are trying to access is out of index")
            top_row = self.indexes_depth[1] * (not include_header) + 1 + self.skip_rows
            end_row = self.body_coordinates[1][1]
            cells.extend(self.rectangle_of_cells([[col_index, top_row],
                                                  [col_index, end_row]]))
        return cells

    # --------------------------------------------------------------------------
    # 2 - Methods used as building blocks of tier 1 methods
    # --------------------------------------------------------------------------
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

    def str_key_to_int(self, str_key: str) -> List[int]:
        """
        Returns a list of integers given a string like column key. Ex: "A,B" -> [0, 1]

        ------------------
        The function intakes a str which is meant to individuate one or more
        spreadsheet columns. The columns can be identified either through
        their label ("Foo") or through their spreadsheet letter ("A").

        In case multiple columns are passed, their references should be combined
        either with commas or a colons. The comma should be used to separate
        distinct items as it happens in a list. A colon implies a range inclusive
        on both sides.

        The spreadsheet letters are unpacked following the mentioned rules and
        translated into indexes given the existing convertion criterion.

        General example:
        - "A" -> [0]
        - "A, B" -> [0, 1]
        - "A:C" -> [0, 1, 2]
        - "A:C, E" -> [0, 1, 2, 4]
        Example given the 2x3 df {"Foo": [1, 2], "Bar": [3, 4], "Fez": [5, 6]}:
        - "A" -> [0]
        - "Foo" -> [0]
        - "Foo, Bar" -> [0, 1]
        - "Foo, B" -> [0, 1]
        - "Foo:Fez" -> [0, 1, 2]
        - "Foo:C" -> [0, 1, 2]
        """
        # First the commas are handled. The str is dived in subelements.
        str_key = str_key.split(",")
        output = []
        for bit in str_key:
            # Colons are now addressed
            if bit.find(":") > -1:
                bit = bit.split(":")

                # Inputs of type "A:C:E" are not accepted
                if len(bit) > 2:
                    raise ValueError("The column key cannot have two colons without a comma in between")

                output.extend(range(self.str_index_to_int(bit[0].strip()),
                                    self.str_index_to_int(bit[1].strip()) + 1))
            else:
                output.append(self.str_index_to_int(bit.strip()))
        return output

    # --------------------------------------------------------------------------
    # 3 - Methods used as building blocks of tier 2 methods
    # --------------------------------------------------------------------------
    def str_index_to_int(self, single_col_str: str) -> int:
        """
        Returns an int given a str identifier for a spreadsheet column. Ex: "A" -> 0

        ------------------
        The function intakes a str which is meant to identify a single
        spreadsheet column and returns the corresponding int index. The str can
        either be (1) the column label of the pandas dataframe or (2) the column
        letter of the corresponding spreadsheet.

        Note that the script prioritizes (1). Consider the second set of examples
        to understand the implication of this.

        Example given the 2x2 dataframe {"Foo" : [1, 2], "Bar" : [3, 4]}:
        - "Foo" -> 0
        - "Bar" -> 1
        - "A" -> 0
        - "B" -> 1
        Example given the 2x2 dataframe {"Foo" : [1, 2], "A" : [3, 4]}:
        - "A" -> 1
        - "B" -> 1
        """
        # Tries first to see if the string appears in the column labels
        try:
            return self.df.columns.to_list().index(single_col_str)
        # If fails it runs the conversion using col str
        except ValueError:
            return self.index_from_letter(single_col_str)

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
