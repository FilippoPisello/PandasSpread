#!/usr/bin/env python
# Created by Filippo Pisello
import string
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
    dataframe : pandas dataframe object
        Dataframe to be considered
    keep_index : Bool
        If True, it is taken into account that the first column of the spreadsheet
        will be occupied by the index. All the dimensions will be adjusted as a 
        consequence.
    skip_rows: int
        The number of rows which should be left empty at the top of the spreadsheet.
        Referring to excel row numbering, the table content starts at skip_rows + 1.
        If 0 content start at row 1.
    skip_cols: int
        The number of columns which should be left empty at the left of the spreadsheet.
        Referring to excel column labelling, the table content starts at letter with
        index skip_cols. If 0 content start at column "A".
    correct_lists: Bool
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
    # 1 - Properties
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
    # 2 - Methods meant to be used directly or to define attributes/properties
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

    def rectangle_of_cells(self, coordinates_list):
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
    
    # --------------------------------------------------------------------------
    # 3 - Methods used at their times in methods of tier 2
    # --------------------------------------------------------------------------
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
        
        raise ValueError("The program does not handle numbers past 18 277 yet")
    
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
        
        raise ValueError("The program does not handle columns past 'ZZZ' yet")

    # -------------------------------------------------------------------------
    # 4 - Methods useful for debugging
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
