# Author: Filippo Pisello
from typing import List

class SpreadsheetElement:
    """
    Class to represent subportions of a spreadsheet in different ways.

    ---------------
    The class intakes as its main argument a pair of coordinates made of int
    indexes. It then allows to access various information about the table part.

    The idea is to be able to pass at will from the coordinates view, to the
    cells and cells_range ones. These all describe in different ways the same
    concept, which is the expected position of some data in a spreadsheet.

    Arguments
    ----------------
    coordinates: list in the form [[int, int], [int, int]], (mandatory)
        A list containing two lists of length two, each containing two numbers
        which univocally identify a cell. The two numbers are respectively the
        row index and the column number. The two cells are the top left and
        bottom right one.
    style: None or style object, default=None
        An object of the adequate type can be passed depending on the spreadsheet
        subclass used. For CustomExcel, one should use and ExcelStyle object. If
        None, no action will be taken.
    """
    def __init__(self, coordinates: List[List], style=None):
        self.coordinates = coordinates
        self.style = style

    @property
    def cells(self) -> List[str]:
        """
        Returns list of cells making up the object in the form ["A1", "A2"].
        """
        from spreadsheet import Spreadsheet
        return Spreadsheet.cells(self.coordinates)

    @property
    def cells_range(self) -> str:
        """
        Returns a str giving info on the top left and bottom right cells of the
        object, in the form "A1:B3".
        """
        from spreadsheet import Spreadsheet
        return Spreadsheet.cells_range(self.coordinates)
