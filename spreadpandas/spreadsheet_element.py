"""Contain the class to describe a portion of a spreadsheet"""
from . import operations
from .custom_types import Cells, CellsRange, CoordinatesPair


class SpreadsheetElement:
    """
    Class to describe a rectangular portion of a spreadsheet.

    ---------------
    Identify a rectangular part of a spreadsheet by providing a pair of
    coordinates that identify that top left and bottom right corner of the
    rectangle.

    The object is useful to access the various forms in which the rectangle
    of cells can be identified (cells, coordinates, cells_range).

    Arguments
    ----------------
    coordinates: list in the form [(int, int), (int, int)], (mandatory)
        A list containing two tuples of length two, each containing two numbers
        which univocally identify a cell. The two numbers are respectively the
        row index and the column number. The two cells are the top left and
        bottom right one.
    """

    def __init__(self, coordinates: CoordinatesPair):
        self.coordinates = coordinates

    @property
    def cells(self) -> Cells:
        """
        Returns tuple of cells making up the object in the form ("A1", "A2").
        """
        return operations.cells_rectangle(self.coordinates)

    @property
    def cells_range(self) -> CellsRange:
        """
        Returns a str giving info on the top left and bottom right cells of the
        object, in the form "A1:B3".
        """
        return operations.cells_range(self.coordinates)
