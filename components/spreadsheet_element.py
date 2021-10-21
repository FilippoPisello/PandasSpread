"""Contain the class to describe a portion of a spreadsheet"""
from components.custom_types import Cells, CellsRange, CoordinatesPair
import components.operations as operations


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
    """

    def __init__(self, coordinates: CoordinatesPair):
        self.coordinates = coordinates

    @property
    def cells(self) -> Cells:
        """
        Returns list of cells making up the object in the form ["A1", "A2"].
        """
        return operations.cells_rectangle(self.coordinates)

    @property
    def cells_range(self) -> CellsRange:
        """
        Returns a str giving info on the top left and bottom right cells of the
        object, in the form "A1:B3".
        """

        return operations.cells_range(self.coordinates)
