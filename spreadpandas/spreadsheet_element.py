"""Class to describe a portion of a spreadsheet"""
from __future__ import annotations

from dataclasses import dataclass

from . import operations
from .custom_types import Cells, CellsRange, CoordinatesPair


@dataclass
class SpreadsheetElement:
    """
    Class to describe a rectangular portion of a spreadsheet.

    Identify a rectangular part of a spreadsheet by providing a pair of
    coordinates that represent the top left and bottom right corner of the
    rectangle.

    This object allows to access the various forms in which the rectangle
    of cells can be identified (cells, coordinates, cells_range).

    Parameters
    ----------
    coordinates : tuple[tuple[int, int], tuple[int, int]]
        A tuple containing two tuples of length two, each containing two numbers
        which univocally identify a cell. The two numbers are indexes of the
        respectively the columns and rows. The two cells are the top left and
        bottom right one.
    """

    coordinates: CoordinatesPair

    @property
    def cells(self) -> Cells:
        """
        Return tuple of cells making up the object in the form ("A1", "A2").
        """
        return operations.cells_rectangle(self.coordinates)

    @property
    def cells_range(self) -> CellsRange:
        """
        Return a str giving info on the top left and bottom right cells of the
        object, in the form "A1:B3".
        """
        return operations.cells_range(self.coordinates)
