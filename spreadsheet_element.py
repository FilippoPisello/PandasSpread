from typing import List

class SpreadsheetElement:
    """
    TBD
    """
    def __init__(self, coordinates: List[List], style=None):
        self.coordinates = coordinates
        self.style = style

    @property
    def cells(self):
        from spreadsheet import Spreadsheet
        return Spreadsheet.cells(self.coordinates)

    @property
    def cells_range(self):
        from spreadsheet import Spreadsheet
        return Spreadsheet.cells_range(self.coordinates)