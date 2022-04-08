"""Contains functions carrying out operations over the spreadsheet which
are independent from the spreadsheet instance."""
import string

from .custom_types import Cells, CellsRange, CoordinatesPair


def cells_rectangle(coordinates: CoordinatesPair) -> Cells:
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
    - If [[0,1],[1,2]] is provided, the output will be [A1, B1, A2, B2]
    """
    starting_letter_pos, starting_number = coordinates[0]
    ending_letter_pos, ending_number = coordinates[1]

    cells = tuple(
        letter_from_index(letter) + str(number)
        for number in range(starting_number, ending_number + 1)
        for letter in range(starting_letter_pos, ending_letter_pos + 1)
    )

    return cells


def cells_range(coordinates: list[list]) -> CellsRange:
    """
    Returns a string in the form "A1:B2" given coordinates in the form
    [[0,1],[1,2]]
    """
    cell1 = letter_from_index(coordinates[0][0]) + str(coordinates[0][1])
    cell2 = letter_from_index(coordinates[1][0]) + str(coordinates[1][1])
    return cell1 + ":" + cell2


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

    hundreds_letter = string.ascii_uppercase[((letter_position - 26) % (26**2)) // 26]
    if letter_position <= (26**2 + 25):
        return hundreds_letter + units_letter

    thousands_letter = string.ascii_uppercase[
        (letter_position - 26**2 - 26) // 26**2
    ]
    if letter_position <= (26**3 + 26**2 + 25):
        return thousands_letter + hundreds_letter + units_letter

    raise ValueError("The program does not handle indexes past 18 277 yet")


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

    hundreds_letter = string.ascii_uppercase.index(spreadsheet_letter[-2]) + 1
    if len(spreadsheet_letter) == 2:
        return 26 * hundreds_letter + units_letter

    thousands_letter = string.ascii_uppercase.index(spreadsheet_letter[-3]) + 1
    if len(spreadsheet_letter) == 3:
        return 26**2 * thousands_letter + 26 * hundreds_letter + units_letter

    raise ValueError("The program does not handle column letters past 'ZZZ' yet")
