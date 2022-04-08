import pytest
from spreadpandas import operations


@pytest.mark.parametrize(
    "index, expected_letter, expect_value_error",
    [
        (0, "A", False),
        (1, "B", False),
        (25, "Z", False),
        (26, "AA", False),
        (18277, "ZZZ", False),
        (18278, None, True),
        (-1, None, True),
    ],
)
def test_letter_from_index(index: int, expected_letter: str, expect_value_error: bool):
    if expect_value_error:
        with pytest.raises(ValueError):
            operations.letter_from_index(index)
        return

    assert operations.letter_from_index(index) == expected_letter


@pytest.mark.parametrize(
    "letter, expected_index, expect_value_error",
    [
        ("A", 0, False),
        ("B", 1, False),
        ("Z", 25, False),
        ("AA", 26, False),
        ("ZZZ", 18277, False),
        ("AAAA", None, True),
    ],
)
def test_index_from_letter(letter: str, expected_index: int, expect_value_error: bool):
    if expect_value_error:
        with pytest.raises(ValueError):
            operations.index_from_letter(letter)
        return

    assert operations.index_from_letter(letter) == expected_index


@pytest.mark.parametrize(
    "index, expected_row_number, expect_value_error",
    [
        (0, "1", False),
        (1, "2", False),
        (999, "1000", False),
        (-1, None, True),
    ],
)
def test_row_number_from_index(
    index: int, expected_row_number: str, expect_value_error: bool
):
    if expect_value_error:
        with pytest.raises(ValueError):
            operations.row_number_from_index(index)
        return
    assert operations.row_number_from_index(index) == expected_row_number


@pytest.mark.parametrize(
    "coordinates, expected_cell_label",
    [
        ((0, 0), "A1"),
        ((2, 1), "C2"),
    ],
)
def test_cell_from_coordinates(coordinates, expected_cell_label):
    assert operations.cell_from_coordinates(coordinates) == expected_cell_label


@pytest.mark.parametrize(
    "coordinates_pair, expected_cells_range",
    [
        (((0, 0), (1, 1)), "A1:B2"),
    ],
)
def test_cells_range(coordinates_pair, expected_cells_range):
    assert operations.cells_range(coordinates_pair) == expected_cells_range
