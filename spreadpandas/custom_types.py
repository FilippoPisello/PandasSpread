"""Module where custom types are defined"""
from __future__ import annotations

from typing import Tuple

Coordinates = Tuple[int, int]
CoordinatesPair = Tuple[Coordinates, Coordinates]
Cell = str
Cells = Tuple[Cell, ...]
CellsRange = str
