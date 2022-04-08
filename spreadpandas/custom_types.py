"""Module where custom types are defined"""
from __future__ import annotations

from typing import Tuple

Coordinate = Tuple[int, int]
CoordinatesPair = Tuple[Coordinate, Coordinate]
Cell = str
Cells = Tuple[Cell, ...]
CellsRange = str
