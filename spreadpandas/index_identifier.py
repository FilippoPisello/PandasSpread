from __future__ import annotations

from abc import ABC, abstractmethod


class IndexIdentifier(ABC):
    def __init__(self, identifier: int | str):
        self.identifier = identifier

    @abstractmethod
    def as_integer(self):
        pass


class RowIdentifier(IndexIdentifier):
    def as_integer(self):
        if isinstance(self.identifier, int):
            return self.identifier
