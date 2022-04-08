from abc import ABC, abstractmethod
from typing import Union


class IndexIdentifier(ABC):
    def __init__(self, identifier: Union[int, str]):
        self.identifier = identifier

    @abstractmethod
    def as_integer(self):
        pass


class RowIdentifier(IndexIdentifier):
    def as_integer(self):
        if isinstance(self.identifier, int):
            return self.identifier
