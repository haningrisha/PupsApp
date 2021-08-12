from abc import ABC
from dataclasses import dataclass
from reporter.exceptions import UnsupportedNameLength
from typing import Any
from datetime import datetime


@dataclass
class ColumnType(ABC):
    value: Any
    column_number: int

    def __repr__(self):
        return f"Value: {self.value}"


class FIO(ColumnType):

    def __init__(self, value: str):
        self._value: str = value
        self.first_name = None
        self.surname = None
        self.second_name = None
        self.parse_fio()

    @property
    def value(self):
        return self.get_surname(), self.get_first_name(), self.get_second_name()

    def parse_fio(self):
        fio_array = self._value.split(" ")
        if len(fio_array) == 3:
            self.surname = fio_array[0]
            self.first_name = fio_array[1]
            self.second_name = fio_array[2]
        elif len(fio_array) == 4:
            self.surname = fio_array[0] + fio_array[1]
            self.first_name = fio_array[2]
            self.second_name = fio_array[3]
        else:
            raise UnsupportedNameLength("Ошибка имени",
                                        "Неподдерживаемый формат имени '{0}'"
                                        .format(self._value))

    def get_surname(self):
        return Surname(value=self.surname)

    def get_first_name(self):
        return FirstName(value=self.first_name)

    def get_second_name(self):
        return SecondName(value=self.second_name)


@dataclass
class Surname(ColumnType):
    value: str
    column_number: int = 0


@dataclass
class FirstName(ColumnType):
    value: str
    column_number: int = 1


@dataclass
class SecondName(ColumnType):
    value: str
    column_number: int = 2


@dataclass
class BirthDay(ColumnType):
    value: datetime
    column_number: int = 3


@dataclass
class Policy(ColumnType):
    value: str
    column_number: int = 4


@dataclass
class DateFrom(ColumnType):
    value: datetime
    column_number: int = 5


@dataclass
class DateTo(ColumnType):
    value: datetime
    column_number: int = 6


@dataclass
class DateCancel(ColumnType):
    value: datetime
    column_number: int = 7

