from abc import ABC
from dataclasses import dataclass
from reporter.exceptions import UnsupportedNameLength
from typing import Any
from datetime import datetime
import re


@dataclass
class Rules:
    intersect_rules: list
    else_rule: Any


@dataclass
class ColumnType(ABC):
    value: Any
    column_number: int

    def __repr__(self):
        return f"Value: {self.value}"


class FIO(ColumnType):

    def __init__(self, value: str):
        self.value: str = value.strip()
        self.first_name = None
        self.surname = None
        self.second_name = None
        self.parse_fio()

    @property
    def value(self):
        return self.get_surname(), self.get_first_name(), self.get_second_name()

    @value.setter
    def value(self, value):
        self._value = value

    def parse_fio(self):
        fio_array = re.split(r' +', self._value)
        if len(fio_array) == 2:
            self.surname = fio_array[0]
            self.first_name = fio_array[1]
            self.second_name = ''
        elif len(fio_array) == 3:
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


@dataclass
class ClinicCode(ColumnType):
    value: str
    column_number: int = 8


@dataclass
class ControlCode(ColumnType):
    value: str
    column_number: int = 9


@dataclass
class MedicinesID(ColumnType):
    value: int
    column_number: int = 10


@dataclass
class Codes:
    clinic_code: ClinicCode
    control_code: ControlCode
    medicine_id: MedicinesID

    @property
    def value(self):
        return self.clinic_code, self.control_code, self.medicine_id


@dataclass
class CodeFilter:
    value: str
    rules: Rules

    def get_codes(self):
        statement = self.value.strip().split()
        for rule in self.rules.intersect_rules:
            if rule.get("rule") & set(statement):
                return rule.get("value")
        return self.rules.else_rule
