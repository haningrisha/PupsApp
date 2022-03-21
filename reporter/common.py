import enum


class AccessTypes(enum.Enum):
    PK = enum.auto()
    KDC = enum.auto()
    SMT = enum.auto()
    UNDEFINED = enum.auto()
