class PupsAppException(Exception):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class UnsupportedNameLength(PupsAppException):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class UnrecognisedType(PupsAppException):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class UnsupportedRenCode(PupsAppException):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class UnsupportedSogazCode(PupsAppException):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message
