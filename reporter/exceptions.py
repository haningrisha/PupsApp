class UnsupportedNameLength(Exception):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class UnrecognisedType(Exception):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message


class UnsupportedRenCode(Exception):
    def __init__(self, expression, message):
        self.expression = expression
        self.message = message
