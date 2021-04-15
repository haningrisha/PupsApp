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


class UnsupportedDateQuantity(PupsAppException):
    def __init__(self, date_cell, file):
        self.expression = "Неподдерживаемый формат даты"
        self.message = f"Найдено слишком много дат в ячейке \n \"{date_cell}\" \n файла {file}"


class NoDateFound(PupsAppException):
    def __init__(self, date_cell, file):
        self.expression = "Дата не распозанан"
        self.message = f"Не удалось найти дату в ячейке\n \"{date_cell}\"\n файла {file}"


class FileTypeWasNotDefined(PupsAppException):
    def __init__(self, file):
        self.expression = "Не удалось определить тип файла"
        self.message = f"Не удалось определить тип файла (прикреп/откреп) для файла {file}"
