from .base import Report, NullReport, ENDING_ROW_CELLS, Config
from . import column_types as ct

ATTACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='Медэкспресс КДЦ Стандарт'),
            control_code=ct.ControlCode(value='0012/СК'),
            medicine_id=ct.MedicinesID(value=921)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='САО "Медэкспресс"'),
            control_code=ct.ControlCode(value='718'),
            medicine_id=ct.MedicinesID(value=921)
        )
    ),
    "header_row": {
        "ф.и.о.": ct.FIO,
        "дата рожд.": ct.BirthDay,
        "номер полиса": ct.Policy,
        "начало полиса": ct.DateFrom,
        "конец полиса": ct.DateTo
    },
    "ending_row_cells": ENDING_ROW_CELLS
})

DETACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='Медэкспресс КДЦ Стандарт'),
            control_code=ct.ControlCode(value='0012/СК'),
            medicine_id=ct.MedicinesID(value=2)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='САО "Медэкспресс"'),
            control_code=ct.ControlCode(value='718'),
            medicine_id=ct.MedicinesID(value=2)
        )
    ),
    "header_row": {
        "фио": ct.FIO,
        "дата рождения": ct.BirthDay,
        "№ полиса": ct.Policy,
        "дата открепления": ct.DateCancel
    },
    "ending_row_cells": ENDING_ROW_CELLS
})


def get_absolut(file, attach=False, detach=False):
    if Report.is_reportable(file):
        if attach:
            return Report(file, ATTACH_CONFIG)
        elif detach:
            return NullReport(file)
    else:
        return NullReport(file)
