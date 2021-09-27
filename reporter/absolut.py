from .base import Report, NullReport, ENDING_ROW_CELLS, Config
from . import column_types as ct

ATTACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='Абсолют Страхование КДЦ'),
            control_code=ct.ControlCode(value="0035/СК"),
            medicine_id=ct.MedicinesID(value=921)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='Абсолют Страхование ПК'),
            control_code=ct.ControlCode(value='9004'),
            medicine_id=ct.MedicinesID(value=921)
        )
    ),
    "header_row": {
        "фио": ct.FIO,
        "дата рождения": ct.BirthDay,
        "№ полиса": ct.Policy,
        "дата начала действия полиса": ct.DateFrom,
        "дата окончания действия полиса": ct.DateTo
    },
    "ending_row_cells": ENDING_ROW_CELLS
})

DETACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='Абсолют Страхование КДЦ'),
            control_code=ct.ControlCode(value="0035/СК"),
            medicine_id=ct.MedicinesID(value=2)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='Абсолют Страхование ПК'),
            control_code=ct.ControlCode(value='9004'),
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
            return Report(file, DETACH_CONFIG)
    else:
        return NullReport(file)
