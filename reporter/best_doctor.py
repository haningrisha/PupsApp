from .base import Report, NullReport, ENDING_ROW_CELLS, Config
from . import column_types as ct

ATTACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='ООО "БестДоктор" КДЦ'),
            control_code=ct.ControlCode(value='055/ЮЛ/2018'),
            medicine_id=ct.MedicinesID(value=921)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='ООО "БестДоктор" ПК'),
            control_code=ct.ControlCode(value='9322'),
            medicine_id=ct.MedicinesID(value=921)
        )
    ),
    "header_row": {
        "фамилия": ct.Surname,
        "имя": ct.FirstName,
        "отчество": ct.SecondName,
        "дата рождения": ct.BirthDay,
        "№ полиса": ct.Policy,
        "дата начала обслуживания": ct.DateFrom,
        "дата окончания обслуживания": ct.DateTo
    },
    "ending_row_cells": ENDING_ROW_CELLS
})

DETACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='ООО "БестДоктор" КДЦ'),
            control_code=ct.ControlCode(value='055/ЮЛ/2018'),
            medicine_id=ct.MedicinesID(value=2)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='ООО "БестДоктор" ПК'),
            control_code=ct.ControlCode(value='9322'),
            medicine_id=ct.MedicinesID(value=2)
        )
    ),
    "header_row": {
        "фамилия": ct.Surname,
        "имя": ct.FirstName,
        "отчество": ct.SecondName,
        "дата рождения": ct.BirthDay,
        "№ полиса": ct.Policy,
        "дата окончания обслуживания": ct.DateCancel
    },
    "ending_row_cells": ENDING_ROW_CELLS
})


def get_best_doctor(file, attach=False, detach=False):
    if Report.is_reportable(file):
        if attach:
            return Report(file, ATTACH_CONFIG)
        elif detach:
            return Report(file, DETACH_CONFIG)
    else:
        return NullReport(file)
