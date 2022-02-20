from functools import partial

from .report import Report, NullReport, ENDING_ROW_CELLS, Config
from . import column_types as ct

ATTACH_CONFIG = Config({
    "header_row": {
        "фио": ct.FIO,
        "дата рождения": ct.BirthDay,
        "№ полиса": ct.Policy,
        "период обслуживания": {
            "с": ct.DateFrom,
            "по": ct.DateTo
        },
        "группа, № договора, организация": partial(
            ct.CodeFilter,
            rules=ct.Rules(
                intersect_rules=[
                    {
                        "rule": {"СЕМРАШ"},
                        "value": (ct.Codes(
                            clinic_code=ct.ClinicCode(value='АльфаСЕМРАШ ПК'),
                            control_code=ct.ControlCode(value='980/24/10-15/21'),
                            medicine_id=ct.MedicinesID(value=921)
                        ),)
                    }
                ],
                else_rule=(
                    ct.Codes(  # KDC
                        clinic_code=ct.ClinicCode(value='Альфа КДЦ Первичные консультации'),
                        control_code=ct.ControlCode(value='ДС №2 к дог. №74-21 от 01.10.2021 г.'),
                        medicine_id=ct.MedicinesID(value=921)
                    ),
                    ct.Codes(  # PK
                        clinic_code=ct.ClinicCode(value='Альфа ПК Первичные консультации'),
                        control_code=ct.ControlCode(value='ДС №2 к дог. №73-21 от 01.10.2021 г.'),
                        medicine_id=ct.MedicinesID(value=921)
                    )
                )
            ),
        )
    },
    "ending_row_cells": ENDING_ROW_CELLS,
    "is_code_filtered": True
})

DETACH_CONFIG = Config({
    "header_row": {
        "фио": ct.FIO,
        "дата рождения": ct.BirthDay,
        "№ полиса": ct.Policy,
        "дата открепления с (с данной даты не обслуживается)": ct.DateCancel,
        "группа, № договора, организация": partial(
            ct.CodeFilter,
            rules=ct.Rules(
                intersect_rules=[
                    {
                        "rule": {"СЕМРАШ"},
                        "value": (ct.Codes(
                            clinic_code=ct.ClinicCode(value='АльфаСЕМРАШ ПК'),
                            control_code=ct.ControlCode(value='980/24/10-15/21'),
                            medicine_id=ct.MedicinesID(value=2)
                        ),)
                    }
                ],
                else_rule=(
                    ct.Codes(  # KDC
                        clinic_code=ct.ClinicCode(value='Альфа КДЦ Первичные консультации'),
                        control_code=ct.ControlCode(value='ДС №2 к дог. №74-21 от 01.10.2021 г.'),
                        medicine_id=ct.MedicinesID(value=2)
                    ),
                    ct.Codes(  # PK
                        clinic_code=ct.ClinicCode(value='Альфа ПК Первичные консультации'),
                        control_code=ct.ControlCode(value='ДС №2 к дог. №73-21 от 01.10.2021 г.'),
                        medicine_id=ct.MedicinesID(value=2)
                    )
                )
            ),
        )
    },
    "ending_row_cells": ENDING_ROW_CELLS,
    "is_code_filtered": True
})


def get_alfa(file, attach=False, detach=False):
    if Report.is_reportable(file):
        if attach:
            return Report(file, ATTACH_CONFIG)
        elif detach:
            return Report(file, DETACH_CONFIG)
    else:
        return NullReport(file)
