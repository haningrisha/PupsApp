from reporter.base import Report, NullReport, ENDING_ROW_CELLS
import reporter.column_types as ct


ATTACH_CONFIG = {
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='САО "ВСК" КДЦ'),
            control_code=ct.ControlCode(value='ДС10к045СК/2016/16180SMU00009'),
            medicine_id=ct.MedicinesID(value=4001554)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='САО "ВСК" ПК'),
            control_code=ct.ControlCode(value='ДС19к17180SMU00092/9020'),
            medicine_id=ct.MedicinesID(value=4001554)
        )
    ),
    "header_row": {
        "фио": ct.FIO,
        "др": ct.BirthDay,
        "полис": ct.Policy,
        "дата действия полиса с": ct.DateFrom,
        "дата действия полиса по": ct.DateTo,
        "дата действия полиса": ct.DateFrom,
        "дата действия полиса до": ct.DateTo
    },
    "ending_row_cells": ENDING_ROW_CELLS
}

DETACH_CONFIG = {
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='САО "ВСК" КДЦ'),
            control_code=ct.ControlCode(value='ДС10к045СК/2016/16180SMU00009'),
            medicine_id=ct.MedicinesID(value=2)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='САО "ВСК" ПК'),
            control_code=ct.ControlCode(value='ДС19к17180SMU00092/9020'),
            medicine_id=ct.MedicinesID(value=2)
        )
    ),
    "header_row": {
        "фио": ct.FIO,
        "др": ct.BirthDay,
        "полис": ct.Policy,
        "дата действия полиса по": ct.DateCancel,
        "дата действия полиса до": ct.DateCancel
    },
    "ending_row_cells": ENDING_ROW_CELLS
}


def get_vsk(file, attach=False, detach=False):
    if Report.is_reportable(file):
        if attach:
            return Report(file, ATTACH_CONFIG)
        elif detach:
            return Report(file, DETACH_CONFIG)
    else:
        return NullReport(file)
