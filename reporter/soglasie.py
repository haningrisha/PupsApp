from .report import Report, Config, ENDING_ROW_CELLS, NullReport
from .column_types import DateTo, DateFrom, DateCancel
from . import column_types as ct


ATTACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='Согласие КДЦ'),
            control_code=ct.ControlCode(value='0021/СК'),
            medicine_id=ct.MedicinesID(value=921)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='Согласие ПК'),
            control_code=ct.ControlCode(value='150'),
            medicine_id=ct.MedicinesID(value=921)
        )
    ),
    "header_row": {
        "фамилия": ct.Surname,
        "имя": ct.FirstName,
        "отчество": ct.SecondName,
        "д/р": ct.BirthDay,
        "№ полиса дмс": ct.Policy,
    },
    "ending_row_cells": ENDING_ROW_CELLS
})

DETACH_CONFIG = Config({
    "codes": (
        ct.Codes(  # KDC
            clinic_code=ct.ClinicCode(value='Согласие КДЦ'),
            control_code=ct.ControlCode(value='0021/СК'),
            medicine_id=ct.MedicinesID(value=2)
        ),
        ct.Codes(  # PK
            clinic_code=ct.ClinicCode(value='Согласие ПК'),
            control_code=ct.ControlCode(value='150'),
            medicine_id=ct.MedicinesID(value=2)
        )
    ),
    "header_row": {
        "ф.и.о": ct.FIO,
        "дата рождения": ct.BirthDay,
        "нормер полиса": ct.Policy,
    },
    "ending_row_cells": ENDING_ROW_CELLS
})


class SoglasieReport(Report):
    def get_data(self):
        self._get_tables()
        self.typed_tables = [self._make_typed_table(table) for table in self.tables]
        self._add_dates()
        self._make_final_table()
        self._add_codes()
        return self.final_table

    def _add_dates(self):
        pass


class SoglasieReportAttach(SoglasieReport):
    def _add_dates(self):
        for table, dates in zip(self.typed_tables, self._get_data_before_table()):
            date_from = DateFrom(dates["date_start"])
            date_to = DateTo(dates["date_end"])
            for row in table:
                row.append(date_from)
                row.append(date_to)

    def _get_data_before_table(self):
        dates = []
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    if "прикрепление" in cell.value.lower().split():
                        date_start = row[cell.column].value
                        date_end = row[cell.column + 3].value
                        dates.append({"date_start": date_start, "date_end": date_end})
        return dates


class SoglasieReporDetach(SoglasieReport):
    def _add_dates(self):
        for table, dates in zip(self.typed_tables, self._get_data_before_table()):
            date_cancel = DateCancel(dates["date_cancel"])
            for row in table:
                row.append(date_cancel)

    def _get_data_before_table(self):
        dates = []
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    if 'снять с:' == cell.value.lower().strip():
                        date_cancel = row[cell.column].value
                        dates.append({"date_cancel": date_cancel})
        return dates


def get_soglasie(file, attach=False, detach=False):
    if Report.is_reportable(file):
        if attach:
            return SoglasieReportAttach(file, ATTACH_CONFIG)
        elif detach:
            return SoglasieReporDetach(file, DETACH_CONFIG)
    else:
        return NullReport(file)
