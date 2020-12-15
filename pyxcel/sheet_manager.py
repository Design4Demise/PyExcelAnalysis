import xlwings as xw
from xlwings.main import Range
from typing import Any
import datetime
import re


class SheetManager:

    def __init__(self, path: str) -> None:

        self.path = path
        self._verify_path(path)

        self.wb = xw.Book(path)
        self.backup_path = self._store_backup()

    def __del__(self) -> None:
        self.wb.save()
        self.wb.close()

    def get_cell(self, sheet: str, cell: str) -> Range:
        return self.wb.sheets[sheet].range(cell)

    def change_cell(self, sheet: str, cell: str, value: Any):
        self.get_cell(sheet, cell).value = value

    def check_cell(self, sheet: str, cell: str) -> Any:
        return self.get_cell(sheet, cell).value

    def restore(self) -> None:
        self.wb.close()
        self.wb = xw.Book(self.backup_path)
        self.wb.save(self.path)

    def _store_backup(self) -> str:

        self.wb.save()

        t_epoch = int(datetime.datetime.now().timestamp())
        backup_path = re.sub(r'(.+)(\.xlsx)', rf'{t_epoch}_\1\2', self.path)
        self.wb.save(backup_path)

        self.wb.close()
        self.wb = xw.Book(self.path)

        return backup_path

    @staticmethod
    def _verify_path(path: str) -> None:
        if not path.endswith('.xlsx'):
            raise ValueError('File must have .xlsx extension.')
