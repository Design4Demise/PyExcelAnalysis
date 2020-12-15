import datetime
import re
from typing import Any

import xlwings as xw
from xlwings.main import Range


class SheetManager:

    def __init__(self, path: str) -> None:

        """Class for Managing Excel Sheets.

        Parameters
        ----------
        path: str
            Path to Excel Workbook.
        """

        self.path = path
        self._verify_path(path)

        self.wb = xw.Book(path)
        self.backup_path = self._store_backup()

    def __del__(self) -> None:

        """Ensures that Workbook is saved and closed."""

        self.wb.save()
        self.wb.close()

    def get_cell(self, sheet: str, cell: str) -> Range:

        """Returns contents of given cell/s.

        Parameters
        ----------
        sheet: str
            Sheet to point to.
        cell: str
            Cell to return.

        Returns
        -------
        Range
            Excel cell/s.
        """

        return self.wb.sheets[sheet].range(cell)

    def change_cell(self, sheet: str, cell: str, value: Any):

        """Change value in given cell/s.

        Parameters
        ----------
        sheet: str
            Sheet to point to.
        cell: str
            Cell to change.
        value: Any
            Value/s to prescribe to specified cell/s.
        """

        self.get_cell(sheet, cell).value = value

    def check_cell(self, sheet: str, cell: str) -> Any:

        """Gets value/s of specified cell/s.

        Parameters
        ----------
        sheet: str
            Sheet to point to.
        cell: str
            Cell to check.

        Returns
        -------
        Any
            Value/s of specified cell/s.
        """

        return self.get_cell(sheet, cell).value

    def restore(self) -> None:

        """Restore Workbook from Backup."""

        self.wb.close()
        self.wb = xw.Book(self.backup_path)
        self.wb.save(self.path)

    def _store_backup(self) -> str:

        """Store Backup of Workbook."""

        self.wb.save()

        t_epoch = int(datetime.datetime.now().timestamp())
        backup_path = re.sub(r'(.+)(\.xlsx)', rf'{t_epoch}_\1\2', self.path)
        self.wb.save(backup_path)

        self.wb.close()
        self.wb = xw.Book(self.path)

        return backup_path

    @staticmethod
    def _verify_path(path: str) -> None:

        """Verifies file is .xlsx.

        Parameters
        ----------
        path: str
            Path to verify.
        """

        if not path.endswith('.xlsx'):
            raise ValueError('File must have .xlsx extension.')
