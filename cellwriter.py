import xlrd
import xlutils.copy
import xlwt
from xlwt import Workbook


class Writer():
    def __init__(self, directory):
        self.directory = directory
        self.inBook = xlrd.open_workbook(self.directory, formatting_info=True)
        self.outBook = xlutils.copy.copy(self.inBook)
        self.outSheet = self.outBook.get_sheet(0)

    def _getOutCell(self, colIndex, rowIndex):
        row = self.outSheet._Worksheet__rows.get(rowIndex)
        if not row:
            return None
        cell = row._Row__cells.get(colIndex)
        return cell

    def _setOutCell(self, col, row, value):
        previousCell = self._getOutCell( col, row)
        self.outSheet.write(row, col, value)
        if previousCell:
            newCell = self._getOutCell( col, row)
            if newCell:
                newCell.xf_idx = previousCell.xf_idx

    def write(self, col, row, text):
        self._setOutCell( row, col, text)

    def save(self):
        self.outBook.save(self.directory)
