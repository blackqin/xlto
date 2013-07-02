import xlrd3 as xlrd
from xlto import XlTo

# Base class to parse .xls file(s) to any other expecting file format~
class XlsTo(XlTo):
    # Variables
    _reservedRows = 4
    _nameRowIndex = 2
    _xls = None

    # Initialization
    def __init__(self):
        pass

    @property
    def reservedRows(self):
        return self._reservedRows

    @reservedRows.setter
    def reservedRows(self, value):
        self._reservedRows = value

    @property
    def nameRowIndex(self):
        return self._nameRowIndex

    @nameRowIndex.setter
    def nameRowIndex(self, value):
        self._nameRowIndex = value

    # Parse one xls file~
    def parseFile(self, filePath, outputDir):
        self._xls = xlrd.open_workbook(filePath)

        for sheet in self._xls.sheets():
            self.parseSheet(sheet, outputDir)

    # Virtual method to parse one xls sheet to expecting file~
    def parseSheet(self, sheet, outputDir):
        print("[XlsTo] Virtual method to parse one xls sheet~")
