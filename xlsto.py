import xlrd3 as xlrd
from xlto import XlTo

# Base class to parse .xls file(s) to any other expecting file format~
class XlsTo(XlTo):
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

    @property
    def outputFileExt(self):
        return self._outputFileExt

    # Parse one xls file~
    def parseFile(self, filePath, outputDir):
        self._xls = xlrd.open_workbook(filePath)

        for sheet in self._xls.sheets():
            if self._isValidSheet(sheet):
                outputContent = self.parseSheet(sheet, outputDir)
                outputFileName = sheet.name + self._outputFileExt
                self._saveFile(outputDir, outputFileName, outputContent)

    # Check validity of the sheet~
    def _isValidSheet(self, sheet):
        # Sheet is empty~
        rows = sheet.nrows - self.reservedRows
        cols = sheet.ncols

        if rows <= 0 or cols <= 0:
            return False

        # Unnecessary to parse the sheet if its first letter is neither uppercase nor lowercase~
        firstChar = sheet.name[0]
        firstCharCode = ord(firstChar)

        if (firstCharCode < 48):
            return False

        return True

    # Virtual method to parse one xls sheet to expecting file~
    def parseSheet(self, sheet, outputDir):
        print("[XlsTo] Virtual method to parse one sheet~")
