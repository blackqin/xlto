import xlrd3 as xlrd
from xlto import XlTo

#TODO
# root tag attributes. (<Sheet version="1.02">)

# Excel parser
class XlsToXml(XlTo):
    # Variables
    _reservedRows = 4
    _nameRowIndex = 2

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
        xls = xlrd.open_workbook(filePath)

        for sheet in xls.sheets():
            self._parseSheet(sheet, outputDir)

    # Parse one sheet~
    def _parseSheet(self, sheet, outputDir):
        name = sheet.name
        rows = sheet.nrows - self.reservedRows
        cols = sheet.ncols

        # Do nothing if the sheet is empty~
        if rows <= 0:
            return

        # Generate xml string~
        xmlStr = self._toXmlStr(sheet, rows)

        # Save as xml file~
        fileName = name + ".xml"
        self._saveFile(outputDir, fileName, xmlStr)

    # Convert one sheet to xml string~
    def _toXmlStr(self, sheet, rows):
        rootTagName = sheet.name
        xmlStr = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n"
        xmlStr += "<" + rootTagName + ">\n"

        for i in range(self.reservedRows, self.reservedRows + rows):
            xmlStr += "\t" + self._toXmlRow(sheet, i) + "\n"

        xmlStr += "</" + rootTagName + ">"

        return xmlStr

    # Convert one row to xml string~
    def _toXmlRow(self, sheet, rowIndex):
        row = sheet.row(rowIndex)
        colIndex = 0
        cellName = ""
        xmlStr = "<row "

        for cell in row:
            (cellName, cellValue) = self._parseCell(sheet, cell, colIndex)

            # Don't add empty cells~
            if (cellName != "" or cellValue != ""):
                xmlStr += cellName + "=\"" + str(cellValue) + "\" "

            colIndex += 1

        xmlStr += "/>"

        return xmlStr

    # Parse one cell~
    def _parseCell(self, sheet, cell, colIndex):
        cellName = (str(sheet.cell_value(self.nameRowIndex, colIndex))).strip()
#TODO
# parse these types: int, uint, bool, time, str
        if cell.ctype == xlrd.XL_CELL_NUMBER and cell.value == int(cell.value):
            cellValue = int(cell.value)
        else:
            cellValue = cell.value

        return (cellName, cellValue)
