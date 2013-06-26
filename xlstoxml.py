import xlrd3 as xlrd
from xlto import XlTo

#TODO
# root tag attributes. (<Sheet version="1.02">)

# Excel parser
class XlsToXml(XlTo):
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
            if (cellName != "" and cellValue != ""):
                xmlStr += cellName + "=\"" + cellValue + "\" "

            colIndex += 1

        xmlStr += "/>"

        return xmlStr

    # Parse one cell~
    def _parseCell(self, sheet, cell, colIndex):
        cellName = (str(sheet.cell_value(self.nameRowIndex, colIndex))).strip()
        cellType = cell.ctype
        cellValue = cell.value

        if cellType == xlrd.XL_CELL_NUMBER:
            intValue = int(cellValue)
            if cellValue == intValue:
                cellValue = intValue
        elif cellType == xlrd.XL_CELL_DATE:
            timeTuple = xlrd.xldate_as_tuple(cellValue, self._xls.datemode)
            cellValue = self._toTimeStr(timeTuple)

        return (cellName, (str(cellValue)).strip())

    # Convert tuple like (2013, 12, 31, 23, 59, 59) to string '2013/12/31 23:59:59'~
    def _toTimeStr(self, timeTuple):
        year = str(timeTuple[0])
        month = str(timeTuple[1])
        date = str(timeTuple[2])
        hours = str(timeTuple[3])
        minutes = str(timeTuple[4])
        seconds = str(timeTuple[5])
        timeStr = year + "/" + month + "/" + date + " " + hours + ":" + minutes + ":" + seconds

        return timeStr
