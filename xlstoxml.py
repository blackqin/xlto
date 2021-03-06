import xlrd3 as xlrd
from xlsto import XlsTo

# Parse .xls file(s) to .xml file(s)~
class XlsToXml(XlsTo):
    # Initialization
    def __init__(self):
        self.reservedRows = 4
        self.functionRowIndex = 0
        self.typeRowIndex = 1
        self.nameRowIndex = 2
        self.commentRowIndex = 3
        self._outputFileExt = ".xml"

    # Parse one sheet~
    def parseSheet(self, sheet, outputDir):
        return self._toXmlStr(sheet)

    # Convert one sheet to xml string~
    def _toXmlStr(self, sheet):
        rows = sheet.nrows - self.reservedRows
        sheetName = sheet.name
        xmlStr = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n"
        xmlStr += "<Sheet name=\"" + sheetName + "\">\n"

        for i in range(self.reservedRows, self.reservedRows + rows):
            xmlStr += "\t" + self._toXmlRow(sheet, i) + "\n"

        xmlStr += "</Sheet>"

        return xmlStr

    # Convert one row to xml string~
    def _toXmlRow(self, sheet, rowIndex):
        row = sheet.row(rowIndex)
        colIndex = 0
        cellName = ""
        xmlStr = "<Row "

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
