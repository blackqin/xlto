import os
import xlrd3 as xlrd

# Excel parser
class XlsToXml:
	# Constants
	RESERVED_ROWS			=	4
	CELL_TYPE_ROW_INDEX		=	0
	DATA_TYPE_ROW_INDEX		=	1
	NAME_ROW_INDEX			=	2
	COMMENT_ROW_INDEX		=	3

	# Initialization
	def __init__(self):
		pass

	# Parse all xls files in "inputDir"
	def parseDir(self, inputDir, outputDir):
		for dirPath, dirNames, fileNames in os.walk(inputDir):
			for fileName in fileNames:
				filePath = os.path.join(dirPath, fileName)
				self.parseXls(filePath, outputDir)

	# Parse more than one xls file at a time~
	def parseXlsList(self, filePathList, outputDir):
		for filePath in filePathList:
			self.parseXls(filePath, outputDir)

	# Parse one xls file~
	def parseXls(self, filePath, outputDir):
		xls = xlrd.open_workbook(filePath)

		for sheet in xls.sheets():
			self._parseSheet(sheet, outputDir)

	# Parse one sheet~
	def _parseSheet(self, sheet, outputDir):
		name = sheet.name
		rows = sheet.nrows - self.RESERVED_ROWS
		cols = sheet.ncols

		# Do nothing if the sheet is empty
		if rows <= 0:
			return

		# Generate xml string~
		xmlStr = self._toXmlStr(sheet, rows)

		# Save as xml file~
		self._saveXml(xmlStr, name, outputDir)

	# Convert one sheet to xml string~
	def _toXmlStr(self, sheet, rows):
		rootTagName = sheet.name
		xmlStr = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n"
		xmlStr += "<" + rootTagName + ">\n"

		for i in range(self.RESERVED_ROWS, self.RESERVED_ROWS + rows):
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
			cellName = sheet.cell_value(self.NAME_ROW_INDEX, colIndex)
			cellValue = self._correctCellValue(cell);
			xmlStr += str(cellName) + "=\"" + str(cellValue) + "\" "
			colIndex += 1
		xmlStr += "/>"

		return xmlStr

	# Correct cell value for sure~
	def _correctCellValue(self, cell):
		if cell.ctype == xlrd.XL_CELL_NUMBER and cell.value == int(cell.value):
			return int(cell.value)
		else:
			return cell.value

	# Save as xml file~
	def _saveXml(self, xmlStr, fileName, outputDir):
		filePath = os.path.join(outputDir, fileName + ".xml")

		file = open(filePath, 'w', encoding='utf-8')
		file.write(xmlStr)
		file.close()
