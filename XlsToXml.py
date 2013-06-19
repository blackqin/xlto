import xlrd

# Excel parser
class XlsToXmlParser:
	# Constants
	RESERVED_ROWS			=	4
	CELL_TYPE_ROW_INDEX		=	0
	DATA_TYPE_ROW_INDEX		=	1
	NAME_ROW_INDEX			=	2
	COMMENT_ROW_INDEX		=	3

	# Initialization
	def __init__(self):
		pass

	#
	def parseDir(self, inputDir, outputDir):
		print "Warning: Incomplete..."

	#
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
			print "Warning: " + name + " sheet is empty!"
			return

		# Generate xml string~
		xmlStr = self._toXmlStr(sheet, rows)
		print "\n" + xmlStr + "\n"

		# Save as xml file~
		self._saveXml(xmlStr, name, outputDir)

	# Parse one sheet to xml string~
	def _toXmlStr(self, sheet, rows):
		rootTagName = sheet.name
		xmlStr = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\n"
		xmlStr += "<" + rootTagName + ">\n"

		for i in range(self.RESERVED_ROWS, self.RESERVED_ROWS + rows):
			xmlStr += "\t" + self._toXmlRow(sheet, i) + "\n"

		xmlStr += "</" + rootTagName + ">"

		return xmlStr

	# Parse one row to xml string~
	def _toXmlRow(self, sheet, rowIndex):
		row = sheet.row(rowIndex)
		colIndex = 0
		cellName = ""
		xmlStr = "<"

		for cell in row:
			cellName = sheet.cell_value(self.NAME_ROW_INDEX, colIndex)
			xmlStr += cellName + "=\"" + unicode(cell.value) + "\" "
			colIndex += 1
		xmlStr += "/>"

		return xmlStr

	# Save as xml file~
	def _saveXml(self, xmlStr, fileName, outputDir):
		utf8Str = xmlStr.encode("utf-8")
		fileName += ".xml"

		file = open(outputDir + "/" + fileName, 'w')
		file.write(utf8Str)
		file.close()
"""
	# Parse xls file(s) which could be found at 'filePathList'~
	def parse(self, filePathList):
		if type(filePathList) is str:
			self.parseXls(filePathList)
		elif type(filePathList) is list:
			for filePath in filePathList:
				self.parseXls(filePath)
"""
