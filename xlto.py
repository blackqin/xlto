import os

# Base class to parse Excel file(s) to other expecting file format~
class XlTo:
	# Initialization
	def __init__(self):
		pass

	# Parse all files in "inputDir"
	def parseDir(self, inputDir, outputDir):
		for dirPath, dirNames, fileNames in os.walk(inputDir):
			for fileName in fileNames:
				filePath = os.path.join(dirPath, fileName)
				self.parseXls(filePath, outputDir)

	# Parse a list of files~
	def parseFileList(self, filePathList, outputDir):
		for filePath in filePathList:
			self.parseFile(filePath, outputDir)

	# Parse one file~
	def parseFile(self, filePath, outputDir):
		print("[XlTo] Virtual method to parse an file~")

	# Save as expecting file format~
	def _saveFile(self, outputDir, fileName, fileExtension, fileContent, fileEncoding="utf-8"):
		fullFileName = fileName + "." + fileExtension
		filePath = os.path.join(outputDir, fullFileName)

		file = open(filePath, 'w', encoding=fileEncoding)
		file.write(fileContent)
		file.close()
