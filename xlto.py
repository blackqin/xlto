import os

# Base class to parse Excel file(s) to any other expecting file format~
class XlTo:
    # Initialization
    def __init__(self):
        pass

    # Parse all files in "inputDir"
    def parseDir(self, inputDir, outputDir):
        for dirPath, dirNames, fileNames in os.walk(inputDir):
            for fileName in fileNames:
                filePath = os.path.join(dirPath, fileName)
                self.parseFile(filePath, outputDir)

    # Parse a list of files~
    def parseFileList(self, filePathList, outputDir):
        for filePath in filePathList:
            self.parseFile(filePath, outputDir)

    # Parse one file~
    def parseFile(self, filePath, outputDir):
        print("[XlTo] Virtual method to parse one file~")

    # Virtual method to parse one xls sheet to expecting file~
    def parseSheet(self, sheet, outputDir):
        print("[XlTo] Virtual method to parse one sheet~")

    # Save as expecting file format~
    def _saveFile(self, outputDir, fileName, fileContent, fileEncoding="utf-8"):
        filePath = os.path.join(outputDir, fileName)

        file = open(filePath, 'w', encoding=fileEncoding)
        file.write(fileContent)
        file.close()
