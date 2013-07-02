import sys
from xlstoxml import XlsToXml

# Parse xls file(s) from arguments~
argvLen = len(sys.argv)
inputPath = "input"
outputPath = "output"

parser = XlsToXml()
parser.reservedRows = 4
parser.nameRowIndex = 2

parser.parseDir(inputPath, outputPath)
