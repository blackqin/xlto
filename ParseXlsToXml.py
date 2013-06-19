import sys
from XlsToXml import *

# Parse xls files passed as arguments~
def parseSpecifiedXlsFiles(xlsFiles):
	parser = XlsToXmlParser()
	parser.parse(xlsFiles)

# Parse all xls files in current folder~
def parseAllXlsFiles():
	print("Error: No xls file accepted!")


# Parse xls file(s) from arguments, or all xls files in current folder~
argvLen = len(sys.argv)

parser = XlsToXmlParser()
#parser.parseDir("input", "output")
#parser.parseXls("input/test1.xls", "output")
parser.parseXlsList(["input/test001.xls", "input/test002.xls"], "output")

"""
if argvLen > 1:
	parseSpecifiedXlsFiles(sys.argv[1:])
else:
	parseAllXlsFiles()
"""

"""
exec test1.xls
exec test1.xls
"""
