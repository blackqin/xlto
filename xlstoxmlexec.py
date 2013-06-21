import sys
from xlstoxml import XlsToXml

# Parse xls file(s) from arguments~
argvLen = len(sys.argv)

parser = XlsToXml()
parser.parseDir("C:/Users/blackqin/Desktop/inputXls", "C:/Users/blackqin/Desktop/outputXml")
#parser.parseXls("input/test1.xls", "output")
#parser.parseXlsList(["input/test001.xls", "input/test002.xls"], "output")
