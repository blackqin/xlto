import sys
from xlstoxml import XlsToXml

# Parse xls file(s) from arguments~
argvLen = len(sys.argv)

parser = XlsToXml()
parser.parseDir("input", "output")
