import os
import sys
from xlstoxml import XlsToXml

# Default input/output paths~
inputDir = "input"
outputDir = "output"

# Prepare to parse~
parser = XlsToXml()
#parser.reservedRows = 4
#parser.functionRowIndex = 0
#parser.nameRowIndex = 2

# Parse xls file(s) from arguments~
argLen = len(sys.argv)

# Parse from default input directory to default output directory~
# Example:  xlstoxmlexec.py
if argLen == 1:
    parser.parseDir(inputDir, outputDir)

# Parse one file or directory~
# Example(directory):   xlstoxmlexec.py input output
# Example(file):        xlstoxmlexec.py input/sheet0.xls output
# Example(files):       xlstoxmlexec.py input/sheet0.xls input/sheet1.xls input/sheet2.xls output
elif argLen >= 3:
    input = sys.argv[1:-1]
    outputDir = sys.argv[-1]

    if os.path.isdir(outputDir):
        if len(input) == 1 and os.path.isdir(input[0]):
            inputDir = input[0]
            parser.parseDir(inputDir, outputDir)
        else:
            parser.parseFileList(input, outputDir)
    else:
        print("Error: Last argument MUST be directory!")
else:
    print("Error: Invalid arguments!")
