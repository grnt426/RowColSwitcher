import glob
import os
import xlrd
from datetime import time

# Grab all XLS files in this directory
for file in glob.glob("*.xls"):

    # Convert the XLS to CSV
    wb = xlrd.open_workbook(file)
    sh = wb.sheet_by_index(0)
    csvOutput = ""
    for rownum in range(sh.nrows):
        rowVal = sh.row_values(rownum)
        csvLine = ""
        for i in range(len(rowVal)):
            val = rowVal[i]

            # Excel internally stores dates as floats, which means if we have a date, we need to convert it
            # into the "real" format that was expected, ie. 00:00.000
            cell_type = sh.cell_type(rownum - 1, i)
            if cell_type == xlrd.XL_CELL_DATE and rownum != 0:
                dt_tuple = xlrd.xldate_as_tuple(val, wb.datemode)
                val = time(*dt_tuple[3:]).strftime("%M:%S.0")

            # Apparently, we can't join on numbers. So we have to manually convert to strings
            if not isinstance(val, str):
                val = str(val)
            csvLine += val + ","
        csvOutput += csvLine + "\n"

    # Rotate the CSV
    outputFile = open(os.path.splitext(file)[0] + ".csv", "w")
    rows = [[], [], []]
    for line in csvOutput.split("\n"):
        cols = line.split(",")
        if len(cols) == 0:
            continue
        for i in range(len(cols)-1):
            rows[i].append(cols[i])
    for row in rows:
        for entry in row:
            outputFile.write(entry + ",")
        outputFile.write("\n")
    outputFile.close()