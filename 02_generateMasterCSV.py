'''
    Check LDMS for non standard software
    uses module 
    - 'csv'             : built-in
    - 'shelve'          : built-in
'''

### Files we will use
output_file = "EXCEL\\__overview_master"

### Databases we will use
db_checkedpcs = "DB\\checkedpcs.db"
db_standardsoft = "DB\\stdsoft.db"
db_foundsoft = "DB\\foundsoft.db"
db_allsoft = "DB\\allsoft.db"
db_aliases = "DB\\aliases.db"
db_knownsoft = "DB\\knownsoft.db"

### Misc variables
allSoft = []
dictAllSoft = {}
dictAllPCInfo = {}

### Modules we will use
import shelve
import csv
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# Opening shelves
s_checkedpcs = shelve.open(db_checkedpcs)
s_foundsoft = shelve.open(db_foundsoft)
s_allsoft = shelve.open(db_allsoft)
s_knownsoft = shelve.open(db_knownsoft)

# Getting list of all software
for soft in s_allsoft:
    allSoft.append(soft)

# Sorting it alphabetically, remember the location
for num,soft in enumerate(sorted(allSoft,key=lambda s: s.lower())):
    dictAllSoft[soft]=num

# Build a dictionary with ticked alphabetically sorted software
for pc in s_foundsoft:
    lstSoft = len(dictAllSoft)*[""]
    foundsoft = s_foundsoft[pc]
    for soft in foundsoft:
        loc = dictAllSoft[soft]
        lstSoft[loc]="X"
    dictAllPCInfo[pc]=lstSoft

# Build CSV file
with open(output_file + ".csv","w") as f:
    header = ["PC name"] + sorted(allSoft,key=lambda s: s.lower())
    writer = csv.writer(f,delimiter=",",lineterminator='\n')
    writer.writerow(header)
    for pc in dictAllPCInfo:
        line = [pc]
        line += dictAllPCInfo[pc]
        writer.writerow(line)

# Build XLSX file
wb = openpyxl.Workbook()
ws = wb.active
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

with open(output_file+".csv") as f:
    numRows = 0
    reader = csv.reader(f, delimiter=",")
    for row in reader:
        ws.append(row)

    # Setting max width
    dims = {}
    for row in ws.rows:
        numRows += 1
        for cell in row:
            if cell.value:
                dims[cell.column] = max((dims.get(cell.column, 0), len(cell.value)))
    for col, value in dims.items():
        ws.column_dimensions[col].width = value
        
    # Formatting Excel file
    lstCol = list(dims.keys())
    for colNum,col in enumerate(lstCol):
        # Running over columns
        if colNum == 0:
            # column width = 17
            ws.column_dimensions[col].width = 17
        else:
            # column width = 3
            ws.column_dimensions[col].width = 2.3
        for rowNum in range(numRows):
            # Running over rows
            curCell = ws[str(col.upper()) + str(rowNum+1)]
            curCell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            if colNum > 0:
                # We do not touch the first column, borders are already set
                if rowNum < 1:
                    # Formatting should be: vertical text + all border + center
                    curCell.alignment = Alignment(horizontal='center',text_rotation=90)
                else:
                    # Formatting should be : all borders + center
                    curCell.alignment = Alignment(horizontal='center')        
    
    # Saving Excel file
    wb.save(output_file+".xlsx")
    print()

# Closing shelves
s_checkedpcs.close()
s_foundsoft.close()
s_allsoft.close()
s_knownsoft.close()
