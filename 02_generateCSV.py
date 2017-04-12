'''
    Generate CSV and Excel files
'''

### Files we will use
check_file = "_input_checkPCS.txt"
csv_file = "EXCEL\\overview_"

### CSV stuff specific file
csv_head1 = ["PC name"]
csv_head2 = len(csv_head1) * [""]

### Databases we will use
db_checkedpcs = "DB\\checkedpcs.db"
db_foundsoft = "DB\\foundsoft.db"

### Modules we will use
import shelve
import csv
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

### Opening shelves
s_checkedpcs = shelve.open(db_checkedpcs)
s_foundsoft = shelve.open(db_foundsoft)

### Functions
def readFile(file="in.txt"):
    print("Reading from",file)
    print()
    lst = []
    with open(file,"r") as f:
        for l in f:
            line = l.strip("\n")
            line.strip()
            if line.startswith("#") or len(line)<2:
                continue
            else:
                lst.append(line)
    return lst

# Generate CSV file
def generateCSV(fileName="",pc="",allSoft=[]):
    print("Generating CSV.")
    with open(fileName,"w") as f:
        writer = csv.writer(f)
        header1 = csv_head1 + allSoft
        csv_head2[0] = pc
        header2 = csv_head2 + len(allSoft) * ["X"]
        writer.writerow(header1)
        writer.writerow(header2)    
    
def generateXL(fileName="",fileName_csv=""):
    delimiter=","
    numRows = 0
    
    print("Generating Excel file")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    
    with open(fileName_csv) as f:
        reader = csv.reader(f, delimiter=delimiter)
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
                ws.column_dimensions[col].width = 4
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
        wb.save(fileName_xl)
        print()
        
### Actual program
allPCS = readFile(check_file)
allSoft = []

for pc in allPCS:
    print("Getting software list for",pc)
    fileName_csv = csv_file + pc + ".csv"
    fileName_xl = csv_file + pc + ".xlsx"
    for soft in s_foundsoft[pc]:
        allSoft.append(soft)
    generateCSV(fileName=fileName_csv,pc=pc,allSoft=allSoft)
    generateXL(fileName=fileName_xl,fileName_csv=fileName_csv)
    allSoft=[]

print("All done.")
print()
    
### Closing shelves
s_checkedpcs.close()
s_foundsoft.close()