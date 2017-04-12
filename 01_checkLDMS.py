'''
    Check LDMS for non standard software
    uses module 
    - 'requests'        : pip install requests
    - 'urllib3'         : pip install urllib3
    - 'requests_ntlm'   : pip install requests_ntlm
    - 'shelve'          : no install needed
    
    Files used:
    - H:\AD.txt (case sensitive!)
        -> Format:
            # Comment line
            ad_user = "AD\\CWID"        (note the double backslash!)
            ad_pwd = "AD password"      
'''

### Files we will use
pwd_file = "H:\\AD.txt"
log_file = "checkLDMS_"
known_file = "_input_knownSoft.txt"
check_file = "_input_checkPCS.txt"

### Databases we will use
db_checkedpcs = "DB\\checkedpcs.db"
db_standardsoft = "DB\\stdsoft.db"
db_foundsoft = "DB\\foundsoft.db"
db_allsoft = "DB\\allsoft.db"
db_aliases = "DB\\aliases.db"
db_knownsoft = "DB\\knownsoft.db"

### Misc variables
count = 0
soft = ""
foundsoft = []
foundtd = False
link = "http://ldms.glpoly.net/LDMSWeb/InventoryDatasheet.aspx?target="

### Modules we will use
import requests
import shelve
from requests_ntlm import HttpNtlmAuth

### Reading username and password from pwd_file
try:
    with open(pwd_file,"r") as f:
        for num,line in enumerate(f):
            if num == 1:
                ad_user = line.split("=")[1].strip().strip('"')
            elif num == 2:
                ad_pwd = line.split("=")[1].strip().strip('"')
            elif num > 2:
                break
except:
    err = "Passwordfile " + pwd_file + " does not exist!"
    raise FileExistsError(err)

### Functions we will use
def readFile(file="in.txt"):
    knownsoft = []
    with open(file,"r") as f:
        for l in f:
            line = l.strip("\n")
            line.strip()
            if line.startswith("#") or len(line)<2:
                continue
            else:
                knownsoft.append(line)
    return knownsoft

def readData(link="http://www.google.be",ad_user=ad_user,ad_pwd=ad_pwd):    
    session = requests.Session()
    session.auth=auth=HttpNtlmAuth(ad_user,ad_pwd)
    data = session.get(link)
    return data.text
            
def chars2line(text=""):
    chars = ""
    for char in text:
        if char == "\n":
            chars += char
            line = chars
            chars = ""
            yield line
        else:
            chars += char
            
def getCorrectData(text=""):
    table = ""
    returntable = ""
    found = False
    end = False
    for l in text:
        line = l.strip()
        if found == False:
            if line.startswith('<div id="ctl00_ctl00_MainContentPlaceHolder_CMS_MainContentPlaceHolder_lblAppsoftware"'):
                table += line + "\n"
                found = True
        else:
            if line.endswith('</table>'):
                table += line + "\n"
                end = True
                return table
            else:
                table += line + "\n"
    return table
 
def findSoft(pcname="",link=link,knownsoft=[],log_file=log_file):
    
    link += pcname
    log_file += pcname + ".txt"

    # Getting and parsing text to find our table containing the data we need
    text = chars2line(readData(link=link))
    table = getCorrectData(text=text)

    with open(log_file,"w") as f:
        for line in chars2line(table):
            if line.startswith("<td><font face"):
                # We found something. If counter > 1 we found software package
                string = line[32:]
                for num,l in enumerate(string):
                    if l == "<":
                        foundpos = num
                        break
                soft = string[:foundpos]

                if soft.startswith("Security Update for") or soft.startswith("Update for Microsoft") or soft.startswith("Service Pack") or "(KB" in soft:
                    # Updates and service packs
                    f.write("Found update (ignoring): ")
                    f.write(soft)
                    f.write("\n")
                elif soft.startswith("Lenovo"):
                    # A line I couldn't get rid of ...
                    f.write("Found Lenvo software. (ignoring): ")
                    f.write(soft)
                    f.write("\n")
                elif soft.startswith('ack" size="2">'):
                    # A line I couldn't get rid of ...
                    f.write("Found html jibberish. (ignoring): ")
                    f.write(soft)
                    f.write("\n")
                elif soft in knownsoft or soft.startswith("Microsoft Office Proofing") or soft.startswith("Outils de ") or soft.startswith("Microsoft Visual C++"):
                    f.write("Found Microsoft tools (ignoring): ")
                    f.write(soft)
                    f.write("\n")
                else:
                    f.write("Found non standard software (adding): ")
                    f.write(soft)
                    f.write("\n")
                    if soft not in foundsoft:
                        foundsoft.append(soft)

    if len(foundsoft)>0:
        print("Found non standard software -",pcname)
        for soft in sorted(foundsoft):
            print(soft)
        print()

    return foundsoft


# Opening shelves
s_checkedpcs = shelve.open(db_checkedpcs)
s_foundsoft = shelve.open(db_foundsoft)
s_allsoft = shelve.open(db_allsoft)
s_knownsoft = shelve.open(db_knownsoft)

# Getting data
knownsoft = readFile(file=known_file)
knownsoft.sort()
PCS = readFile(file=check_file)
PCS.sort()

# Adding our known standard software in the database
for soft in knownsoft:
    if soft not in s_knownsoft:
        s_knownsoft[soft] = True
     
for pc in PCS:
    if pc in s_checkedpcs:
        pass
    else:
        print(pc,"has not been checked. Adding to database")
        foundsoft = findSoft(pcname=pc,knownsoft=knownsoft)
        s_checkedpcs[pc]=True
        s_foundsoft[pc]=foundsoft
        for soft in foundsoft:
            if soft not in s_allsoft:
                s_allsoft[soft]=True

print()                

# # Printing all software lists
# for pc in s_foundsoft:
    # print(pc)
    # print(s_foundsoft[pc])

# Deleting knownsoft from all software database. Known soft changes constantly
for soft in s_knownsoft:
    if soft in s_allsoft:
        del s_allsoft[soft]
        
# Finding newly added knownsoft in our already checked database
print("Checking database for changes in 'Standard Software'")
for pc in s_foundsoft:
    found = False
    c = 1
    while c > 0:
        lst = sorted(s_foundsoft[pc])
        for nonstdsoft in lst:
            if nonstdsoft in s_knownsoft:
                print("Found",nonstdsoft,"... Removing it.")
                lst.remove(nonstdsoft)
                found = True
        if found == True:
            s_foundsoft[pc]=lst
            c = 1
        else:
            c = 0
        found = False

# # Printing all software lists
# for pc in s_foundsoft:
    # print(pc)
    # print(s_foundsoft[pc])

# PC = "BXA04205"
# for soft in s_foundsoft[PC]:
    # print(soft)

print()
print("All Done")
print()
    
# Closing shelves
s_checkedpcs.close()
s_foundsoft.close()
s_allsoft.close()

