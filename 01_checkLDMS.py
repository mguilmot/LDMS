'''
    Check LDMS for non standard software
    uses module 
    - 'requests'        : pip install requests
    - 'urllib3'         : pip install urllib3
    - 'requests_ntlm'   : pip install requests_ntlm
    - 'shelve'          : no install needed
    - 'beautifulsoup'   : pip install beautifulsoup
    
    Files used:
    - H:\AD.txt (case sensitive!)
        -> Format:
            # Comment line
            ad_user = "AD\\CWID"        (note the double backslash!)
            ad_pwd = "AD password"      
'''

### Files we will use
pwd_file = "H:\\AD.txt"
log_file = "LOG\\checkLDMS_"
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
import shelve
import requests
from requests_ntlm import HttpNtlmAuth
import bs4

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
    data = []
    with open(file,"r") as f:
        for l in f:
            line = l.strip("\n")
            line.strip()
            if line.startswith("#") or len(line)<2:
                continue
            else:
                data.append(line)
    return data

# Starting a session to LDMS
session = requests.Session()
session.auth=auth=HttpNtlmAuth(ad_user,ad_pwd)    

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
print("Generating our known software database.")
for soft in knownsoft:
    if soft not in s_knownsoft:
        s_knownsoft[soft] = True
     
for pc in PCS:
    foundsoft = []
    if pc in s_checkedpcs:
        pass
    else:
        with open(log_file + pc + ".log","w") as f:
            print(pc,"has not been checked. Adding to database")
            url = link + pc
            data = session.get(url)
            html = data.text
            soup = bs4.BeautifulSoup(html,"html.parser")
            try:
                table = soup.find_all("table")[-1]
                trs = table.find_all('tr')
                for tr in trs:
                    tds = tr.find_all('td')
                    soft = tds[0].text
                    if soft.startswith("Security Update for") or soft.startswith("Update for Microsoft") or soft.startswith("Service Pack") or "(KB" in soft:
                        # MS updates
                        f.write("Found MS:" + soft + "\n")
                    elif soft.startswith("Lenovo"):
                        # Standard Lenovo software
                        f.write("Found Lenovo:" + soft + "\n")
                    elif soft in knownsoft or soft.startswith("Microsoft Office Proofing") or soft.startswith("Outils de ") or soft.startswith("Microsoft Visual C++"):
                        # MS proofing tools and C++
                        f.write("Found Proofing:" + soft + "\n")
                    elif soft == "suitename":
                        # Ignore table header
                        continue
                    else:
                        # Finally the software we are interested in.
                        print("Found non standard soft:",soft)
                        f.write("Found non standard software:" + soft + "\n")
                        foundsoft.append(soft)
            except:
                print(pc,"could not be checked.")
                f.write("Error: could not check PC. No LDMS contract?\n")
                foundsoft = []
        
        s_checkedpcs[pc]=True
        if len(foundsoft)==0:
            foundsoft = ["No extra software"]
        s_foundsoft[pc]=foundsoft
        for soft in foundsoft:
            if soft not in s_allsoft:
                s_allsoft[soft]=True

print()                

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

print()
print("All Done")
print()

# Closing shelves
s_checkedpcs.close()
s_foundsoft.close()
s_allsoft.close()

