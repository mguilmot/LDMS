Author: Mike Guilmot - http://www.mikeguilmot.be

Specific scripts I wrote to make our task easier 
to replace PCs and laptops.

It checks the software used on a web interface, in a specific 
table in the html. Searches for all software installed on that 
PC (info on that webpage).
Uses that data to fill the databases.

Generates a shelve 'database' with "non-standard" software.
Generates 1 excel file per PC.

Uses:
- Python3
- requests
- requests_ntlm
- csv
- shelve
- bs4

Uses beautifulsoup to make things real easy.

Before use, create following directories:
- DB
- EXCEL
- LOG
