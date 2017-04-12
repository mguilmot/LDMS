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

Did not really bother making it the prettiest code or the most
efficient, as we'll only need to run this every once in a while.

I was more interested in "spending a few hours to win many many 
hours of lookup the same data over and over again."