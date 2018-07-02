"""
Creates a JSON-like string of OP number : surplus key-value pairs
and stores data in a file Surpluses_<YYYY>-<MM>.txt.
Requires updated Poslovni.pregled.xlsx file (current month must contain valid data for current click goals in column M).
Surplus is a deficit of clicks if value is negative.
"""

from os import getcwd, path
from os.path import dirname
from datetime import date
from openpyxl import load_workbook

months = { 1: "JAN",
           2: "FEB",
           3: "MAR",
           4: "APR",
           5: "MAJ",
           6: "JUN",
           7: "JUL",
           8: "AVG",
           9: "SEP",
          10: "OKT",
          11: "NOV",
          12: "DEC"}
path = getcwd().replace("\\", "\\\\") + "\\\\"
parent_dir = dirname(getcwd()).replace("\\", "\\\\") + "\\\\"
today = date.today()
year = str(today.year)[2:]
month = months[today.month]
packages = {49: 200, 99: 400, 199: 800, 399: 1600}
wb = load_workbook("T:\\SkrbZaUporabnike\\POO\\Poslovni paketi\\Poslovni-pregled.xlsx", data_only = True)
sheet = wb[month + " " + year]
surpluses_string = "{" + "// surplus data for {}-{}\n".format(today.year, today.month if today.month >= 10 else "0" + str(today.month))
for i in range(1, len(sheet["A"]), 2): # starts with 2nd row to omit 1st row and jumps by 2 since each company takes 2 rows
    if sheet["A"][i].value and sheet["E"][i].value in (49, 99, 199, 399):
        company = sheet["A"][i].value
        OP = sheet["C"][i].value
        surplus = packages[sheet["E"][i].value] - sheet["M"][i].value # clicks deficit if negative
        surpluses_string += "'{} - {}': {},\n".format(OP, company, surplus)
    elif sheet["A"][i].value: # stores total clicks goal - possible clicks from previous months as new goal for custom Poslovni paket/Preusmeritve
        company = sheet["A"][i].value
        OP = sheet["C"][i].value
        remaining = sheet["M"][i].value # goal met if <= 0
        surpluses_string += "'{}P - {}': {},\n".format(OP, company, remaining)
surpluses_string = surpluses_string[:-2] + "\n};"
open(parent_dir + "Archive/Surpluses_{}-{}.txt".format(today.year, today.month if today.month >= 10 else "0" + str(today.month)), "w", encoding="UTF-8").write(surpluses_string)
print("Clicks surplus data saved as {}Archive/Surpluses_{}-{}.txt".format(parent_dir.replace("\\\\", "/"), today.year, today.month if today.month >= 10 else "0" + str(today.month)))