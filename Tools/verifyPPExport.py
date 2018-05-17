"""
Verifies export of data on GAdW and TSmedia networks in previous month against data in Poslovni-pregled.xlsx table and outputs results.
Requires Poslovni-pregled.xlsx table with clicks data for previous month and Poslovni Paketi <YYYY>-<MM>.csv exported data for previous month file.
Outputs discrepancies in clicks and match percentage for each account/campaign in Poslovni Paketi <YYYY>-<MM>.csv file.
"""

from datetime import date
from os import getcwd, path, pardir
from csv import reader, writer, register_dialect, QUOTE_NONE
from openpyxl import load_workbook

# functions
def getCSVdata(file, dlm=","):
    """
    Gets data from a CSV file and returns it as a 2D list.
    :param file: str; CSV file to process
    :return: list; CSV data in a 2D list
    """
    with open(file) as stats: return list(reader(stats, delimiter=dlm))

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
path_cwd = getcwd().replace("\\", "\\\\") + "\\\\"
path_parent = path.abspath(path.join(path_cwd, pardir)).replace("\\", "\\\\") + "\\\\"
today = date.today()
prev_month_year = today.year
prev_month_month = today.month - 1
if prev_month_month == 0:
    prev_month_year -= 1
    prev_month_month = 12
prev_month_month = str(prev_month_month)
if len(prev_month_month) == 1: prev_month_month = "0" + prev_month_month
prev_month_year = str(prev_month_year)
register_dialect('myDialect', delimiter=';', quoting=QUOTE_NONE)
data_prev = getCSVdata(path_parent + "Archive/Poslovni paketi {}-{}.csv".format(prev_month_year, prev_month_month), ";")
wb = load_workbook(path_parent + "Poslovni-pregled.xlsx", data_only = True)
sheet = wb[months[int(prev_month_month)] + " " + prev_month_year[2:]]
for campaign in data_prev:
    for i in range(1, len(sheet["A"]), 2):
        if sheet["A"][i].value and sheet["C"][i].value == campaign[0]:
            difference = int(campaign[2]) - sheet["Q"][i].value
            matching_percentage = str(round(abs(sheet["Q"][i].value / int(campaign[2])) * 100, 2)) + "%"
            print(campaign[0], campaign[3], difference, matching_percentage)
            campaign.extend([difference, matching_percentage])
while True:
    try:
        fh = open(path_parent + "Archive/Poslovni Paketi Discrepances {}-{}.csv".format(prev_month_year, prev_month_month), "w", newline="")
        write = writer(fh, dialect="myDialect")
        write.writerows(data_prev)
        fh.close()
        break
    except PermissionError: input("\nPlease close Poslovni Paketi Discrepances {}-{}.csv and press any key. ".format(prev_month_year, prev_month_month))

print("\nData saved in file {}Archive/Poslovni Paketi Discrepances {}-{}.csv".format(path_parent.replace("\\\\", "/"), prev_month_year, prev_month_month))
