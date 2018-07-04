"""
Sums impressions and clicks data from GAdW and TSmedia (and optionally SN) and saves new data in a file 'Poslovni Paketi <YEAR>-<PREVIOUS MONTH>.csv'.
Required file for processing GAdW data: 'GAdW_CSV_<YEAR>-<PREVIOUS MONTH>.csv'.
Required file for processing TSmedia data: 'campaign-ctr-<YEAR>-<PREVIOUS MONTH>.csv'.
Optional file for exceptions: 'Exceptions_prev_<YEAR>-<PREVIOUS MONTH>'.
Removes rows from TSmedia data containing field 'custom' (Preusmeritve).
Removes rows from TSmedia data containing 0 impression and 0 clicks (assumed inactive campaigns).
"""
from os import getcwd, path
from os.path import dirname
from datetime import date
from csv import reader, writer, register_dialect, QUOTE_NONE
from json import load

def getCSVdata(file, dlm=","):
    """
    Gets data from a CSV file and returns it as a 2D list.
    :param file: str; CSV file to process
    :return: list; CSV data in a 2D list
    """
    with open(file, encoding="UTF-8") as stats:
        return list(reader(stats, delimiter=dlm))

path = getcwd().replace("\\", "\\\\") + "\\\\"
parent_dir = dirname(getcwd()).replace("\\", "\\\\") + "\\\\"
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
GAdW_prev = getCSVdata(parent_dir.replace("\\\\", "/") + "Archive/GAdW_CSV_{}-{}.csv".format(prev_month_year, prev_month_month), ";")
TSmedia_prev = [row[:7] for row in list(filter(lambda row: int(row[3]) + int(row[4]) != 0 and int(row[3]) >= int(row[4]), list(filter(lambda row: "custom" not in row[6], getCSVdata(parent_dir.replace("\\\\", "/") + "Archive/campaign-ctr-{}-{}.csv".format(prev_month_year, prev_month_month), ";")[1:]))))]
exceptions_prev = dict()
try: exceptions_prev = load(open(parent_dir.replace("\\\\", "/") + "Archive/Exceptions_prev_{}-{}.txt".format(prev_month_year, prev_month_month), encoding="UTF-8"))
except FileNotFoundError: print("No exceptions for previous month found.")
total_prev = [[row[0], row[2], row[3], row[1]] for row in GAdW_prev] # 2nd ([1]) column with account/company name added as last column
to_append = list() # campaign not in GAdW data (TSmedia only campaigns) to be appended to the output file
for row_TSmedia in TSmedia_prev:
    OP = row_TSmedia[0][:9] # assumes campaign string starts with valid OP
    found = False
    for row_total in total_prev: # iterates over rows in GAdW data
        if row_total[0] == OP: # if OP from GAdW data matches OP from TSmedia data, impressions and clicks are summed
            row_total[1] = int(row_total[1]) + int(row_TSmedia[3])
            row_total[2] = int(row_total[2]) + int(row_TSmedia[4])
            found = True
            break
    if not found: to_append.append([OP, int(row_TSmedia[3]), int(row_TSmedia[4]), row_TSmedia[0]]) # if OP from GAdW data does not matche OP from TSmedia data, TSmedia data is added to to_append list
if exceptions_prev: # adds exceptions data if exceptions found
    for OP in exceptions_prev.keys():
        for row in to_append: # iterates over rows in to append data
            if row[0] == OP: # if OP from total data matches OP from exceptions data, clicks are summed
                row[2] = row[2] + exceptions_prev[OP]
                break
for row in total_prev: row.append("insert into crm..DAILY_STATISTICS_GOOGLE_ADWORDS (campaign, clicks, impressions, datum, datum_vpisa) select '{}','{}','{}', DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0), getdate();".format(row[0], row[2], row[1]))
for row in to_append: row.append("insert into crm..DAILY_STATISTICS_GOOGLE_ADWORDS (campaign, clicks, impressions, datum, datum_vpisa) select '{}','{}','{}', DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0), getdate();".format(row[0], row[2], row[1]))
while True:
    try:
        fh = open(parent_dir + "Archive\\\\" + "Poslovni Paketi {}-{}.csv".format(prev_month_year, prev_month_month), "w", newline="")
        write = writer(fh, dialect="myDialect")
        write.writerows(total_prev)
        write.writerows(to_append)
        fh.close()
        break
    except PermissionError: input("\nPlease close Poslovni Paketi {}-{}.csv and press any key. ".format(prev_month_year, prev_month_month))

print("\nData saved in file {}Archive/Poslovni Paketi {}-{}.csv".format(parent_dir.replace("\\\\", "/"), prev_month_year, prev_month_month))
