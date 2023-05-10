# esc_parser.py - Download the results of ESC voting from a google sheet,
# count the scores and upload the results to a separate sheet.

import re
import gspread

SPREADSHEET_NAME = "ESC 2023"
gs = gspread.oauth()

# Open the spreadsheet and load all the rows in as a list of dictionaries.
sh = gs.open(SPREADSHEET_NAME)
worksheet = sh.worksheet("Form Responses")
votes = worksheet.get_all_records()
results = {}

performer_regex = re.compile(r"\d+\) (.+) \[(.+)\]")

# Sum the results into a dictionary of dictionaries (keys of first are
# the performers, the keys of the second are the categories)
for vote in votes:
    for k, v in vote.items():
        # Ignore the Ocenjevalec and Timestamp data
        if k == "Ocenjevalec:" or k == "Timestamp":
            continue
        mo = performer_regex.search(k)
        performer = mo[1]
        category = mo[2]
        results.setdefault(performer, {})
        results[performer].setdefault(category, 0)
        results[performer][category] += v

result_sheet = sh.add_worksheet(title="Results", rows=100, cols=20)

# Update the results sheet with the data from the dictionary.
# The data needs to be assembled into lists in order to be used in the update
# function.
i = 1
for performer, scores in results.items():
    row_data, header_data = [], []
    row_data.append(i)
    row_data.append(performer)
    for category, score in scores.items():
        row_data.append(score)
        header_data.append(category)
    print(row_data)
    # The range which needs to be updated is calculated from the length of the
    # list of data to be updated. For a length of 8, at the first row, the
    # result is 'A2:H2'
    row_range = f"A{i+1}:" + chr(ord("@") + len(row_data)) + str(i + 1)
    result_sheet.update(row_range, [row_data])
    i += 1

# Populate the header row
header_data.insert(0, "No.")
header_data.insert(1, "Izvajalec")
row_range = "A1:" + chr(ord("@") + len(header_data)) + "1"
result_sheet.update(row_range, [header_data])
result_sheet.format(row_range, {"textFormat": {"bold": True}})
