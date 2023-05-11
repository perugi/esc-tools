# esc_parser.py - Download the results of ESC voting from a google sheet,
# count the scores and upload the results to a separate sheet.

import re, json, sys
import gspread
import argparse

parser = argparse.ArgumentParser(
    description="Parse the results from an ESC response sheet and store them as a table on a separate sheet"
)
parser.add_argument(
    "-l",
    "--lang",
    type=str,
    choices=["si", "en"],
    default="si",
    help="Language of the generated table, by default si.",
)
parser.add_argument(
    "-n",
    "--name",
    type=str,
    required=True,
    help='Name of the Google Sheet on which to operate, e.g. "ESC 2023")',
)

args = parser.parse_args()
selected_lang = args.lang
spreadsheet_name = args.name

if sys.platform == "linux":
    credentials = "~/.config/gspread/credentials.json"
elif sys.platform == "win32":
    credentials = "%APPDATA%/Roaming/gspread/credentials.json"
else:
    print("ERROR: Unrecognized OS")
    sys.exit()

gs = gspread.oauth(credentials_filename=credentials)

# lang.json contains the translation of all the language specific strings.
with open("lang.json") as f:
    lang_data = f.read()
lang = json.loads(lang_data)

# Open the spreadsheet and load all the rows in as a list of dictionaries.
sh = gs.open(spreadsheet_name)
worksheet = sh.worksheet("Form Responses")
votes = worksheet.get_all_records()
results = {}

performer_regex = re.compile(r"\d+\) (.+) \[(.+)\]")

# Sum the results into a dictionary of dictionaries (keys of first are
# the performers, the keys of the second are the categories)
for vote in votes:
    for k, v in vote.items():
        # Ignore the Ocenjevalec and Timestamp data
        if k == lang[selected_lang]["judge"] or k == lang[selected_lang]["timestamp"]:
            continue
        mo = performer_regex.search(k)
        performer = mo[1]
        category = mo[2]
        results.setdefault(performer, {})
        results[performer].setdefault(category, 0)
        results[performer][category] += v

result_sheet = sh.add_worksheet(title=lang[selected_lang]["results"], rows=100, cols=20)

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
header_data.insert(0, lang[selected_lang]["no"])
header_data.insert(1, lang[selected_lang]["performer"])
row_range = "A1:" + chr(ord("@") + len(header_data)) + "1"
result_sheet.update(row_range, [header_data])
result_sheet.format(row_range, {"textFormat": {"bold": True}})
