#!/usr/bin/python

import uu
import gspread
from sys import argv, exit
from os import remove
from oauth2client.service_account import ServiceAccountCredentials


# AUTH
# -------------------------------------------------------]]
# Authentication files & configuration.
# Be sure to replace your own json file name here.
# -------------------------------------------------------]]

# Insert the path & name of your own .json auth file here.
# This is the only thing you need to edit in the script itself.
json_file = 'credentials.json'
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file, scope)

def authorize_and_get_spreadsheet(sheet_name):
    try:
        gc = gspread.authorize(credentials)
        try:
            spread = gc.open(sheet_name)
        except gspread.exceptions.SpreadsheetNotFound:
            print("Creating new spreadsheet", sheet_name)
            spread = gc.create(sheet_name)
        print("Logged into Sheets!")
    except Exception as e:
        print('Exception\n', e)
        exit("Error logging into Google Sheets. Check your authentication.")
    return spread

# Split a string into chunks so we can work around
# the Google Sheets' per-cell value limit.
def chunk_str(bigchunk, chunk_size):
    return (bigchunk[i:i + chunk_size] for i in range(0, len(bigchunk), chunk_size))


# UPLOAD
# -------------------------------------------------------]]
# File upload process for 'put'.
# -------------------------------------------------------]]
def sheetpost_put(worksheet, filename):
    wks = worksheet

    # UU-encode the source file
    uu.encode(filename, filename + ".out")

    print("Encoded file into uu format!")

    with open(filename + ".out", "r") as uploadfile:
        encoded = uploadfile.read()
    uploadfile.close()

    # Wipe the sheet of existing content.
    print("Wiping the existing data from the sheet.")
    row_sweep = 1
    column_sweep = 1
    while wks.cell(row_sweep, column_sweep).value != "":
        if row_sweep == 1000:
            row_sweep = 1
            column_sweep += 1
        wks.update_cell(row_sweep, column_sweep, "")
        print("Wipe:", row_sweep, column_sweep)
        row_sweep += 1

    # Write the chunks to Drive
    cell = 1
    column = 1
    chunk = chunk_str(encoded, 49500)

    print("Writing the chunks to the sheet. This'll take a while. Get some coffee or something.")
    for part in chunk:
        if cell == 1000:
            print("Ran out of rows, adding a column.")
            cell = 1
            column += 1
        # Add a ' to each line to avoid it being interpreted as a formula
        part = "'" + part
        wks.update_cell(cell, column, part)
        print("Write:", cell, column, "Part:", part[:20])
        cell += 1

    print("Cells used:", cell, column)

    # Delete the UU-encoded file
    remove(filename + ".out")
    print("All done! " + str(cell) + " cells filled in Sheets.")


# DOWNLOAD
# -------------------------------------------------------]]
# File download process for 'get'.
# -------------------------------------------------------]]
def sheetpost_get(worksheet, filename):

    wks = worksheet
    downfile = filename + ".uu"

    row_sweep = 1
    column_sweep = 1
    values_list = []
    values_final = []

    # Trim out the extra single quotes
    while True:
        print("Trim:", row_sweep, column_sweep)
        val = wks.cell(row_sweep, column_sweep).value

        if val == '':
            print('Break:', row_sweep, column_sweep)
            break

        values_list = wks.col_values(column_sweep)
        for value in values_list:
            if value == '':
                print("End reading")
                break

            if row_sweep > 1 and column != 1:
                # dont trim for cell,col == 1,1
                # Trim initial "'", put in there while writing
                value = value[1:]
            values_final += value
        column_sweep += 1
    values_final = "".join(values_final)

    # Save to file
    with open(downfile, "w+") as recoverfile:
        recoverfile.write(values_final)
    recoverfile.close()

    print("Saved Sheets data to decode! Decoding now. Beep boop.")
    uu.decode(downfile, filename)
    remove(downfile)
    print("Data decoded! All done!")


# HELP
# -------------------------------------------------------]]
# The help message that displays if none or too
# few arguments are given to properly execute.
# -------------------------------------------------------]]
help_message = '''To upload a sheetpost:
\t sheetpost.py put [GSheets key from URL] [Input filename]"
To retrieve a sheetpost:
\t sheetpost.py get [GSheets key from URL] [Output filename]"'''


# MAIN
# -------------------------------------------------------]]
# Where the magic happens!
# -------------------------------------------------------]]

if __name__ == '__main__':
    file_key = '1MT6l2bMJimjRdtqmZUDs3kOKxoOi3Wca8uB7C11SLDc'
    filename = 'XCHG.jpg'
    filename = 'learn_py.pdf'

    spreadsheet = authorize_and_get_spreadsheet(filename + '_sheet')
    worksheet = spreadsheet.sheet1

    print("BEGIN FILE PUT")
    sheetpost_put(worksheet, filename)
    print("END FILE PUT")

    print("BEGIN FILE GET")
    sheetpost_get(worksheet, filename.replace('.', '_retrieved.'))
    print("END FILE GET")
    exit('End main')


if argv[0] == 'python':
    # if windows user
    argv = argv[1:]

if len(argv) < 4:
    print("Too few arguments!")
    exit(help_message)

sheet_id = str(argv[2])
filename = str(argv[3])

if argv[1] == "put":
    sheetpost_put(sheet_id, filename)

elif argv[1] == "get":
    sheetpost_get(sheet_id, filename)

else:
    print("Unknown operation (accepts either 'get' or 'put')")
    exit(help_message)
print("End of program")


