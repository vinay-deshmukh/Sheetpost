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

WKS_RANGE = 'A1:B1000'

def authorize_and_get_spreadsheet(sheet_name):
    try:
        gc = gspread.authorize(credentials)
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
    out_file = filename + '.out'
    uu.encode(filename, out_file)

    print("Encoded file into uu format!")

    with open(out_file, "r") as uploadfile:
        encoded = uploadfile.read()
        print('Out file read')
        print(encoded[:20])

    # Wipe the sheet of existing content.
    print("Wiping the existing data from the sheet.")
    row_sweep = 1
    column_sweep = 1

    # Increse the range for larger files
    all_cells = sorted(wks.range(WKS_RANGE), key=lambda x: x.col)

    i = 0
    while True:
        val = all_cells[i].value
        if val == '':
            print("End of original file contents")
            break

        # Clear cell
        all_cells[i].value = ''
        print("Wiping:", all_cells[i].row, all_cells[i].col)
        i += 1
    total_wipes = i + 1
    # Update all the cells at once
    '''
    print('Size of value', len(all_cells[0].value))
    wks.update_acell('A1', all_cells[0].value)
    print('exit in put')
    exit('EXIT in put')
    '''

    cell_chunk = 100 # 100 cells written per call
    for i in range(0, total_wipes, cell_chunk):
        print("Wipe:", i, i + cell_chunk)
        wks.update_cells(all_cells[i: i + cell_chunk])

    # Get iterator over the file contents
    # chunk_size should be less than (50000-1)
    # since cell limit is 50000, and one character is used to prepend to data
    chunk = chunk_str(encoded, chunk_size=49500)

    print("Writing the chunks to the sheet. This'll take a while. Get some coffee or something.")
    for i, part in enumerate(chunk):
        '''if cell == 1000:
            print("Ran out of rows, adding a column.")
            cell = 1
            column += 1
        '''
        cell = all_cells[i]
        cell.value = part
        # Add a ' to each line to avoid it being interpreted as a formula
        #print('Before prepend:', repr(cell.value[:20]))
        part = "'" + part
        # wks.update_cell(cell, column, part)
        cell.value = part
        # Using repr so '\n' is shown, and not interpreted as newline
        print("Write:", cell.row, cell.col, "Part:", repr(cell.value[:20]))
        #cell += 1
    total_cells_written = i + 1

    # Update the edited cells
    cell_chunk = 100 # Update only 100 cells at a time
    for i in range(0, total_cells_written, cell_chunk):
        print("Uploading:", i, i + cell_chunk)
        first_cell, last_cell = all_cells[i], all_cells[i + cell_chunk]
        print("Cells:", first_cell.row, first_cell.col,
                        last_cell.row,  last_cell.col)
        wks.update_cells(all_cells[i: i + cell_chunk])

    print("Cells used:", cell.row, cell.col)

    # Delete the UU-encoded file
    #remove(out_file)
    print("All done! " + str(cell.row * cell.col) + " cells filled in Sheets.")


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

    # Trim and read the data
    all_cells = sorted(wks.range(WKS_RANGE), key=lambda x: x.col)
    i = 0
    while True:
        cell = all_cells[i]
        value = cell.value
        if value == '':
            break
        # Remove the prepended "'"
        value = value[1:]
        print('Trim:', cell.row, cell.col, repr(cell.value[:20]))
        values_final.append(value)

        i += 1
    values_final = ''.join(values_final)

    # Save to file
    with open(downfile, "w+") as recoverfile:
        recoverfile.write(values_final)

    print("Saved Sheets data to decode! Decoding now. Beep boop.")
    uu.decode(downfile, filename)
    #remove(downfile)
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
    filename = 'XCHG.jpg'
    filename = 'learn_py.pdf'
    #filename = 'byte_python.pdf'
    filename = 'lc16.pdf'
    filename = 'algo18.chm'
    filename = 'pjava22.pdf'

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


