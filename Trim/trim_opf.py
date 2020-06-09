import os
import pyvyu as pv
import csv

# The following script will parse all opf files found in the source_folder; trim and shift all columns and cells 
# according to the onset and offset of the reference cell.
# The script will find the column_ref and cell_ref onset and offset and will trim all columns found in the columns list 
# accordingly, if not found the column will be removed from the result.

# Important Note: pyvyu creates temporary files when saving the the opf spreadsheet, this is glitch in pyvyu and will be fixed in future 
# releases, please ignore them

#  The reference column where to get the onset and offset to be used to trim the opf file
column_ref = "COLUMN_NAME"

# The cell index where to find the onset and offset, keep in mind that it starts from 0 (means first cell)
# if you would like to have the second cell of the reference column you should put 1 in cell_ref
cell_ref = 0

# Put the columns that you would like to keep here!
columns = ["COLUMN_NAME_1", "COLUMN_NAME_1", "..."]

# Path to your OPF files folder
source_folder = "SOURCE/FOLDER/PATH"

# Path to the folder where generated opf files will be saved
# make sure that the folder exists and is not a sub directory of
# the source_folder
target_folder = "TARGET/FOLDER/PATH"

#  This just to keep a record of all opf files processed, no need to be changed
fields = ["opf_original", "opf_trimmed", "onset_timestamp", "offset_timestamp", "onset_millis", "offset_millis"]
rows = []

# change the prefix of the created opf files
file_cut = "{}_cut.opf"


for root, dirs, files in os.walk(source_folder):
    for file in files:
        if file.endswith(".opf"):
            print("Loading file: {}".format(file))
            original_file = os.path.join(root, file)
            sheet = pv.load_opf(original_file)

            col = sheet.get_column(column_ref)
            if col is None:
                continue

            onset = col.cells[cell_ref].onset
            offset = col.cells[cell_ref].offset
            print("Found onset: {} offset: {} in {}".format(pv.to_timestamp(onset), pv.to_timestamp(offset), column_ref))

            if len(columns) != 0:
                sheet.columns = {colname: col for (colname, col) in sheet.columns.items() if colname in columns}

            sheet = pv.trim_sheet(onset, offset, sheet)
            file_cut = os.path.join(target_folder, os.path.splitext(file)[0] + '_cut.opf')
            # save opf generate tmp files, please ignore them it will be fixed in future pyvyu releases
            pv.save_opf(sheet, file_cut)
            rows.append([original_file, file_cut, pv.to_timestamp(onset), pv.to_timestamp(offset), onset, offset])

# CSV file where we keep, onsets offsets and files paths
output_file_path = os.path.join(target_folder, "{}.csv".format(column_ref))
with open(output_file_path, 'w') as file:
    writer = csv.writer(file)
    # writing the fields
    writer.writerow(fields)
    # writing the data rows
    writer.writerows(rows)
