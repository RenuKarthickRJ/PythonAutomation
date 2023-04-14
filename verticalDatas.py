import pandas as pd
import os
import csv

# set the directory where the Excel files are located
dir_path = "File_path_here"

# loop through each Excel file in the directory and convert to CSV
for file_name in os.listdir(dir_path):
    if file_name.endswith(".xlsx"):
        # read the Excel file into a Pandas DataFrame
        xl_file = pd.read_excel(os.path.join(dir_path, file_name), sheet_name=None)
        # loop through each sheet and save as a separate CSV file
        for sheet_name in xl_file:
            df = xl_file[sheet_name]
            # save the CSV file with the name of the original Excel file and sheet name
            csv_name = file_name.replace(".xlsx", "") + "_" + sheet_name + ".csv"
            df.to_csv(os.path.join(dir_path, csv_name), index=False)

# set the directory where the CSV files are located
dir_path = "File_path_here"

# create a new CSV file to store the filtered data
with open("filtered", "w", newline='') as outfile:
    writer = csv.writer(outfile)
    # loop through each CSV file in the directory
    for file_name in os.listdir(dir_path):
        if file_name.endswith(".csv"):
            # get the name of the sheet from the CSV file name
            sheet_name = file_name.replace(".csv", "").split("_")[-1]
            # read the CSV file into a list of lists
            with open(os.path.join(dir_path, file_name), "r") as infile:
                reader = csv.reader(infile)
                data = [row for row in reader]
            # loop through the rows in the list and write the filtered data to the output file
            for i in range(27, 35):
                row_data = [file_name, data[i][0], data[i][1]]
                writer.writerow(row_data)
            # write an empty row between each sheet's data
            writer.writerow([])