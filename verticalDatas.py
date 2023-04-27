import pandas as pd
import os
import csv

# set the directory where the Excel files are located
dir_path = "D:\\PythonAutomation\\PythonAutomation\\Excel Docs"

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
dir_path = "D:\\PythonAutomation\\PythonAutomation\\Excel Docs"

# create a new CSV file to store the filtered data
with open("Vertical_filtered1.csv", "w", newline='') as outfile:
    writer = csv.writer(outfile)
    # loop through each CSV file in the directory
    for file_name in os.listdir(dir_path):
        if file_name.endswith(".csv") and "Buffer 1" in file_name:
            # get the name of the sheet from the CSV file name
            sheet_name = file_name.replace(".csv", "")
            # read the CSV file into a list of lists
            with open(os.path.join(dir_path, file_name), "r") as infile:
                reader = csv.reader(infile)
                data = [row for row in reader]
            # loop through the rows in the list and write the filtered data to the output file
            for i in range(27, 35):
                row_data = [sheet_name, data[i][0], data[i][1]]
                writer.writerow(row_data)
            # write an empty row between each sheet's data
            writer.writerow([])

with open("Vertical_filtered2.csv", "w", newline='') as outfile:
    writer = csv.writer(outfile)
    # loop through each CSV file in the directory
    for file_name in os.listdir(dir_path):
        if file_name.endswith(".csv") and "Buffer 2" in file_name:
            # get the name of the sheet from the CSV file name
            sheet_name = file_name.replace(".csv", "")
            # read the CSV file into a list of lists
            with open(os.path.join(dir_path, file_name), "r") as infile:
                reader = csv.reader(infile)
                data = [row for row in reader]
            # loop through the rows in the list and write the filtered data to the output file
            for i in range(27, 35):
                row_data = [sheet_name, data[i][0], data[i][1]]
                writer.writerow(row_data)
            # write an empty row between each sheet's data
            writer.writerow([])

with open("Vertical_filtered3.csv", "w", newline='') as outfile:
    writer = csv.writer(outfile)
    # loop through each CSV file in the directory
    for file_name in os.listdir(dir_path):
        if file_name.endswith(".csv") and "Buffer 3" in file_name:
            # get the name of the sheet from the CSV file name
            sheet_name = file_name.replace(".csv", "")
            # read the CSV file into a list of lists
            with open(os.path.join(dir_path, file_name), "r") as infile:
                reader = csv.reader(infile)
                data = [row for row in reader]
            # loop through the rows in the list and write the filtered data to the output file
            for i in range(27, 35):
                row_data = [sheet_name, data[i][0], data[i][1]]
                writer.writerow(row_data)
            # write an empty row between each sheet's data
            writer.writerow([])

with open("Vertical_filtered4.csv", "w", newline='') as outfile:
    writer = csv.writer(outfile)
    # loop through each CSV file in the directory
    for file_name in os.listdir(dir_path):
        if file_name.endswith(".csv") and "Buffer 4" in file_name:
            # get the name of the sheet from the CSV file name
            sheet_name = file_name.replace(".csv", "")
            # read the CSV file into a list of lists
            with open(os.path.join(dir_path, file_name), "r") as infile:
                reader = csv.reader(infile)
                data = [row for row in reader]
            # loop through the rows in the list and write the filtered data to the output file
            for i in range(27, 35):
                row_data = [sheet_name, data[i][0], data[i][1]]
                writer.writerow(row_data)
            # write an empty row between each sheet's data
            writer.writerow([])

with open("Vertical_filtered_all.csv", "w", newline='') as outfile:
    writer = csv.writer(outfile)
    # loop through each CSV file in the directory
    for file_name in os.listdir(dir_path):
        if file_name.endswith(".csv"):
            # get the name of the sheet from the CSV file name
            sheet_name = file_name.replace(".csv", "")
            # read the CSV file into a list of lists
            with open(os.path.join(dir_path, file_name), "r") as infile:
                reader = csv.reader(infile)
                data = [row for row in reader]
            # loop through the rows in the list and write the filtered data to the output file
            for i in range(27, 35):
                row_data = [sheet_name, data[i][0], data[i][1]]
                writer.writerow(row_data)
            # write an empty row between each sheet's data
            writer.writerow([])