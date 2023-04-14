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

# create a dictionary to store the data from each CSV file
data_dict = {}

# loop through each CSV file in the directory
for file_name in os.listdir(dir_path):
    if file_name.endswith(".csv"):
        # get the name of the sheet from the CSV file name
        sheet_name = file_name.replace(".csv", "").split("_")[-1]
        # read the CSV file into a list of lists
        with open(os.path.join(dir_path, file_name), "r") as infile:
            reader = csv.reader(infile)
            data = [row for row in reader]
        # loop through the rows in the list and add the data to the dictionary
        data_dict[file_name.replace(".csv", "")] = [(data[i][0], data[i][1]) for i in range(27, 35)]
print(data_dict.items())

# create a list of frequencies
freq_list = [data_dict[key][i][0] for key in data_dict for i in range(len(data_dict[key]))]
freq_list = list(set(freq_list))
freq_list.sort()
print(freq_list)

# create a new CSV file to store the filtered data
with open("filtered_data5.csv", "w", newline='') as outfile:
    writer = csv.writer(outfile)
    # write the headers row
    headers = ["Freq"] + [file_name.replace(".csv", "") for file_name in os.listdir(dir_path) if file_name.endswith(".csv")]
    writer.writerow(headers[1:])
    print(headers)
    # loop through the frequency values and write the data to the output file
    for freq in freq_list:
        row_data = [freq]
        for file_name in headers[1:]:
            for data in data_dict[file_name]:
                if data[0] == freq:
                    row_data.append(data[1])
                    break
            else:
                row_data.append('')
        writer.writerow(row_data)
#Completed