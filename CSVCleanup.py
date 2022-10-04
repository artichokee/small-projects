# This Python script takes a CSV downloaded from the Office 365 Admin user 
# export page and cleans the data so that:
# 1. Only relevant columns Display Name and User principle name are displayed
# 2. Only users with licenses are listed

import pandas as pd
import csv
import re

# intake file name
path = input("Enter csv file path: ")

# open csv to get column names
# will be used to pop non-related names
with open(path) as csv_file:

    # create csv reader object
    csv_reader = csv.reader(csv_file, delimiter = ',')

    column_names = []

    # iterate through first row of csv to get column names
    for row in csv_reader:

        column_names.append(row)

        # only need column names, not data, so break on first row
        break

# read file in pandas (for column deletion)
sheet = pd.read_csv(path)

# Delete the weird extra characters from first column
column_names[0][0] = re.sub('ï»¿', '', column_names[0][0])

# Delete non-name and email columns
for outer_column in column_names:
    for column in outer_column:
        if (column != "Display name") and (column != "Licenses") and (column != "User principal name"):
            sheet.drop(column, inplace=True, axis=1)

# Remove all rows where unlicensed
sheet = sheet.dropna(subset=['Licenses'])

# Remove License column
sheet.drop("Licenses", inplace=True, axis=1)

# Create new CSV
csv_file_path = ""
for char in path[::-1]:
    if char == '\\':
        break
    else:
        if char == ".":
            csv_file_path = csv_file_path + ".1"
        else:
            csv_file_path = csv_file_path + char

csv_file_path = csv_file_path[::-1]
csv_path = "C:\\Users\\Sam\Downloads\\" + csv_file_path

sheet.to_csv(csv_path, index=False)




