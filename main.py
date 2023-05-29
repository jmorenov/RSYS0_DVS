import csv
import os
from typing import List

import openpyxl


def list_csv_directory() -> List[str]:
    files = []

    dir_path = 'CSV'
    for path in os.listdir(dir_path):
        if os.path.isfile(os.path.join(dir_path, path)) and path.endswith('.csv'):
            files.append(os.path.join(dir_path, path))

    return files


def open_excel():
    workbook = openpyxl.load_workbook('020_Fatigue_Frames_Selection_Class.xlsx')
    worksheet = workbook['Spectra_Results']

    return workbook, worksheet


def insert_data_in_sheet(worksheet, cell, text) -> None:
    worksheet[cell] = text


def read_csv(file_name) -> List[str]:
    with open(file_name, newline='') as csvfile:
        csvdata = csv.reader(csvfile, delimiter=';', quotechar='|')
        csv_text = []

        for row in csvdata:
            csv_text.append(row[1])

        return csv_text


if __name__ == '__main__':
    csv_files = list_csv_directory()
    workbook, worksheet = open_excel()
    index = 3

    for csv_file in csv_files:
        csv_data = read_csv(csv_file)
        row = 0
        for column in ['B', 'C', 'D', 'E', 'F']:
            insert_data_in_sheet(worksheet, column + str(index), csv_data[row])
            row += 1

        index += 1

    workbook.save('test.xlsx')