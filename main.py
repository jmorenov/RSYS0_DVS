import csv
import os
from typing import List

import openpyxl


def list_txts_directory() -> List[tuple[str, str]]:
    files = []

    dir_path = 'TXTS'
    for path in os.listdir(dir_path):
        if os.path.isfile(os.path.join(dir_path, path)) and path.endswith('.TXT'):
            files.append((os.path.join(dir_path, path), path))

    return files


def open_excel(workbook, txt_file_name: str):
    return workbook.create_sheet(txt_file_name)


def insert_data_in_sheet(worksheet, text) -> None:
    worksheet.append(text)


def read_txt(file_name: str) -> List[List[float]]:
    with open(file_name, 'r', newline='') as txtfile:
        txt_lines = txtfile.readlines()
        text_rows = []

        for line in txt_lines:
            if line.startswith('    ') and line.startswith('    NODE') is False:
                row_txt = line.split('    ')
                row_float = [float(x) for x in row_txt if x != '']
                text_rows.append(row_float)

        return text_rows


if __name__ == '__main__':
    workbook = openpyxl.Workbook(write_only=True)
    txts_files = list_txts_directory()

    for txt_file in txts_files:
        txt_path, txt_name = txt_file
        worksheet = open_excel(workbook, txt_name)
        txt_data = read_txt(txt_path)

        insert_data_in_sheet(worksheet, ['NODE', 'SX', 'SY', 'SZ', 'SXY', 'SYZ', 'SXZ'])

        for row in txt_data:
            insert_data_in_sheet(worksheet, row)

    workbook.save('test.xlsx')