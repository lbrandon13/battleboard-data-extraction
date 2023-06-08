#!/usr/local/bin/python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook

from textLists import skillList

source_folder = "d:\Work\Coding\\battleboard_data_extraction\\battleboard-data-extraction\source_files\\"

test_workbook = load_workbook(source_folder + "Wulfric_baneguard_current_v_2021.1.xlsm")

print(test_workbook.sheetnames)

sheet = test_workbook['The Character']

skillMap = {}

# openpyxl starts at 1 for row and column
for i in range(14,397):
    
    if sheet.cell(row=i,column=3).value != None:
        # allow for offset between row number on sheet and position in list
        try:
            skillMap[skillList[(i-14)]] = sheet.cell(row=i,column=3).value
        except Exception as e:
            print("Incorrect Number of skills, likely wrong battleboard version")
            raise e

# print(sheet.cell(row=15,column=3).value)

# print(len(skillList))
print(skillMap)

