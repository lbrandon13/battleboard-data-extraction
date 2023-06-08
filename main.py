#!/usr/local/bin/python
# -*- coding: utf-8 -*-

from openpyxl import load_workbook

source_folder = "d:\Work\Coding\\battleboard_data_extraction\\battleboard-data-extraction\source_files\\"

test_workbook = load_workbook(source_folder + "Wulfric_baneguard_current_v_2021.1.xlsm")

print(test_workbook.sheetnames)

