from openpyxl import load_workbook

workbook = load_workbook("D:\Work\Coding\battleboard_data_extraction\Wulfric_baneguard_current_v_2021.1.xlsm")

print(workbook.sheetnames)

