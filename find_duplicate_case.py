import json
import re

from openpyxl import load_workbook, Workbook

NEW_4C_FIRST_ROW = 4
NEW_4C_ADDRESS_COL = "E"
NEW_4C_NAME_COL = "B"
NEW_4C_SHEET_NAME = "Batch 4c - other direct SW refe"

OLD_4B_FIRST_ROW = 3
OLD_4B_ADDRESS_COL = "E"
OLD_4B_NAME_COL = "B"
OLD_4B_SHEET_NAME = "Batch 4b - BedokGeylang Serai"

OLD_4A_FIRST_ROW = 3
OLD_4A_ADDRESS_COL = "E"
OLD_4A_NAME_COL = "B"
OLD_4A_SHEET_NAME = "Batch 4a - Bukit MerahKreta Aye"

OLD_3_FIRST_ROW = 3
OLD_3_ADDRESS_COL = "C"
OLD_3_NAME_COL = "E"
OLD_3_SHEET_NAME = "Batch 3"

OLD_2_FIRST_ROW = 3
OLD_2_ADDRESS_COL = "C"
OLD_2_NAME_COL = "E"
OLD_2_SHEET_NAME = "Batch 2"

def load_excel_with_address(first_row, address_col, sheet_name, filename):
    limiter_col = address_col
    if limiter_col == "C":
        limiter_col = "F"
    finished = False
    cur_row = first_row
    cur_row_str = str(cur_row)
    families = {}

    wb = load_workbook(filename)
    sheet = wb[sheet_name]

    while not finished:
        address = sheet[address_col + cur_row_str].value.strip()
        try:
            address = sheet[address_col + cur_row_str].value.strip()
            #postal_code = re.findall('\d{6}', address)[0]
            unit_number = re.findall('#\d+-\d+', address)[0]
            key = unit_number
            if key in families:
                print("CONFLICT IN SAME BATCH", key, cur_row_str, address_col, sheet_name, filename)
            else:
                families[key] = {
                    #"postal_code": postal_code,
                    "address": address,
                    "unit_number": unit_number,
                    "row": cur_row,
                    "sheet_name": sheet_name,
                    "filename": filename
                }
        except Exception as e:
            print(e, address, cur_row_str, sheet_name, filename)
        
        cur_row += 1
        cur_row_str = str(cur_row)
        if not sheet[limiter_col + cur_row_str].value:
            finished = True
    return families

def merge_dicts(dict1, dict2):
    for key in dict2:
        if key in dict1:
            print("CONFLICT", dict2[key], dict1[key])
        else:
            dict1[key] = dict2[key]
    return dict1

def get_dupes(base_data, new_data):
    for key in new_data:
        if key in base_data:
            print("DUPLICATE", new_data[key], base_data[key])


if __name__ == "__main__":
    batch2 = load_excel_with_address(OLD_2_FIRST_ROW, OLD_2_ADDRESS_COL, OLD_2_SHEET_NAME, "duplicate_case_data/Confirmed beneficiaries – enrolment_acknowledgements.xlsx")
    batch3 = load_excel_with_address(OLD_3_FIRST_ROW, OLD_3_ADDRESS_COL, OLD_3_SHEET_NAME, "duplicate_case_data/Confirmed beneficiaries – enrolment_acknowledgements.xlsx")
    batch4a = load_excel_with_address(OLD_4A_FIRST_ROW, OLD_4A_ADDRESS_COL, OLD_4A_SHEET_NAME, "duplicate_case_data/(4-X) PSS Assessment of Beneficiary Clusters.xlsx")
    batch4b = load_excel_with_address(OLD_4B_FIRST_ROW, OLD_4B_ADDRESS_COL, OLD_4B_SHEET_NAME, "duplicate_case_data/(4-X) PSS Assessment of Beneficiary Clusters.xlsx")
    batch4c = load_excel_with_address(NEW_4C_FIRST_ROW, NEW_4C_ADDRESS_COL, NEW_4C_SHEET_NAME, "duplicate_case_data/(4-X) PSS Assessment of Beneficiary Clusters.xlsx")
    
    base_data = {}
    base_data = merge_dicts(base_data, batch2)
    base_data = merge_dicts(base_data, batch3)
    base_data = merge_dicts(base_data, batch4a)
    base_data = merge_dicts(base_data, batch4b)

    get_dupes(base_data, batch4c)
    
