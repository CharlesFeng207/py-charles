# coding:utf8

import datetime
import json
import shutil
import sys
import time

from openpyxl import Workbook
from openpyxl.compat import range
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter

def find_tracking_rows():
    lst = []

    src_data_start_row = int(json_data["src_data_start_row"])

    for row in range(src_data_start_row, sheet_src.max_row+1):
            for col in time_range_cols:
                cell_name = "{}{}".format(col, row)
                cell_value = str(sheet_src[cell_name].value)

                if cell_value == "None":
                    continue

                try:
                    cell_time = time.strptime(
                        str(sheet_src[cell_name].value), '%Y-%m-%d %H:%M:%S')
                except:
                    print cell_name, "is not a valid time!"

                if cell_time >= time_start and cell_time <= time_end:
                   lst.append(row)
    
    return lst

if len(sys.argv) == 3:
    json_path = sys.argv[1]

    with open(json_path) as json_file:

        json_str = json_file.read().replace("\\", "\\\\")
        json_data = json.loads(json_str)

        time_start = time.strptime(json_data["time_start"], "%Y/%m/%d")
        time_end = time.strptime(json_data["time_end"], "%Y/%m/%d")
        time_range_cols = json_data["time_range_cols"].split(",")

        workbook_src = load_workbook(json_data["src"])
        sheet_src = workbook_src[json_data["src_sheet"]]

        selected_rows = find_tracking_rows()
        print selected_rows

        working_output_path = "{}{}({} to {}).xlsx".format(sys.argv[2], json_data["working_table_name"], time.strftime(
            "%Y-%m-%d", time_start), time.strftime("%Y-%m-%d", time_end))

        shutil.copy(json_data["temp_working"], working_output_path)

        workbook_working = load_workbook(working_output_path)

        delay_output_path = "{}{}({} to {}).xlsx".format(sys.argv[2], json_data["delay_table_name"], time.strftime(
            "%Y-%m-%d", time_start), time.strftime("%Y-%m-%d", time_end))

        shutil.copy(json_data["temp_delay"], delay_output_path)

        workbook_delay = load_workbook(delay_output_path)


else:
    print "please input two parameters!"


