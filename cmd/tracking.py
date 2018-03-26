# coding:utf8

import datetime
import json
import shutil
import sys
import time
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter

def find_tracking_row_objs():

    src_data_start_row = int(json_data["src_data_start_row"])
    src_id_row = json_data["src_id_row"]
    time_range_cols = json_data["time_range_cols"].split(",")

    column_letters = map(lambda i: get_column_letter(i), range(1, sheet_src.max_column + 1))
    column_letters_id_dic = { k:sheet_src["{}{}".format(k, src_id_row)].value for k in column_letters}
    
    objs = []

    for row in range(src_data_start_row, sheet_src.max_row + 1):
        # will select this row if any of column achieve condition
        for col in time_range_cols: 
            cell_name = "{}{}".format(col, row)
            cell_value = str(sheet_src[cell_name].value)

            if cell_value == "None":
                continue

            try:
                cell_time = time.strptime(cell_value, '%Y-%m-%d %H:%M:%S')
            except:
                print cell_name, "is not a valid time!"

            if cell_time >= time_start and cell_time <= time_end:
                cell_obj = {} # collect all data as a diectionary
                for column_letter in column_letters: # each column as a key
                    column_id = column_letters_id_dic[column_letter]

                    if column_id == "None":
                        continue

                    cell_obj[column_id] = sheet_src["{}{}".format(column_letter, row)].value

                objs.append(cell_obj)
                print "row {} selected ({})".format(row, cell_value)
                break # avoid duplicated selecting

    return objs

def process_working_table():

    working_output_path = "{}{}({} to {}).xlsx".format(sys.argv[2], json_data["working_table_name"], time.strftime(
        "%Y-%m-%d", time_start), time.strftime("%Y-%m-%d", time_end))

    shutil.copy(json_data["temp_working"], working_output_path)
    workbook_working = load_workbook(working_output_path)
    sheet_working = workbook_working.active

    writing_row_index = int(json_data["temp_working_id_row"])

    column_letters = map(lambda i: get_column_letter(i), range(1, sheet_working.max_column + 1))
    column_letters_id_dic = { k:sheet_working["{}{}".format(k, writing_row_index)].value for k in column_letters}

    temp_working_data_start_row = int(json_data["temp_working_data_start_row"])
    for obj in selected_cell_objs:
        for letter in column_letters_id_dic.keys():
            cell_id = column_letters_id_dic[letter]
            if cell_id in obj:
                sheet_working["{}{}".format(letter, temp_working_data_start_row)].value = obj[cell_id]
           
        temp_working_data_start_row += 1
    
    workbook_working.save(working_output_path)

    print "{} created!".format(working_output_path)

def process_delay_table():
    
    delay_output_path = "{}{}({} to {}).xlsx".format(sys.argv[2], json_data["delay_table_name"], time.strftime(
        "%Y-%m-%d", time_start), time.strftime("%Y-%m-%d", time_end))

    shutil.copy(json_data["temp_delay"], delay_output_path)
    sheet_delay = load_workbook(delay_output_path).active

if len(sys.argv) == 3:
    json_path = sys.argv[1]
    
    with open(json_path) as json_file:

        # init json config
        json_str = json_file.read().replace("\\", "\\\\")
        json_data = json.loads(json_str)

        # read time info
        time_start_str = json_data["time_start"]
        time_end_str = json_data["time_end"]
        time_start = time.strptime(time_start_str, "%Y/%m/%d")
        time_end = time.strptime(time_end_str, "%Y/%m/%d")

        # load source excel
        path_src = json_data["src"]
        print "loading {}".format(path_src)

        workbook_src = load_workbook(json_data["src"])
        sheet_src = workbook_src[json_data["src_sheet"]]

        print "load complete! wait for processing ..."

        # find tracking rows by time info and save as dictionary
        selected_cell_objs = find_tracking_row_objs()

        print "from {} to {}, there are {} rows selected!".format(time_start_str, time_end_str, len(selected_cell_objs))

        process_working_table()

        process_delay_table()

else:
    print "please input two parameters!"
