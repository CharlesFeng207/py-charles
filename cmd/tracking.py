# coding:utf8

from datetime import datetime
import json
import shutil
import sys
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter

def find_tracking_row_objs():

    src_data_start_row = int(json_data["src_data_start_row"])
  
    time_range_cols = json_data["time_range_cols"].split(",")

    column_letters = map(lambda i: get_column_letter(i), range(1, sheet_src.max_column + 1))
    column_letters_id_dic = { k:sheet_src["{}{}".format(k, src_id_row)].value for k in column_letters}
    
    objs = []

    for row in range(src_data_start_row, sheet_src.max_row + 1):
        # will select this row if any of column achieve condition
        for col in time_range_cols: 
            cell_name = "{}{}".format(col, row)
            cell_value = sheet_src[cell_name].value

            if cell_value == None:
                continue

            try:
                cell_time = datetime.strptime(str(cell_value), '%Y-%m-%d %H:%M:%S')
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

    working_output_path = "{}{}({} to {}).xlsx".format(sys.argv[2], json_data["working_table_name"], to_str(time_start), to_str(time_end))

    shutil.copy(json_data["temp_working"], working_output_path)

    make_table(working_output_path, int(json_data["temp_working_id_row"]), int(json_data["temp_working_data_start_row"]), src_selected_objs)

def process_delay_table():
    
    delay_output_path = "{}{}({} to {}).xlsx".format(sys.argv[2], json_data["delay_table_name"], to_str(time_start), to_str(time_end))

    shutil.copy(json_data["temp_delay"], delay_output_path)

    delay_combine_number_id = json_data["delay_combine_number_id"]
    delay_combine_id = sheet_src["{}{}".format(json_data["delay_combine_col"], src_id_row)].value

    delay_obj_dic = {} # { id_value : { every_key_in_src : [every_values] } }

    for obj in src_selected_objs:
        if delay_combine_id not in obj:
            continue

        delay_combine_id_value = obj[delay_combine_id]

        if delay_combine_id_value not in delay_obj_dic:
            delay_obj_dic[delay_combine_id_value] = {}
        
        key_list_obj = delay_obj_dic[delay_combine_id_value]

        for k in obj:
            if k not in key_list_obj:
                key_list_obj[k] = []

            if obj[k] not in key_list_obj[k]:
                key_list_obj[k].append(obj[k])
    
    delay_objs = map(lambda x:{k : x[k] for k in x}, delay_obj_dic.values())

    make_table(delay_output_path, int(json_data["temp_delay_id_row"]), int(json_data["temp_delay_data_start_row"]), delay_objs)

def to_str(obj):
    if type(obj) is str:
        return obj

    if type(obj) is unicode:
        return obj.encode('UTF-8')

    if type(obj) is datetime:
        return obj.strftime("%Y-%m-%d")

    if type(obj) is list:
        return ','.join(map(to_str, obj))

    return str(obj)

def make_table(output_path, id_row, writing_row, target_objs):
    
    workbook_target = load_workbook(output_path)
    sheet_target = workbook_target.active

    column_letters = map(lambda i: get_column_letter(i), range(1, sheet_target.max_column + 1))
    column_letters_id_dic = { k:sheet_target["{}{}".format(k, id_row)].value for k in column_letters}

    for obj in target_objs:
        for letter in column_letters_id_dic.keys():
            cell_id = column_letters_id_dic[letter]
            if cell_id in obj:
                sheet_target["{}{}".format(letter, writing_row)].value = to_str(obj[cell_id])
           
        writing_row += 1
    
    workbook_target.save(output_path)

    print "{} created!".format(output_path)

if len(sys.argv) == 3:
    json_path = sys.argv[1]
    
    with open(json_path) as json_file:

        # init json config
        json_str = json_file.read().replace("\\", "\\\\")
        json_data = json.loads(json_str)

        # read time info
        time_start_str = json_data["time_start"]
        time_end_str = json_data["time_end"]
        time_start = datetime.strptime(time_start_str, "%Y/%m/%d")
        time_end = datetime.strptime(time_end_str, "%Y/%m/%d")

        # load source excel
        path_src = json_data["src"]
        print "loading {}".format(path_src)

        workbook_src = load_workbook(json_data["src"])
        sheet_src = workbook_src[json_data["src_sheet"]]
        src_id_row = json_data["src_id_row"]

        print "load complete! wait for processing ..."

        # find tracking rows by time info and save as dictionary
        src_selected_objs = find_tracking_row_objs()

        print "from {} to {}, there are {} rows selected!".format(time_start_str, time_end_str, len(src_selected_objs))

        # make working table
        process_working_table()

        # make delay table
        process_delay_table()

else:
    print "please input two parameters!"
