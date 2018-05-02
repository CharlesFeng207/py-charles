# coding:utf8

from datetime import datetime
import json
import shutil
import sys
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter


def find_tracking_row_objs():

    src_data_start_row = int(global_json_data["src_data_start_row"])

    time_range_cols = table_json_data["time_range_cols"].split(",")

    column_letters = map(lambda i: get_column_letter(i),
                         range(1, sheet_src.max_column + 1))
    column_letters_id_dic = {k: src_letter_to_id(k) for k in column_letters}

    objs = []

    for row in range(src_data_start_row, sheet_src.max_row + 1):
        # will select this row if any of column achieve condition
        for col in time_range_cols:
            cell_name = "{}{}".format(col, row)
            cell_value = sheet_src[cell_name].value

            if cell_value == None:
                continue

        # print cell_value, type(cell_value)

            if check_time_range(cell_value):
                cell_obj = {}  # collect all data as a diectionary
                for column_letter in column_letters:  # each column as a key
                    column_id = column_letters_id_dic[column_letter]

                    if column_id == "None":
                        continue

                    cell_obj[column_id] = sheet_src["{}{}".format(column_letter, row)].value

                    objs.append(cell_obj)
                    print "row {} selected ({})".format(row, cell_value)
                    break  # avoid duplicated selecting

    return objs

def check_time_range(cell_value):
    
    if type(cell_value) is unicode:
        t1 = map(lambda x: datetime.strptime(x, "%Y/%m/%d"), cell_value.split())
        t2 = map(lambda x: check_time_range(x), t1)
        t3 = any(t2)
        return t3

    if type(cell_value) is datetime:
        if cell_value >= time_start and cell_value <= time_end:
            return True

    return False

def process_working_table():

    working_table_name = table_json_data["working_table_name"]
   
    if working_table_name == u'':
        print "working_table_name is null"
        return
    
    working_output_path = "{}{}({} to {}).xlsx".format(user_output_folder, working_table_name, to_str(time_start), to_str(time_end))

    shutil.copy(table_json_data["temp_working"], working_output_path)

    combine_number_id = table_json_data["work_combine_number_id"]
    combine_id = src_letter_to_id(table_json_data["work_combine_col"])

    combined_obj_dic = {} # { id_value : { every_key_in_src : [every_values] } }

    for obj in src_selected_objs:
        
        # just select col which isn't filled with data
        check_value = obj[src_letter_to_id(table_json_data["delay_check_col"])]
        if check_value != None:
            continue

        if combine_id not in obj:
            continue

        combine_id_value = obj[combine_id]

        if combine_id_value not in combined_obj_dic:
            combined_obj_dic[combine_id_value] = {combine_number_id:0}
            
        key_list_obj = combined_obj_dic[combine_id_value]
        key_list_obj[combine_number_id] += 1

        for k in obj:
            if k not in key_list_obj:
                key_list_obj[k] = []

            if obj[k] not in key_list_obj[k]:
                key_list_obj[k].append(obj[k])
    
    combined_objs = map(lambda x:{k : x[k] for k in x}, combined_obj_dic.values())

    attach_number_col(combined_objs)

    make_table(working_output_path, int(table_json_data["temp_working_id_row"]), int(table_json_data["temp_working_data_start_row"]), combined_objs)

def process_delay_table():
    delay_table_name = table_json_data["delay_table_name"]

    if delay_table_name == u'':
        print "delay_table_name is null"
        return

    delay_output_path = "{}{}({} to {}).xlsx".format(user_output_folder, delay_table_name, to_str(time_start), to_str(time_end))

    shutil.copy(table_json_data["temp_delay"], delay_output_path)

    delay_combine_number_id = table_json_data["delay_combine_number_id"]
    delay_combine_id = src_letter_to_id(table_json_data["delay_combine_col"])

    delay_obj_dic = {} # { id_value : { every_key_in_src : [every_values] } }

    for obj in src_selected_objs:
        
        # just select col which isn't filled with data
        check_value = obj[src_letter_to_id(table_json_data["delay_check_col"])]
        if check_value != None:
            continue

        if delay_combine_id not in obj:
            continue

        delay_combine_id_value = obj[delay_combine_id]

        if delay_combine_id_value not in delay_obj_dic:
            delay_obj_dic[delay_combine_id_value] = {delay_combine_number_id:0}
            
        key_list_obj = delay_obj_dic[delay_combine_id_value]
        key_list_obj[delay_combine_number_id] += 1

        for k in obj:
            if k not in key_list_obj:
                key_list_obj[k] = []

            if obj[k] not in key_list_obj[k]:
                key_list_obj[k].append(obj[k])
    
    delay_objs = map(lambda x:{k : x[k] for k in x}, delay_obj_dic.values())

    attach_number_col(delay_objs)

    make_table(delay_output_path, int(table_json_data["temp_delay_id_row"]), int(table_json_data["temp_delay_data_start_row"]), delay_objs)

def src_letter_to_id(letter):
    return sheet_src["{}{}".format(letter, src_id_row)].value

# add a col to indicate each data order automatically
def attach_number_col(objs):
    for i, obj in enumerate(objs):
        obj["No."] = i + 1

def to_str(obj):
    if obj == None:
        return ""

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

if len(sys.argv) == 4:
    
    print "sys.argv:", sys.argv

    global_json_path = sys.argv[1]
    table_json_path = sys.argv[2]
    user_output_folder = sys.argv[3]

    # init json config
    with open(global_json_path) as global_json_file:
        print "global json config loaded: ", global_json_path

        global_json_str = global_json_file.read().replace("\\", "\\\\")
        print global_json_str, type(global_json_str)

        global_json_data = json.loads(global_json_str)

        with open(table_json_path) as table_json_file:
            print "table json config loaded: ", table_json_path

            table_json_str = table_json_file.read().replace("\\", "\\\\")
            print table_json_str, type(table_json_str)

            table_json_data = json.loads(table_json_str)

            # read time info
            time_start_str = global_json_data["time_start"]
            time_end_str = global_json_data["time_end"]
            time_start = datetime.strptime(time_start_str, "%Y/%m/%d")
            time_end = datetime.strptime(time_end_str, "%Y/%m/%d")

            # load source excel
            path_src = global_json_data["src"]
            print "loading {}".format(path_src)

            workbook_src = load_workbook(path_src)
            sheet_src = workbook_src[global_json_data["src_sheet"]]
            src_id_row = global_json_data["src_id_row"]

            print "load complete! wait for processing ..."

            # find tracking rows by time info and save as dictionary
            src_selected_objs = find_tracking_row_objs()

            print "from {} to {}, there are {} rows selected!".format(time_start_str, time_end_str, len(src_selected_objs))

            # make working table
            process_working_table()

            # make delay table
            process_delay_table()
       

else:
    print "please input three parameters! (global json path, table json path, out put folder path)"
