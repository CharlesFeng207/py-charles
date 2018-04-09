# coding:utf8

import json
import shutil
import time
import traceback
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter
from copy import deepcopy

class PartChanging:
    
    # when a part first time to appear.
    op_init = 'init'

    # if new part is just old part, will override old part data
    op_override = 'override' 

    # if old part updated to new part, will clear usage which the new part using, 
    # after that build a relationship between the new and old
    op_update = 'update'

    # if new part id is invalid, will clear all usage for old part id
    op_delete = 'delete' 

    def __init__(self, new_part_id, old_part_id, new_part_data, old_part_data, row_number):
        self.new_part_id = new_part_id
        self.old_part_id = old_part_id
        self.old_part_data = old_part_data
        self.new_part_data = new_part_data
        self.row_number = row_number

    def __str__(self):
         return '< PartChanging id:{} {} -> {} {} row:{} >'.format(
             self.old_part_id, is_valid_part_id(self.old_part_id), self.new_part_id, is_valid_part_id(self.new_part_id), self.row_number)
    
    def do(self, parts_wrapper_objs):

        print self

        PartChanging.init_before_operation(parts_wrapper_objs, self.old_part_id, self.old_part_data, 
        '{} {} by row {}'.format(PartChanging.op_init, self.old_part_id, self.row_number))

        PartChanging.init_before_operation(parts_wrapper_objs, self.new_part_id, self.new_part_data, 
        '{} {} by row {}'.format(PartChanging.op_init, self.new_part_id, self.row_number))

        if self.new_part_id == None: # delete
            empty_data = PartDataRecord()
            empty_data.op_info = '{} {} by row {}'.format(PartChanging.op_delete, self.old_part_id, self.row_number)
            print empty_data.op_info
            parts_wrapper_objs[self.old_part_id].append_part_data_record(empty_data)

        elif self.new_part_id == self.old_part_id: # override
            modified_data = deepcopy(self.new_part_data)
            modified_data.op_info = '{} {} by row {}'.format(PartChanging.op_override, self.old_part_id, self.row_number)
            print modified_data.op_info

            parts_wrapper_objs[self.old_part_id].append_part_data_record(modified_data)
        else: # upate: override old part + init new part + build relationship

            # override updated data

            modified_data = deepcopy(self.old_part_data)
            modified_data.op_info = '{} {} by row {}, new part id: {}'.format(
                PartChanging.op_override, self.old_part_id, self.row_number, self.new_part_id)

            print modified_data.op_info

            for k in self.new_part_data.car_usage:
                if k in modified_data.car_usage and self.new_part_data.car_usage[k] > 0:
                    modified_data.car_usage[k] = 0

            old_part_state = parts_wrapper_objs[self.old_part_id]
            new_part_state = parts_wrapper_objs[self.new_part_id]

            old_part_state.append_part_data_record(modified_data)

            # build relationship
            if new_part_state not in old_part_state.next_part:
                old_part_state.next_part.append(new_part_state)

            new_part_state.pre_part = old_part_state
            
        print '\n'
           
    @staticmethod
    def init_before_operation(parts_wrapper_objs, part_id, part_data, op_info):
        if part_id != None and part_id not in parts_wrapper_objs:
            init_data = deepcopy(part_data)
            init_data.op_info = op_info
            print op_info

            t = PartWrapper(part_id)
            t.append_part_data_record(init_data)
            parts_wrapper_objs[part_id] = t

class PartWrapper:
    
    def __init__(self, part_id):
        self.pre_part = None
        self.next_part = []
        self.part_id = part_id
        self.part_data_records = []

    def __str__(self):
        return str(self.part_id)

    def get_detail_str(self):
        arr = [str(item) for item in self.part_data_records]
        arr.insert(0, str(self)) # add self
        arr_str = '\n'.join(arr)
        return arr_str

    def is_root_wrapper(self):
        return self.pre_part == None

    def is_final_wrapper(self):
        return len(self.next_part) == 0
    
    def find_root_wrapper(self):
        t = self
        while not t.is_root_wrapper():
            t = t.pre_part
        return t

    def append_part_data_record(self, part_data_record):
        self.part_data_records.append(part_data_record)

    def get_initial_part_data(self):
        if len(self.part_data_records) > 0:
            return self.part_data_records[0]
        return None
    
    def get_newest_part_data(self):
        if len(self.part_data_records) > 0:
            return self.part_data_records[-1]
        return None

class PartDataRecord:
    
    def __init__(self):
        self.car_usage = {}
        self.op_info = None

    def __str__(self):
        return "< op_info:{} car_usage:{} >".format(self.op_info, self.format_car_usage())
        
    def is_avalible(self):
        t = map(lambda x: x > 0, self.car_usage.values())
        return any(t)
    
    def set_usage(self, car_usage):
        self.car_usage = car_usage
        pass

    def format_car_usage(self):
        arr = [str(self.car_usage[column_letters_id_dic[col]]) for col in new_part_usage_cols]
        arr_str = ''.join(arr)
        return arr_str

def src_letter_to_id(letter):
    return sheet_src["{}{}".format(letter, src_id_row)].value

def load_cell_part_id(sheet, col, row):
    cell_value = sheet["{}{}".format(col, row)].value

    if type(cell_value) is unicode:
        cell_value = u"".join(cell_value.split()) # delete nbsp
        return cell_value.encode('utf-8')

    if type(cell_value) is str:
        return cell_value

    if cell_value == None:
        return None
        
    return str(cell_value)

def load_cell_car_usage(sheet, col, row):
    
    cell_value = sheet["{}{}".format(col, row)].value

    if type(cell_value) is None:
        return 0

    if type(cell_value) is not long and type(cell_value) is not float and type(cell_value) is int:
        print "Error: {}{} {}({}) is not number!!".format(col, row, cell_value, type(cell_value))

    return cell_value

def print_parts_wrapper_objs(parts_wrapper_objs):

    print "< print_parts_wrapper_objs >\n"

    for item in parts_wrapper_objs.values():
        print item.get_detail_str()
    
    print "\n< print_parts_wrapper_objs end >"

def is_valid_part_id(part_id):
    if type(part_id) is str:
        t1 = part_id.split('-')
        if len(t1) > 2:
            t2 = map(lambda x:x == '0' or x == '', t1)
            t3 = all(t2) == False
            return t3
    
        
    return False

# def process_parts_initial_table():
#     parts_target_table_path = u'D:\\Repositories\\py_charles\\cmd\\1234.xlsx'
#     targert_parts_col = 'AR'
#     targert_parts_row_start = 10
#     targert_parts_row_end = 100

src_path = u'D:\\Repositories\\py_charles\\cmd\\123.xlsx'

print 'loading... ', src_path
since = time.time()

workbook_src = load_workbook(src_path, data_only=True)

print time.time() - since
print "load complete! wait for processing ..."

sheet_src = workbook_src['Change Log']

src_id_row = 5
src_data_start_row = 7

new_part_id_col = 'M'
old_part_id_col = 'AR'
new_part_usage_cols = ['O','P','Q','R','S','T','U','V']
old_part_usage_cols = ['AT','AU','AV','AW','AX','AY','AZ','BA']

column_letters = map(lambda i: get_column_letter(i), range(1, sheet_src.max_column + 1))
column_letters_id_dic = {k: src_letter_to_id(k) for k in column_letters}

changelist = []
for row in range(src_data_start_row, sheet_src.max_row + 1):
    
    new_part_id = load_cell_part_id(sheet_src, new_part_id_col, row)
    new_part_usage = {}
    for col in new_part_usage_cols:
        cell_value = load_cell_car_usage(sheet_src, col, row)
        new_part_usage[column_letters_id_dic[col]] = cell_value

    new_part_data = PartDataRecord()
    new_part_data.set_usage(new_part_usage)

    old_part_id = load_cell_part_id(sheet_src, old_part_id_col, row)
    old_part_usage = {}
    for col in old_part_usage_cols:
        cell_value = load_cell_car_usage(sheet_src, col, row)
        old_part_usage[column_letters_id_dic[col]] = cell_value
    
    old_part_data = PartDataRecord()
    old_part_data.set_usage(old_part_usage)

    part_change = PartChanging(new_part_id, old_part_id, new_part_data, old_part_data, row)
    changelist.append(part_change)

try:
    parts_wrapper_objs = {}
    
    print len(changelist)
    
    for item in changelist:
        item.do(parts_wrapper_objs)
    
    print 'build parts object complete! \n \n'

    print_parts_wrapper_objs(parts_wrapper_objs)

except Exception as err:
    print err
    print traceback.format_exc()

raw_input()