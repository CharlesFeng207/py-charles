# coding:utf8

import json
import shutil
import time
from openpyxl import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.utils import get_column_letter
from copy import deepcopy

class PartChanging:
    
    op_init = 'op_init'
    op_override = 'op_override'
    op_delete = 'op_delete'

    def __init__(self, new_part_id, old_part_id, new_part_data, old_part_data, row_number):
        self.new_part_id = new_part_id
        self.old_part_id = old_part_id
        self.old_part_data = old_part_data
        self.new_part_data = new_part_data
        self.row_number = row_number

    def __str__(self):
         return 'PartChanging id:{} -> {} row:{}'.format(
             self.old_part_id, self.new_part_id, self.row_number)
    
    def do(self, parts_wrapper_objs):
        
        if self.old_part_id == None:
            print 'self.old_part_id == None in row {}'.format(self.row_number)
            return

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
        else: # new operation: override old part + init new part + build
            modified_data = deepcopy(self.old_part_data)
            modified_data.op_info = '{} {} by row {}, new part id: {}'.format(
                PartChanging.op_override, self.old_part_id, self.row_number, self.new_part_id)

            print modified_data.op_info

            for k in self.new_part_data.car_usage:
                if k in modified_data.car_usage and self.new_part_data.car_usage[k] > 0:
                    modified_data.car_usage[k] = 0
            parts_wrapper_objs[self.old_part_id].append_part_data_record(modified_data)
           
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
        self.next_part = None
        self.part_id = part_id
        self.part_data_records = []        

    def is_root_wrapper(self):
        return self.pre_part == None

    def is_final_wrapper(self):
        return self.next_part == None
    
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
        
    def is_avalible(self):
        t = map(lambda x: x > 0, self.car_usage.values())
        return any(t)
    
    def update_from_new_usage(self, card_usage):
        for k in card_usage:
            if k in self.car_usage:
                if card_usage[k] > 0:
                    self.car_usage[k] = 0
        pass

    def clear_usage(self):
        self.car_usage = {}
        pass
    
    def set_usage(self, car_usage):
        self.car_usage = car_usage
        pass

def src_letter_to_id(letter):
    return sheet_src["{}{}".format(letter, src_id_row)].value

def format_cell_value(cell_value):
    
    if type(cell_value) is unicode:
        cell_value = u"".join(cell_value.split()) # delete nbsp
        return cell_value.encode('utf-8')

    if type(cell_value) is str:
         return cell_value

    return str(cell_value)

def print_parts_wrapper_objs(parts_wrapper_objs):
    for item in parts_wrapper_objs.values():
        if item.is_root_wrapper():
            print item.part_id

src_path = u'D:\\Repositories\\py_charles\\cmd\\Change Log - Copy.xlsx'

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
    
    new_part_id = format_cell_value(sheet_src["{}{}".format(new_part_id_col, row)].value)
    new_part_usage = {}
    for col in new_part_usage_cols:
        cell_value = sheet_src["{}{}".format(col, row)].value
        new_part_usage[column_letters_id_dic[col]] = cell_value

    new_part_data = PartDataRecord()
    new_part_data.car_usage = new_part_usage

    old_part_id = format_cell_value(sheet_src["{}{}".format(old_part_id_col, row)].value)
    old_part_usage = {}
    for col in old_part_usage_cols:
        cell_value = sheet_src["{}{}".format(col, row)].value
        old_part_usage[column_letters_id_dic[col]] = cell_value
    
    old_part_data = PartDataRecord()
    old_part_data.car_usage = old_part_usage

    part_change = PartChanging(new_part_id, old_part_id, new_part_data, old_part_data, row)
    changelist.append(part_change)

try:
    parts_wrapper_objs = {}
    
    print len(changelist)
    
    for item in changelist:
        print item
        item.do(parts_wrapper_objs)
        
    print_parts_wrapper_objs(parts_wrapper_objs)

except Exception as err:
    print err

raw_input()