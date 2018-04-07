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
    
    def do(self, parts_wrapper_objs):
        
        if self.old_part_id == None:
            print 'self.old_part_id == None in row {}'.format(self.row_number)
            return

        init_op_info = '{} by row {}'.format(op_init,  self.row_number)
        PartChanging.init_before_operation(parts_wrapper_objs, self.old_part_id, self.old_part_data, init_op_info)
        PartChanging.init_before_operation(parts_wrapper_objs, self.new_part_id, self.new_part_data, init_op_info)

        if self.new_part_id == None: # delete
            
            pass
        else if self.new_part_id == self.old_part_id: # override
            
            pass
        else: # new operation: override old part + init new part
            
            pass
            
        pass

    @staticmethod
    def init_before_operation(parts_wrapper_objs, part_id, part_data, op_info):
        if part_id != None and part_id not in parts_wrapper_objs:
            init_data = deepcopy(part_data)
            init_data.op_info = op_info

            t = PartWrapper(part_id)
            t.append_part_data_record(init_data)
            parts_wrapper_objs[part_id] = t
    
    def op_type(self):
        if self.new_part_id == None:
            return op_delete
        if self.new_part_id == self.old_part_id:
            return op_override
        else:
            return op_new

class PartWrapper:
    
    def __init__(self, part_id):
        self.pre_part = None
        self.next_part = None
        self.part_id = part_id
        self.part_data_records = []

    
    def is_root_wrapper(self):
        return pre_part == None

    def is_final_wrapper(self):
        return next_part == None
    
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
    
    def __init__(self, car_usage):
        self.car_usage = car_usage
        self.change_info = None
        pass
    
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