from datetime import date, datetime
import regex as re
from typing import Literal,Any

DEBUG = False

class KSSerializedItem:
    __uid = 0
    def __init__(self, 
                 serial_num:str = "N/A", 
                 flags:list[str] = [],
                 parent = None,
                 active:bool = True,
                 new = True):
        self.serial_num:str = serial_num.upper()
        self.active = active
        self._flags:set[str] = set(flags)
        self._parent:KSItem = parent
        self._uid = KSSerializedItem.__uid
        self.new = new
        KSSerializedItem.__uid += 1

    @property
    def flags(self):
        return self._flags
    @property
    def id(self):
        return self._uid
    @property
    def uid(self):
        return self._uid
    @property
    def parent(self):
        return self._parent

    def set_flags(self, flags:list[str]):
        for flag in flags:
            self.set_flag(flag)   
    def remove_flags(self, flags:list[str]):
        for flag in flags:
            self.remove_flag(flag)
    def set_flag(self, flag:str):
        if len(flag) <= 0: return
        self._flags.add(flag.lower())
    def remove_flag(self, flag:str):
        if len(flag) <= 0: return
        try:
            self._flags.remove(flag.lower())
        except KeyError:
            return
    def clear_flags(self):
        self.flags.clear()
        
    def has_flags_allof(self, flags:list[str] = []):
        for flag in flags:
            if not(flag.lower() in self._flags):
                return False
        return True
    
    def has_flags_oneof(self,flags:list[str] = []):
        for flag in flags:
            if flag.lower() in self._flags:
                return True
        return False

    def get_file_string(self):
        return f"{self.serial_num}\x1F" + ('\x07'.join(self._flags) + f"\x1F{'1' if self.active else '0'}\x1F{'1' if self.new else '0'}")
    @staticmethod
    def from_file_string(line:str):
        cereal = line.split("\x1F")
        return KSSerializedItem(
            serial_num = cereal[0], 
            flags = cereal[1].split('\x07') if len(cereal[1]) > 0 else [],
            active= cereal[2] == '1',
            new = cereal[3] == '1')

    def restore(self, update_qoh = True):
        if not self.parent:
            raise RuntimeError("Parent Error")
        self.parent.restore_serial_num(self,update_qoh)
    def remove(self, update_qoh = True):
        if not self.parent:
            raise RuntimeError("Parent Error")
        self.parent.remove_serial_num(self,update_qoh)
    
    def __repr__(self):
        return f"{self.id}: {self.serial_num:<50} {';'.join(self.flags)}{' removed' if not self.active else ''}{' new' if self.new else ''}"
        # date format if needed: {self.valid.strftime(r'%m/%d/%y')}\t

class KSItem:
    def __init__(self,
                 id:int,
                 prod_code:str = "Not Defined",
                 desc:str = "N/A",
                 qoh:float = 0,
                 phys:float = None,
                 serial_nums:list[KSSerializedItem] = [],
                 last_count:date = date(date.today().year - 50,1,1),
                 serialized:bool = True,
                 cost: float = 0.0,
                 retail: float = 0.0,
                 flags:list[str] = []):
        self._id:int = id
        self.prod_code:str = prod_code
        self.desc:str = desc
        self._qoh:float = float(qoh)
        self.serial_nums:list[KSSerializedItem] = serial_nums.copy()
        self.last_count:date = last_count
        self.serialized:bool = serialized
        self._flags:set[str] = set(flags)
        self._phys:float = 0.0
        self.cost:float = cost
        self.retail:float = retail
        
        for sn in self.serial_nums:
            sn._parent = self
        
        if phys == None:
            self.phys = qoh
        else:
            self.phys = float(phys)
        
        if self.phys < 0:
            self.phys = 0.0
    
    @property
    def id(self):
        return self._id
    @property
    def flags(self):
        return self._flags

    @property
    def qoh(self):
        return self._qoh
    @qoh.setter
    def qoh(self,value):
        self._qoh = float(value)

    @property
    def phys(self):
        return self._phys
    @phys.setter
    def phys(self,value):
        self._phys = float(value)
    
    @property
    def sn_count(self):
        return self.get_serial_count()
    
    @property
    def tsn_count(self):
        return len(self.serial_nums)
    
    def increase_count(self,amount = 1):
        self.phys += amount

    def decrease_count(self,amount = 1):
        self.phys -= amount

    def reset_count(self, value:float = 0):
        self.phys = float(value)

    def remove_flags(self, flags:list[str] = []):
        if len(self._flags) < 1: return
        try:
            for flag in flags:
                self._flags.remove(flag)
        except KeyError:
            return
    def clear_flags(self):
        self._flags.clear()

    def has_flaged(self, flags:list[str] = [], nflags:list[str] = []):
        for s in self.serial_nums:
            if s.has_flags_oneof(nflags): 
                return False
            if s.has_flags_allof(flags):
                return True
        return False
    
    def get_flaged(self, flags:list[str] = [], nflags:list[str] = []):
        return_val:list[KSSerializedItem] = []
        for s in self.serial_nums:
            if s.has_flags_oneof(nflags): continue
            if s.has_flags_allof(flags):
                return_val.append(s)
        return return_val

    def get_new(self, flags:list[str] = [], nflags:list[str] = [], invert = False):
        return_list:list[KSSerializedItem] = []
        for s in self.serial_nums:
            if s.has_flags_oneof(nflags): continue
            if s.new != invert and s.has_flags_allof(flags):
                return_list.append(s)
        return return_list
    
    def get_removed(self, flags:list[str] = [], nflags:list[str] = [], invert = False):
        return_list:list[KSSerializedItem] = []
        for s in self.serial_nums:
            if s.has_flags_oneof(nflags): continue
            if s.active == invert and s.has_flags_allof(flags):
                return_list.append(s)
        return return_list
    
    def get_file_string(self, flags = [], nflags = []):
        return_vals = (
            str(self.id),
            self.prod_code,
            self.desc,
            str(self.qoh),
            str(self.phys),
            self.last_count.strftime(r'%m/%d/%Y'),
            str(self.serialized),
            str(self.cost),
            str(self.retail),
            "\x1D".join( (x.get_file_string() for x in self.serial_nums if x.has_flags_allof(flags) and not x.has_flags_oneof(nflags)) ),
            "\n"
        )
        return "\t".join(return_vals)
    @staticmethod
    def from_file_string(line:str):
        fields = line.split("\t")
        if len(fields) < 10:
            # support for old files
            item = KSItem(
                id = int(fields[0]),
                prod_code = fields[1],
                desc = fields[2],
                qoh = float(fields[3]),
                phys = float(fields[4]),
                last_count = datetime.strptime(fields[5],r'%m/%d/%Y').date(),
                serialized = fields[6] == "True",)
            for serial in fields[7].split("\x1D"):
                    if len(serial) > 0:
                        item.add_serial_item(KSSerializedItem.from_file_string(serial), update_qoh=False)
        else:
            item = KSItem(
                id = int(fields[0]),
                prod_code = fields[1],
                desc = fields[2],
                qoh = float(fields[3]),
                phys = float(fields[4]),
                last_count = datetime.strptime(fields[5],r'%m/%d/%Y').date(),
                serialized = fields[6] == "True",
                cost = float(fields[7]),
                retail = float(fields[8]))
            for serial in fields[9].split("\x1D"):
                    if len(serial) > 0:
                        item.add_serial_item(KSSerializedItem.from_file_string(serial), update_qoh=False)
        return item
    
    def merge(self, new_item, options = [], exclude_flags = []):
        self.qoh = max(new_item.qoh, self.qoh)
        self.phys += new_item.phys
        
        if new_item.last_count > self.last_count:
            self.last_count = new_item.last_count
        
        if not self.serialized:
            self.serialized = new_item.serialized 
            
        my_sns = self.get_flaged(nflags=exclude_flags)
        item_sns = new_item.get_flaged(nflags=exclude_flags)
        
        unique_sns:dict[str,list[KSSerializedItem]] = {}
        new_sns:list[KSSerializedItem] = []
        for new_sn in item_sns:
            if not new_sn.new:
                if not new_sn.serial_num in unique_sns:
                    unique_sns[new_sn.serial_num] = []
                unique_sns[new_sn.serial_num].append(new_sn)
            else:
                new_sns.append(new_sn)
        
        pending:dict[str,list[KSSerializedItem]] = {}
        for sn in my_sns:
            try:
                if len(unique_sns[sn.serial_num]) < 1:
                    unique_sns.pop(sn.serial_num)
                    continue
                new_sn = unique_sns[sn.serial_num].pop()
                if new_sn.active:
                    if not sn.active:
                        sn.active = new_sn.active
                        sn.set_flags(new_sn.flags)
                    else:
                        if sn.serial_num not in pending:
                            pending[sn.serial_num] = []
                        pending[sn.serial_num].append(new_sn)
                else:
                    if not sn.active:
                        if len(pending[sn.serial_num]) < 1:
                            pending.pop(sn.serial_num)
                            sn.set_flags(new_sn)
                        else:
                            new_sn = pending[sn.serial_num].pop()
                            sn.active = new_sn.active
                            sn.set_flags(new_sn.flags)
            except KeyError:
                continue
    
        for value in unique_sns.values():
            if len(value) < 1: continue
            
            for new_sn in value:
                self.add_serial_item(new_sn, update_qoh=False)
        
        for new_sn in new_sns:
            self.add_serial_item(new_sn,update_qoh=False)

        if self.serialized:
            self.phys = float(self.sn_count)
            
    def update(self, new_item, options = [], exclude_flags = []):
        unique_sns:dict[str,list[KSSerializedItem]] = {}
        my_sns = self.get_new(nflags=exclude_flags,invert = True)
        item_sns = new_item.get_new(nflags=exclude_flags,invert = True)
        
        
        if("keep_phys" not in options):
            self.phys += new_item.qoh - self.qoh
        self.qoh = new_item.qoh
        self.last_count = new_item.last_count
        self.prod_code = new_item.prod_code
        self.desc = new_item.desc
        self.serialized = new_item.serialized
        self.cost = new_item.cost
        self.retail = new_item.retail
        
        for sn in my_sns:
            if not sn.serial_num in unique_sns:
                unique_sns[sn.serial_num] = []
            unique_sns[sn.serial_num].append(sn)
        
        for sn in item_sns:
            try:
                if len(unique_sns[sn.serial_num]) < 1:
                    self.add_serial_item(sn, update_qoh = False)
                    unique_sns.pop(sn.serial_num)
                    continue
                
                my_sn = unique_sns[sn.serial_num].pop()
                my_sn.set_flags(sn.flags)
            except KeyError:
                self.add_serial_item(sn, update_qoh=False)
                continue
    
        for value in unique_sns.values():
            if len(value) < 1: continue
            
            for sn in value:
                sn.parent._delete_serial_item(sn, update_qoh=False)
        if self.serialized:
            self.phys = float(self.sn_count)

    def add_serial_num(self, serial_num:str, update_qoh = True, new = True,flags = []):
        new_item = KSSerializedItem(serial_num.upper(), flags=flags, new = new, parent=self)
        return self.add_serial_item(new_item, update_qoh = update_qoh)
    def add_serial_item(self, serial_item:KSSerializedItem, update_qoh = True, flags = []):
        serial_item.set_flags(flags)
        serial_item._parent = self
        self.serial_nums.append(serial_item)
        if update_qoh:
            self.increase_count()
        return serial_item
    
    def _delete_serial_item(self, serial_item:KSSerializedItem, update_qoh = True):
        try:
            self.serial_nums.remove(serial_item)
            if update_qoh and serial_item.active:
                self.decrease_count()
        except ValueError:
            return
        
    def remove_serial_num(self, serial_item:KSSerializedItem, update_qoh = True):
        if not serial_item.active: return
        serial_item.active = False
        if update_qoh:
            self.decrease_count()
        if serial_item.new:
            self._delete_serial_item(serial_item,update_qoh)
        
    def restore_serial_num(self, serial_item:KSSerializedItem, update_qoh = True):
        if serial_item.active: return
        serial_item.active = True
        if update_qoh:
            self.increase_count()

    # Gets the count of ACTIVE serial numbers, flaged numbers, or numbers not flaged
    def get_serial_count(self, flags = [], nflags = []):
        count = 0
        for sn in self.serial_nums:
            if sn.active and sn.has_flags_allof(flags) and not sn.has_flags_oneof(nflags):
                count += 1
        return count
    
    def get_serial_by_id(self, id:int):
        for s in self.serial_nums:
            if id == s.id:
                return s
        return None

    def __repr__(self):
        q = "'"
        c = ":"
        out_str = f"{self.prod_code + c:<8} \'{self.desc + q:<32} QOH: {int(self.qoh) if self.qoh.is_integer() else self.qoh:<3} Phys: {int(self.phys) if self.phys.is_integer() else self.phys:<3}" + (f"S/N's: {self.get_serial_count():<3}" if self.serialized else "          ") + f" Avg Cost: ${self.cost:.2f}"
        return out_str

class KSSearchResult:
    def __init__(self, search_terms:dict[str,str] = {}, search_results:dict[KSItem,list[KSSerializedItem]] = {}):
        self.search_terms:dict[str,str] = search_terms
        self.result:dict[KSItem,list[KSSerializedItem]] = search_results
        self.keys = self.result.keys
        self.values = self.result.values
        self.items = self.result.items
        
    def update(self, newdata):
        self.search_terms = newdata.search_terms
        self.result = newdata.result
        self.keys = newdata.result.keys
        self.values = newdata.result.values
        self.items = newdata.result.items
    
    def get_first_sn(self):
        for value in self.values():
            return value[0]
    def get_first_item(self):
        for key in self.keys():
            return key
    def get_items(self):
        return list(self.result.keys())
    def get_serial_nums(self):
        rval = []
        for value in self.result.values():
            rval.extend(value.copy())
        return rval
    def __repr__(self):
        rval = ""
        for key,value in self.result.items():
            rval += repr(key) + "\n"
            if not key.serialized: continue
            for sn in value:
                rval += "\t" + repr(sn) + "\n"
        return rval
    def __getitem__(self, key):
        return self.result[key]

class KSDataSet:
    def __init__(self):
        self.items:dict[int,KSItem] = {}
        self.filename = ""
        self.last_save = date.today()
        self.meta_data = []
        self.__created_ids__ = 0
        KSSerializedItem.__uid = 0

    def reset(self):
        self.items.clear()
        self.filename = ""
        self.meta_data = []
        self.last_save = date.today()
        KSSerializedItem.__uid = 0
        self.__created_ids__ = 0

    def import_file(self, file_name:str, resolve_conflicts:Literal["skip","merge","update","replace"] = "merge", exclude_flags = []):
        with open(file_name,encoding="cp1252") as in_file:
            
            self._import_file_name = file_name
            if DEBUG:
                out_file = open("./test_records/field_map.txt","w")
            else:
                out_file = None
            
            line = in_file.readline().strip()
            while len(line) > 0: 
                fields = line.split('\t')
                #  0 = KeyStoke ID
                #  1 = Product Code
                #  3 = Description
                #  5 = Base Price
                #  6 = Sale Price
                #  7 = Avg Cost
                #  9 = QOH
                # 13 = Service Item
                # 32 = Serial Numbers
                # 42 = Serialized Item
                # 61 = Last varianced
                if DEBUG and out_file != None:
                    for i,f in enumerate(fields):
                        if f != '\n':
                            if len(f) > 15:
                                f = f[:15]
                            out_file.write(f"{i} : {f:<15}  ")
                    out_file.write("\n")
                
                if len(fields) < 131:
                    print("Error reading import file")
                    return
                
                if fields[13] == "0":
                    last_count = date.today()
                    try:
                        last_count = datetime.strptime(fields[61],r'%m/%d/%y').date()
                    except ValueError:
                        pass
                    try:
                        item = KSItem(
                            id = int(fields[0]),
                            prod_code = fields[1],
                            desc = fields[3],
                            qoh = float(fields[9]),
                            last_count = last_count,
                            serialized = fields[42] == "1",
                            cost = float(fields[7]),
                            retail = float(fields[5]))
                    except ValueError:
                        pass
                    else:
                        for cereal in fields[32].split("\xff"):
                            if cereal != "":
                                cereal = cereal.upper()
                                item.add_serial_num(cereal, update_qoh = False , new = False)
                        if item.serialized:
                            item.phys = float(item.sn_count)
                        exists = item.id in self.items
                        if not exists:
                            self.items[item.id] = item
                        else:
                            if resolve_conflicts == "replace":
                                self.items[item.id] = item
                            elif resolve_conflicts == "merge":
                                self.items[item.id].merge(item, exclude_flags=exclude_flags)
                            elif resolve_conflicts == "update":
                                self.items[item.id].update(item, exclude_flags=exclude_flags)
                
                line = in_file.readline().strip()
            
            if out_file != None:
                out_file.close()
        self.refesh_serial_ids()
    
    def export_file(self, file_name:str, flags = [], nflags = []):
        # By default Keystroke uses KEYSTOK\Forms\KSIMPORT.KSI which is a binary file that defines the import field mapping.
        # It can be setup using the Importer custom module in Keystroke.
        # KSIMPORT.KSI must have the following definition:
        #   Number:         1
        #   Code:           2
        #   Description:    3
        #   QOH:            4
        #   Serial Numbers: 5
        #   Last Count:     6
        #   Tab delimited
        with open(file_name,"w", encoding="cp1252") as out_file:
            for key,value in self.items.items():
                if key < 0: continue
                if not value.has_flaged(flags): continue
                serial_string = ""
                for x in value.serial_nums:
                    if x.has_flags_oneof(nflags): continue
                    if not x.has_flags_allof(flags): continue
                    if x.active:
                        serial_string += x.serial_num + "\xff"
                
                line = "\t".join((
                    str(key),           #1
                    value.prod_code,    #2
                    value.desc,         #3
                    str(value.phys),     #4
                    serial_string,      #5
                    date.today().strftime(r"%m/%d/%y"),    #6
                    '\n'
                ))
                out_file.write(line)
                
    # Included for backwards compatability
    def export_flaged(self, file_name:str, flags = []):
        return self.export_file(file_name,flags)

    def _prep_for_variance(self, value):
        serial_string = ""
        instruct_string = ""
        export_phys = value.phys

        avalues = []
        for x in value.get_new():
            avalues.append(x.serial_num)
        
        rvalues = []
        for x in value.get_removed():
            rvalues.append(x.serial_num)
        
        diff = export_phys - value.qoh
        if diff > 0:
            i = 0
            while i < diff and i < len(avalues):
                serial_string += f"S/N:{avalues[i]}|"
                i += 1
            
            j = 0
            while i < len(avalues) and j < len(rvalues):
                serial_string += f"Replace: '{rvalues[j]}' with '{avalues[i]}'|"
                instruct_string += f"Replace: '{rvalues[j]}' with '{avalues[i]}'\n"
                i += 1
                j += 1

            while j < len(rvalues):
                serial_string += f"Remove: '{rvalues[j]}'|"
                instruct_string += f"Remove: '{rvalues[j]}'\n"
                j+= 1

        elif diff < 0:
            i = 0
            while i < -diff and i < len(rvalues):
                serial_string += f"S/N:{rvalues[i]}|"
                i += 1
            
            j = 0
            while i < len(rvalues) and j < len(avalues):
                serial_string += f"Replace: '{rvalues[i]}' with '{avalues[j]}'|"
                instruct_string += f"Replace: '{rvalues[i]}' with '{avalues[j]}'\n"
                i += 1
                j += 1

            while j < len(avalues):
                serial_string += f"Add: '{avalues[j]}'|"
                instruct_string += f"Add: '{avalues[j]}'\n"
                j+= 1
        else:
            i = 0
            j = 0
            while i < len(avalues) and j < len(rvalues):
                serial_string += f"Replace: '{rvalues[j]}' with '{avalues[i]}'|"
                instruct_string += f"Replace: '{rvalues[j]}' with '{avalues[i]}'\n"
                i += 1
                j += 1
                
            while j < len(rvalues):
                serial_string += f"Remove: '{rvalues[j]}'|"
                instruct_string += f"Remove: '{rvalues[j]}'\n"
                j+= 1

            while i < len(avalues):
                serial_string += f"Add: '{avalues[i]}'|"
                instruct_string += f"Add: '{avalues[i]}'\n"
                i+= 1


        serial_string = serial_string.strip("|")
        
        return (export_phys, serial_string, instruct_string)
    
    # TODO: Allow configuration for product code or stock number exports.
    def export_variance(self, file_name:str):
        instructions = ""
        with open(file_name,"w", encoding="cp1252") as out_file:
            for key,value in self.items.items():
                prep_result = self._prep_for_variance(value)
                instructions += prep_result[2]
                # Variance CSV format: 
                # "[Stock #]","[phys_qty]","[RESERVED]","[comment/sn 1]|[comment/sn 2]|...|[comment/sn n]"
                line = ",".join((
                    f'"{key}"',    #1
                    f'"{str(prep_result[0])}"',     #2
                    '""',                      #3
                    f'"{prep_result[1]}"\n'     #4
                ))
                out_file.write(line)
        return instructions

    def read_file(self, filename:str, resolve_conflicts:Literal["skip","merge","update","replace"] = "merge", options = [], exclude_flags = []):
        with open(filename,"r") as in_file:
            line = in_file.readline()
            if line[-1] == "\n":
                line = line[:-1]
            
            init_data = line.strip("\x1F").split("\x1F")    
            if init_data[0] != "!KSData File":
                print("Invalid File")
                return
            self.last_save = datetime.strptime(init_data[1],r'%m/%d/%Y').date()
            if len(init_data) > 2:
                self.meta_data = init_data[2:]
            
            line = in_file.readline()
            while len(line) > 0:
                item = KSItem.from_file_string(line)
                if not item.id in self.items:
                    self.items[item.id] = item
                else:
                    if resolve_conflicts == "replace":
                        self.items[item.id] = item
                    elif resolve_conflicts == "merge":
                        self.items[item.id].merge(item,options=options,exclude_flags=exclude_flags)
                    elif resolve_conflicts == "update":
                        print(options)
                        self.items[item.id].update(item, options=options, exclude_flags=exclude_flags)

                line = in_file.readline()
        self.filename = filename
        self.refesh_serial_ids()

    def save_file(self, file_name:str = "", flags:list[str] = [], nflags = []):
        if file_name == "":
            if self.filename != "":
                file_name = self.filename
            else:
                print("Error: No previous save found. Filename not provided.")
                return
        self.filename = file_name
        with open(file_name,"w",encoding="utf-8") as out_file:
            out_file.write(f"!KSData File\x1F{date.today().strftime(r'%m/%d/%Y')}\x1F" + "\x1F".join(self.meta_data) + "\n")
            for key,value in self.items.items():
                out_file.write(value.get_file_string(flags, nflags))

    def _print_all(self):
        for key,value in self.items.items():
            print(repr(value))
            if len(value.serial_nums) > 0:
                print("----------------------")
                for x in value.serial_nums:
                    print(repr(x))
                print()

    def move_serial_item(self, serial_num:KSSerializedItem, to_item:KSItem):
        if not serial_num.parent:
            raise RuntimeError("Parent Error")
        serial_num.parent.remove_serial_num(serial_num)
        to_item.add_serial_item(serial_num)

    def forget_items(self,items:list[KSItem]):
        for item in items:
            self.items.pop(item.id)

    def forget_item(self,item:KSItem):
        self.items.pop(item.id)
    # def clean_all(self):
    #     for value in self.items.values():
    #         value.clean_up_inactive()

    # created items have negative IDs and will not be exported into Keystroke Database 
    # files, but can be saved for future reference and added in properly by hand.
    #
    # This is to avoid overlapping Stock Numbers in Keystroke
    #
    # ~~Created items will be exported to Keystroke Variance files since Variancing only
    # uses Product Codes~~
    #
    # Edit: Varianceing may use stock numbers if configured correctly. Created Items will
    # not be exportable in that case.
    def create_item(self, item:KSItem):
        self.__created_ids__ -= 1
        item._id = self.__created_ids__
        self.items[item.id] = item
        return item
    
    def assert_parents(self):
        for item in self.items.values():
            for sn in item.serial_nums:
                if sn._parent == None:
                    raise RuntimeError("Parent Error")
    
    def search(self,serial_id:list[str] = [], nserial_id:list[str] = [], prod_code:list[str] = [], nprod_code:list[str] = [], desc:list[str] = [], ndesc:list[str] = [],
                 eval_str: str = "", serial_num:list[str] = [], nserial_num:list[str] = [], last_count:list[str] = [], nlast_count:list[str] = [],
                 flags:list[str] = [], nflags:list[str] = [], item_only = False, scope:KSSearchResult|None = None, errors = 0, all_items = False, 
                 new = True, removed = True, counted = True):
        
        og_terms:dict[str,str] = {}
        results:dict[KSItem, list[KSSerializedItem]] = {}
        
        if scope == None:
            search_targets = self.items.values()
        else:
            og_terms["scope"] = scope
            search_targets = scope.get_items()
        

        target_buffer:list[KSItem] = []
        regex_flags = re.BESTMATCH|re.IGNORECASE
        
        if len(prod_code) > 0:
            og_terms["prod_code"] = prod_code
            patterns:list[re.Pattern] = []
            for word in prod_code:
                if errors is None or errors != 0:
                    patterns.append(re.compile(r'(' + word + r'){e}',regex_flags))
                else:
                    patterns.append(re.compile(word,regex_flags))
            mindex = 0
            emin = None
            j = 0
            for i in search_targets:
                pemax = None
                for pattern in patterns:
                    result = re.search(pattern, i.prod_code)
                    if result == None:
                        break

                    esum = result.fuzzy_counts[0] + result.fuzzy_counts[1] + result.fuzzy_counts[2]
                    if pemax == None or esum > pemax:
                        pemax = esum
                else:
                    if pemax == None:
                        pemax = len(i.prod_code)
                    if emin == None or pemax < emin:
                        emin = pemax
                        mindex = j
                    if pemax <= emin:
                        if errors == None or errors < 0 or (pemax <= errors):
                            target_buffer.append(i)
                            j += 1
            search_targets = target_buffer[mindex:]
            target_buffer:list[KSItem] = []

        if len(nprod_code) > 0:
            og_terms["nprod_code"] = nprod_code
            patterns:list[re.Pattern] = []
            for word in nprod_code:
                patterns.append(re.compile(word,regex_flags))
            for i in search_targets:
                for pattern in patterns:
                    if re.search(pattern,i.prod_code) != None:
                        break
                else:
                    target_buffer.append(i)
            search_targets = target_buffer.copy()
            target_buffer:list[KSItem] = []

        if len(desc) > 0:
            og_terms["desc"] = desc
            patterns:list[re.Pattern] = []
            for word in desc:
                if errors is None or errors != 0:
                    patterns.append(re.compile(r'(' + word + r'){e}',regex_flags))
                else:
                    patterns.append(re.compile(word,regex_flags))

            mindex = 0
            emin = None
            j = 0
            for i in search_targets:
                pemax = None
                for pattern in patterns:
                    result = re.search(pattern, i.desc)
                    if result == None:
                        break

                    esum = result.fuzzy_counts[0] + result.fuzzy_counts[1] + result.fuzzy_counts[2]
                    if pemax == None or esum > pemax:
                        pemax = esum
                else:
                    if pemax == None:
                        pemax = len(i.desc)
                    if emin == None or pemax < emin:
                        emin = pemax
                        mindex = j
                    if pemax <= emin:
                        if errors == None or errors < 0 or (pemax <= errors):
                            target_buffer.append(i)
                            j += 1
                    
            search_targets = target_buffer[mindex:]
            target_buffer:list[KSItem] = []

        if len(ndesc) > 0:
            og_terms["ndesc"] = ndesc
            patterns:list[re.Pattern] = []
            for word in ndesc:
                patterns.append(re.compile(word,regex_flags))
            for i in search_targets:
                for pattern in patterns:
                    if re.search(pattern, i.desc) != None:
                        break
                else:
                    target_buffer.append(i)
            search_targets = target_buffer.copy()
            target_buffer:list[KSItem] = []
            
        if len(last_count) > 0:
            og_terms["last_count"] = last_count
            # TODO: implement
            print("KSDataSet: search: last_count not implemented yet")
        
        

        id_search_terms = []
        if len(serial_id) > 0:
            og_terms["serial_id"] = serial_id
            for term in serial_id:
                values = term.split("..")
                rmin = int(values[0])
                rmax = int(values[-1])
                id_search_terms.append((min(rmin,rmax),max(rmin,rmax)))

        nid_search_terms = []
        if len(nserial_id) > 0:
            og_terms["nserial_id"] = nserial_id
            for term in nserial_id:
                values = term.split("..")
                rmin = int(values[0])
                rmax = int(values[-1])
                nid_search_terms.append((min(rmin,rmax),max(rmin,rmax)))

        serial_patterns:list[re.Pattern] = []
        if len(serial_num) > 0:
            og_terms["serial_num"] = serial_num
            for word in serial_num:
                serial_patterns.append(re.compile(word,regex_flags))

        nserial_patterns:list[re.Pattern] = []
        if len(nserial_num) > 0:
            for word in nserial_num:
                og_terms["nserial_id"] = nserial_id
                nserial_patterns.append(re.compile(word,regex_flags))
        
        if item_only:
            og_terms["item_only"] = item_only

        all_sns = len(serial_id) == 0 and len(nserial_id) == 0 \
                and len(serial_num) == 0 and len(nserial_num) == 0 \
                and len(flags) == 0 and len(nflags) == 0
                
        filter_sns = len(serial_id) > 0 or len(nserial_id) > 0 \
                or len(serial_num) > 0 or len(nserial_num) > 0 \
                or len(flags) > 0 or len(nflags) > 0
        
        if len(eval_str) > 0:
            og_terms["eval_str"] = eval_str
            # TODO: change to ast
            for x in search_targets:
                if eval(eval_str, {"x":x}):
                    target_buffer.append(x)
            search_targets = target_buffer.copy()
            target_buffer:list[KSItem] = []
                
        for i in search_targets:
            if all_items or not filter_sns:
                results[i] = []
            if item_only:
                results[i] = []
                continue
            
            if not i.serialized:
                continue
            
            if scope == None:
                search_sns = i.serial_nums
            else:
                search_sns = scope[i]

            # if all_sns:
            #     results[i] = search_sns.copy()
            #     continue

            result_sns = []
            for s in search_sns:
                if s.has_flags_oneof(nflags): continue
                if not s.has_flags_allof(flags): continue
                
                if (not new) and s.new: continue
                if (not removed) and (not s.active): continue
                if (not counted) and s.active and not s.new: continue
                
                is_match = True
                
                for pattern in serial_patterns:
                    if (re.search(pattern, s.serial_num) == None):
                        is_match = False
                        break
                
                for pattern in nserial_patterns:
                    if (re.search(pattern, s.serial_num) == None):
                        break
                else:
                    if len(nserial_patterns) > 0:
                        is_match = False

                for term in id_search_terms:
                    if (s.id >= term[0] and s.id <= term[1]):
                        break
                else:
                    if len(id_search_terms) > 0:
                        is_match = False

                for term in nid_search_terms:
                    if (s.id >= term[0] and s.id <= term[1]):
                        is_match = False
                        break

                if is_match:
                    result_sns.append(s)
            if len(result_sns) > 0:
                results[i] = result_sns
        return KSSearchResult(og_terms, results)

    def refesh_serial_ids(self):
        next_id = 0
        for key,value in self.items.items():
            for sn in value.serial_nums:
                sn._uid = next_id
                next_id += 1
        KSSerializedItem.__uid = next_id
    

if __name__ == "__main__":
    DEBUG = True
    def print_results(result:dict[KSItem,list[KSSerializedItem]]):
        for key,value in result.items():
            print(key.prod_code + ":")
            for sn in value:
                print("\t" + repr(sn))
    data = KSDataSet()
    data.import_file("./test_records/EXPORT_HANDBARS.TXT")
    print("------------- import data -------------")
    print_results(data.search())
    
