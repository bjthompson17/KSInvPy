import KSData as ksd
import regex as re
import win32gui,win32con,os
from typing import Literal, Any, Callable
import win32print as wp
import json

data = ksd.KSDataSet()
last_filter_options:tuple[ksd.KSSearchResult | None, dict[str,str], set[str], list[str]] = (None,{},set(),[])
inv_ranges: list[tuple[str]] = []
current_store = 0

CONFIG_PATH = os.path.expanduser("~") + r"\AppData\Local\Python\KSInventoryApp\config;" + os.path.abspath(r".\config")
LOCAL_CONFIG_PATH = os.path.expanduser("~") + r"\AppData\Local\Python\KSInventoryApp\config"
SHARED_CONFIG_PATH = os.path.abspath(r".\config")
MACRO_PATH = os.path.expanduser("~") + r"\AppData\Local\Python\KSInventoryApp\config\macros;" + os.path.abspath(r".\config\macros")
LOCAL_MACRO_PATH = LOCAL_CONFIG_PATH + r"\macros"
SHARED_MACRO_PATH = SHARED_CONFIG_PATH + r"\macros"
MACRO_EXTENSION = ".invm"

def init_config():
    for p in CONFIG_PATH.split(";"):
        if not os.path.exists(p):
            os.makedirs(p + "\macros")
init_config()

## Helpers

def _global_docs(func: Callable):
    """Fill in help documentation with global help info"""
    if func.__doc__ is None:
        return func
    func.__doc__ = func.__doc__.format("","","","","",
        search_switches = """
        /id :
        /nid :
        
        /item :
        /nitem :
        
        /desc :
        /ndesc : 
        
        /sn :
        /nsn :
        
        /flags :
        /nflags :
        
        /errors : 
        """,
        search_options = """
        -all : 
        -item_only, -x : 
        """
    )
    return func

def get_configs(subdir:str = "."):
    rval = {"local":{},"shared":{}}
    local_path = LOCAL_CONFIG_PATH
    shared_path = SHARED_CONFIG_PATH
    local_configs = os.listdir(local_path + "\\" + subdir)
    shared_configs = os.listdir(shared_path + "\\" + subdir)
    
    for file in local_configs:
        abs_path = os.path.abspath(local_path + "\\" + file)
        if os.path.isfile(abs_path) and abs_path.endswith(".json"):
            rval["local"][os.path.splitext(os.path.basename(abs_path))[0]] = abs_path

    for file in shared_configs:
        abs_path = os.path.abspath(shared_path + "\\" + file)
        if os.path.isfile(abs_path) and abs_path.endswith(".json"):
            rval["shared"][os.path.splitext(os.path.basename(abs_path))[0]] = abs_path
    return rval

def load_config(filename:str):
    rval = {}
    try:
        with open(filename, "r") as file:
            rval = json.load(file)
    except FileNotFoundError:
        print("File not found")
        return {}
    except json.JSONDecodeError as err:
        print(f"JSON error: {err.msg}")
    return rval

def save_config(filename:str, config_dict:dict, overwrite = True):
    if os.path.exists(filename) and not overwrite:
        print(f"File '{filename}' already exists")
        return False
    with open(filename, "w") as file:
        json.dump(config_dict,file)
    return True

def open_file_dialogue(title:str = "Select File", default_name:str = "EXPORT", default_dir = os.environ['userprofile'], types:list[tuple[str,str]] = [("Text Files","txt")]):
    filter=''
    for x in types:
        filter += f"{x[0]}\0*.{x[1]}\0"
    
    customfilter='Other file types\0*.*\0'
    try:
        fname, customfilter, flags=win32gui.GetOpenFileNameW(
            InitialDir=default_dir,
            Flags=win32con.OFN_ALLOWMULTISELECT|win32con.OFN_EXPLORER,
            File=default_name, DefExt=types[0][1],
            Title=title,
            Filter=filter,
            CustomFilter=customfilter,
            FilterIndex=0)
        
        file_names:list[str] = []
        fnames = str(fname).split("\x00")
        if len(fnames) > 1:
            for name in fnames[1:]:
                file_names.append(fnames[0]+"\\"+name)
        else:
            file_names.append(fnames[0])
        
        return file_names
    except:
        print("File not found")
        return None
def save_file_dialogue(title:str = "Select File", default_name:str = "EXPORT", default_dir = os.environ['userprofile'], types:list[tuple[str,str]] = [("Text Files","txt")]):
    filter=''
    for x in types:
        filter += f"{x[0]}\0*.{x[1]}\0"
    
    customfilter='Other file types\0*.*\0'
    try:
        Flags = win32con.OFN_EXPLORER
        fname, customfilter, flags=win32gui.GetSaveFileNameW(
            InitialDir=default_dir,
            Flags=Flags,
            File=default_name, DefExt=types[0][1],
            Title=title,
            Filter=filter,
            CustomFilter=customfilter,
            FilterIndex=0)
        if os.path.exists(fname):
            result = win32gui.MessageBox(None,f'"{fname}" already exists.\nWould you like to overwrite it?',"Overwrite file?",win32con.MB_YESNO|win32con.MB_DEFBUTTON2)
            if result == 7:
                raise RuntimeError("Cancel write file.")
            else:
                old_directory = os.path.dirname(fname) + "\\.old\\"
                if not os.path.exists(old_directory):
                    os.mkdir(old_directory)
                print(fname)
                print(old_directory)
                os.system(f'copy /Y "{fname}" "{old_directory}"')
        return str(fname)
    except:
        print("Error writing to file.")
        return None

def get_inv_range():
    rval = []
    min = None
    max = None
    for value in data.items.values():
        if value.id < 0: continue
        if min == None or value.prod_code < min:
            min = value.prod_code
        if max == None or value.prod_code > max:
            max = value.prod_code
            
    return (min,max)
def get_inventory(export_file:str|None = "K:\\InventoryExports\\TEMP_EXPORT.EXP", inv_range_list:list[tuple[str]] = [], sort_field = 2, conflicts:Literal["replace","update","merge","skip"] = "update", store_num = 0):
    global inv_ranges, current_store
    if export_file == None:
        export_file = open_file_dialogue(title = "Import file", default_name="KSEXPORT.EXP", types=[("Keystoke Export","*")], multi_select=False)
    
    if export_file == None: 
        print("Import file not found")
        return False
    
    for inv_range in inv_range_list:
        if None in inv_range: continue
        for i,r in enumerate(inv_ranges):
            # New range is within an existing range.
            if inv_range[0] >= r[0] and inv_range[1] <= r[1]:
                break
            # New range envelopes existing range
            elif inv_range[0] < r[0] and inv_range[1] > r[1]:
                inv_ranges[i] = inv_range
                break
            # New range overlaps lower end of existing range
            elif inv_range[0] < r[0] and (inv_range[1] >= r[0] and inv_range[1] <= r[1]):
                inv_ranges[i] = (inv_range[0], r[1])
                break
            # New range overlaps upper end of existing range
            elif inv_range[1] > r[1] and (inv_range[0] >= r[0] and inv_range[0] <= r[1]):
                inv_ranges[i] = (r[0], inv_range[1])
                break
        else:
            inv_ranges.append(inv_range)
    
    
    current_store = store_num
    
    # Configurable
    # store_data stores the path to each store's data folder
    store_data = (
        "DATA",
        "DATACSU",
        "DATAWH"
    )
    
    if not os.path.exists("K:\\KEYSTROK"):
        print("KEYSTOKE directory does not exist or is not set up right.")
        return False
    os.chdir("K:\\KEYSTROK")
    
    if conflicts == "replace":
        data.reset()
        
    if len(inv_ranges) > 0:
        for r in inv_ranges:
            os.system(f'echo Retrieving Items in range {r[0]},{r[1]}...')
            os.system(f'KSEXPORT.EXE /NOP /NODISPLAY /D {store_data[store_num]} FILE="{export_file}" TAB=ON TYPE=DI SORT={sort_field} START={r[0]} END={r[1]} INIFILE=K:\KEYSTROK\INIS\MAIN\KSEXP.INI')
            data.import_file(export_file,conflicts, exclude_flags = ["new"])
    else:
        os.system(f'echo No import range given ...')
    
    return True

def send_inventory(flags = []):
    if len(data.items) < 1: 
        print("Nothing to send")
        return False
    
    if not os.path.exists("K:\\KEYSTROK"):
        print("KEYSTOKE directory does not exist or is not set up right.")
        return False
    
    get_inventory()
    print("Sending to Keystroke...")
    try:
        data.export_file("K:\\InventoryExports\\IMPORT.IMP", flags = flags)
    except Exception:
        print("Error exporting data")
        return False
    os.chdir("K:\\KEYSTROK")
    os.system(f'echo Sending to Keystoke... & IMP.EXE /NOP /NODISPLAY METHOD=MERGE DBNAME=INV')
    print("Done")
    return True

def get_label_printers(valid_drivers:list[str] = ["Generic / Text Only"]):
    printer_names = []
    for x in wp.EnumPrinters(wp.PRINTER_ENUM_LOCAL):
        printer = wp.OpenPrinter(x[2])
        driver_name = None
        try:
            printer_data = wp.GetPrinter(printer,2)
            driver_name = printer_data["pDriverName"]
        finally:
            wp.ClosePrinter(printer)
        
        if driver_name != None:
            for vd in valid_drivers:
                if vd in driver_name:
                    printer_names.append(x[2])
                    break
    return printer_names.copy()
def print_labels(selected_printer:str,dpi:int,language:str, print_items:ksd.KSSearchResult):
        printer = wp.OpenPrinter(selected_printer)
        output_file = None
        output_type = "RAW"
        printer_info = wp.GetPrinter(printer,2)

        if(printer_info["pPortName"] == "FILE:"):
            output_file = save_file_dialogue(title = "Choose output file", types = [("Text Files","txt"),("Other","*")])
        # elif(printer_info["pPortName"] == "PORTPROMPT:"):
        #     printer_menu.winfo_children()[0].entryconfigure(selected_printer.get(),{"label":"Nope","command":tk._setit(selected_printer,"Nope")})
        #     selected_printer.set("Nope")
        #     return
        try:
            wp.StartDocPrinter(printer,1,("Label",output_file,output_type))
            try:
                for item,value in print_items.items():
                    for sn in value:
                        item_num = item.prod_code
                        serial_num = sn.serial_num
                        if item_num == "" and serial_num == "": continue

                        string = ""
                        if(dpi == 203 and language == "EPL"):
                            string = \
f"""N
Q195,70
R200,0
S1
D7
ZB
B20,25,0,1,2,3,50,B,"{item_num}"
b20,115,P,400,70,s{max(0,int( (64 - len(serial_num)) / 16 ))},f0,x3,y5,"{serial_num}"
P1
"""
                        elif(dpi == 300 and language == "ZPL"):
                            string = \
f"""
^XA^LH60,133^FWN,0
^FO0,10^A@N,40,35,E:CON000.TTF
^BCN,75,Y,N,N,A
^FD{item_num}^FS
^FO0,150^B7N,6,3,14,10,N
^FD{serial_num}^FS
^XZ
"""
                        byte_data = string
                        if output_type == "RAW":
                            byte_data = bytes(string,"utf-8")
                        
                            wp.StartPagePrinter(printer)
                        
                            wp.WritePrinter(printer, byte_data)

                            wp.EndPagePrinter(printer)
            finally:
                wp.EndDocPrinter(printer)
        finally:
            wp.ClosePrinter(printer)

def parse_search_args(switches:dict[str,str], options:set[str]):
    search_args = {}
    for switch,value in switches.items():
        if switch == "id":
            search_args["serial_id"] = qp_split(value,delimiters="\s,")
        elif switch == "nid":
            search_args["nserial_id"] = qp_split(value,delimiters="\s,")
        elif switch == "item":
            search_args["prod_code"] = qp_split(value)
        elif switch == "nitem":
            search_args["nprod_code"] = qp_split(value)
        elif switch == "desc":
            search_args["desc"] = qp_split(value)
        elif switch == "ndesc":
            search_args["ndesc"] = qp_split(value)
        elif switch == "flags":
            search_args["flags"] = qp_split(value,",")
        elif switch == "nflags":
            search_args["nflags"] = qp_split(value,",")
        elif switch == "sn" or switch == "serial":
            search_args["serial_num"] = qp_split(value)
        elif switch == "nsn" or switch == "nserial":
            search_args["nserial_num"] = qp_split(value)
        elif switch == "eval":
            search_args["eval_str"] = value
        elif switch == "errors":
            try:
                search_args["errors"] = int(value)
            except ValueError:
                search_args["errors"] = -1
    
    for option in options:
        if option in ("items_only", "items", "x"):
            search_args["item_only"] = True
        if option in ("all", "a"):
            search_args["all_items"] = True
        if option in ("nnew", "nn"):
            search_args["new"] = False
        if option in ("nremoved", "nr"):
            search_args["removed"] = False
        if option in ("ncounted", "nc"):
            search_args["counted"] = False
                
    return search_args

def parse_command_str(command_str:str):
    command_str = command_str
    seperated_command_pattern = re.compile(r'(?:[^&"]|"(?:\\[\\"]|[^"])*")+')
    
    value_pattern = r'(?:"(?:\\[\\"]|.)+?"|(?:\\[\\"\s]|[^\s"])+?)+'
    command_pattern = re.compile(r'(^\w+)|/(\w+)(?:\s*=\s*|\s+)(' + value_pattern + r')|-(\w+)|(' + value_pattern + r')')
    
    command_matches = re.findall(seperated_command_pattern, string = command_str)
    
    def repl_func(match: re.Match):
        s:str = match.group(0)
        if s.startswith('\\'):
            return s[1:]
        return ''
    
    rval = []
    for unique_command_str in command_matches:
        matches:list[re.Match] = re.findall(pattern = command_pattern, string = unique_command_str.strip())
        if len(matches) <= 0:
            return
        command:str = matches[0][0].lower()
        
        switches:dict[str,str] = {}
        options:set[str] = set()
        values:list[str] = []

        
        for arg in matches[1:]:
            if arg[1] != "":
                switches[arg[1].lower()] = re.sub(r'\\\s|\\"|"',repl_func,arg[2])
            if arg[3] != "":
                options.add(arg[3].lower())
            if arg[4] != "":
                # values.append(qp_unescape(arg[4],[r'\s']))
                values.append(re.sub(r'\\\s|\\"|"',repl_func,arg[4]))

        rval.append((command,switches,options,values))
    return rval

def register_command(name:str, func):
    command_list[name] = func
        
def qp_escape(escape_str:str, escape_chars = [r'\\',r'"'], escape_sequence = '\\'):
    sub_str = "|".join(x for x in escape_chars)
    def repl_func(s:re.Match):
        return escape_sequence + s.group(0)
    return re.sub(sub_str, repl_func, escape_str)
def qp_unescape(escape_str:str, escape_chars = [r'\\',r'"'], escape_sequence = '\\'):
    sub_str = "|".join((re.escape(escape_sequence) + x) for x in escape_chars)
    def repl_func(match:re.Match):
        s:str = match.group(0)
        if s.startswith(escape_sequence):
            return s[len(escape_sequence):]
        return ''
    return re.sub(sub_str, repl_func, escape_str)
def qp_split(search_str:str, delimiters = "\s", quote_modifier:Callable[[str],str] = re.escape):
        results = []

        # re.sub(r'\\"|\\\\', repl_func, "")
        # quote_pattern = r'"(?:\\[\\"]|.)*?"'
        term_pattern = re.compile(r'(?:(?:\\[\\"]|[^' + delimiters + r'"])+|"(?:\\[\\"]|.)*?")+')
        for match in re.findall(term_pattern, search_str):
            sub_str = ""
            parse_pattern = re.compile(r'((?:\\[\\"]|[^"])+)|"((?:\\[\\"]|.)*?)"')
            for m in re.findall(parse_pattern, match):
                if m[0] != "":
                    sub_str += qp_unescape(m[0],[r"\\",r'"'])
                elif m[1] != "":
                    sub_str += quote_modifier(qp_unescape(m[1],[r"\\",r'"']))
            results.append(sub_str)
        return results

# ===================== Commands =============================
def pack_command(command:str, switches:dict[str,str]|None = None, options:set[str]|None = None, values:list[str]|None = None):
    rval = command
    if switches != None:
        for switch,value in switches.items():
            rval += f' /{switch} "{value}"'
    if options != None:
        for option in options:
            rval += f' -{option}'
    if values != None:
        for value in values:
            rval += f' "{value}"'
    return rval
def pack_switches(**kwargs):
    return kwargs
def pack_options(**kwargs):
    rval = set()
    for key,value in kwargs.items():
        if value:
            rval.add(key)
    return rval
def pack_values(*args):
    return args
def hasoneof(options:set[str], *args):
    for arg in args:
        if arg in options:
            return True
    return False

## Commands
# @_global_docs
def filter(scope = None,switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []) -> ksd.KSSearchResult:
    """Filters all items globaly and returns a KSSearchResult object
    containing items that match. The filter terms used are stored for
    later use.
    """
    global last_filter_options
    search_args = parse_search_args(switches, options)
    filter_results = data.search(**search_args)
    last_filter_options = (scope,switches,options,values)
    return filter_results
# @_global_docs
def refresh(scope = None,switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    """Reuses the search arguments used in the last call to filter() and
    returns another KSSearchResults object.
    """
    search_args = parse_search_args(last_filter_options[1], last_filter_options[2])
    filter_results = data.search(**search_args)
    return filter_results
# @_global_docs
def find(scope = None,switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []) -> ksd.KSSearchResult:
    """Search for items within a given KSSearchResults scope object. 
    If scope is None, searches globally instead. Returns a KSSearchResults 
    object.
    """
    search_args = parse_search_args(switches, options)
    filter_results = data.search(scope = scope, **search_args)
    return filter_results
# @_global_docs
def add(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    """Add a new serial number to the first serialized item produced from 
    a search. If the serial number is successfully added, returns the added 
    KSSerializedItem object. Otherwise returns None.
    """
    search_args = parse_search_args(switches,options)
    search_args["item_only"] = True
    results = data.search(scope = scope, **search_args)
    items = results.get_items()
    rval = []
    item = None
    
    for x in items:
        if not x.serialized:
            continue
        else:
            item = x
            break
    else:
        return None
    
    for value in values:
        rval.append(item.add_serial_num(value))
    return rval
# @_global_docs
def remove(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    """Remove all serial numbers matching the search terms and id's provided as values.

    Args:
        scope (KSSearchResult): Scope to search within. If None, searches all items.
        switches (dict[str,str], optional): Search filter switches. 
            {search_switches}
        Defaults to {}.
        options (set[str], optional): All search filter options plus . Defaults to set().
        values (list[str], optional): _description_. Defaults to [].

    Returns:
        list[KSItem]: 
    """
    search_args = parse_search_args(switches,options)
    if len(values) > 0 and not "serial_id" in search_args:
        search_args["serial_id"] = []
    for value in values:
        search_args["serial_id"].append(value)
    results = data.search(scope = scope, **search_args)
    rval = []
    for key,value in results.items():
        for sn in value:
            rval.append(sn)
            key.remove_serial_num(sn, not hasoneof(options,"k","keep_qoh"))
    return rval
# @_global_docs
def restore(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    """Restore all serial numbers matching the search terms and id's provided as values."""
    search_args = parse_search_args(switches,options)
    if len(values) > 0 and not "serial_id" in search_args:
        search_args["serial_id"] = []
    for value in values:
        search_args["serial_id"].append(value)
    results = data.search(scope = scope, **search_args)
    rval = []
    for key,value in results.items():
        for sn in value:
            rval.append(sn)
            key.restore_serial_num(sn, not ("keep_qoh" in options))
    return rval
# @_global_docs
def forget(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    """Stop tracking inventory for the items that match the search terms provided

    Args:
        scope (KSSearchResult): Scope to search within. If None, all items are searched.
        switches (dict[str,str], optional): Search switches. Defaults to {}.
        options (set[str], optional): Search options. Defaults to set().
        values (list[str], optional): Not used, but provided for command compatability. Defaults to [].

    Returns:
        list[KSItem]: A list of what was removed. These can be captured or left for garbage collection.
    """
    search_args = parse_search_args(switches,options)
    search_args["item_only"] = True
    results = data.search(scope = scope, **search_args)
    data.forget_items(results.keys())
    return list(results.keys())
# @_global_docs
def flag(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    """Flag all serial numbers that match the search arguments with the flags provided as values.

    Args:
        scope (KSSearchResult | None): Scope to search within. If None, all items are searched.
        switches (dict[str,str], optional): _description_. Defaults to {}.
        options (set[str], optional): _description_. Defaults to set().
        values (list[str], optional): _description_. Defaults to [].

    Returns:
        list[KSItem]: _description_
    """
    search_args = parse_search_args(switches,options)
    results = data.search(scope = scope, **search_args)
    
    rval = []
    for key,value in results.items():
        for sn in value:
            rval.append(sn)
            sn.set_flags(values)
    return rval
# @_global_docs
def unflag(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    search_args = parse_search_args(switches,options)
    results = data.search(scope = scope, **search_args)
    rval = []
    for key,value in results.items():
        for sn in value:
            rval.append(sn)
            if len(set(("a","all")).intersection(options)) > 0:
                sn.clear_flags()
            else:
                sn.remove_flags(values)
    return rval
# @_global_docs
def reflag(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    search_args = parse_search_args(switches,options)
    results = data.search(scope = scope, **search_args)
    rval = []
    for key,value in results.items():
        for sn in value:
            rval.append(sn)
            sn.clear_flags()
            sn.set_flags(values)
    return rval
# @_global_docs
def reset(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    global inv_ranges
    data.reset()
    inv_ranges = []
    return ksd.KSSearchResult()
# @_global_docs
def import_file(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    conflicts = "merge"
    if hasoneof(options,"m","merge"):
        conflicts = "merge"
    elif hasoneof(options,"u","update"):
        conflicts = "update"
    elif hasoneof(options,"s","skip"):
        conflicts = "skip"
    elif hasoneof(options,"r","replace"):
        conflicts = "replace"
    def _wrapper(values:list[str], conflicts = "merge", nflags = "", file = None, **kwargs):
        file_names = values.copy()
        if file != None:
            file_names.append(file)
        elif len(file_names) < 1:
            file_names = open_file_dialogue(title="Import File",types=[("KS Export Files","txt"),("Other","*")])

        if file_names != None:
            for file_name in file_names:
                    data.import_file(file_name, resolve_conflicts=conflicts, exclude_flags=qp_split(nflags,","))
            return file_names
    return _wrapper(values, conflicts, **switches)
# @_global_docs
def export_file(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    file = ""
    if len(values) > 0:
        file = values[0]
    if "file" in switches:
        file = switches["file"]
    elif file == "":
        file = save_file_dialogue(title = "Export file", default_name="KSIMPORT.IMP", types=[("Keystoke Import File","imp")])
    if file == None: return False

    def _wrapper(flags = "",nflags = "",**kwargs):
        data.export_file(file, flags = qp_split(flags,","), nflags = qp_split(nflags,","))
        return file
    return _wrapper(**switches)
# @_global_docs
def export_variance(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    file = ""
    if len(values) > 0:
        file = values[0]
    if "file" in switches:
        file = switches["file"]
    elif file == "":
        file = save_file_dialogue(title = "Export Variance file", default_dir="K:\\InventoryExports", default_name="KSVARIANCE.LOG", types=[("Keystoke Variance File","log")])
    if file == None: return False

    def _wrapper(flags = "",nflags = "",**kwargs):
        get_inventory()
        instructions = data.export_variance(file)
        with open(f"{os.path.dirname(file)}/Instructions.txt","w") as fout:
            fout.write(instructions)
        os.startfile(f"{os.path.dirname(file)}/Instructions.txt")
        return file
    return _wrapper(**switches)
# @_global_docs
def load(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    conflicts = "merge"
    sub_options = []
    if hasoneof(options,"m","merge"):
        conflicts = "merge"
    elif hasoneof(options,"u","update"):
        conflicts = "update"
        if hasoneof(options,"k","keep_phys"):
            sub_options.append("keep_phys")
    elif hasoneof(options,"s","skip"):
        conflicts = "skip"
    elif hasoneof(options,"r","replace"):
        conflicts = "replace"
    def _wrapper(values:list[str], conflicts = "replace", nflags = "", file = None, **kwargs):
        file_names = values.copy()
        if file != None:
            file_names.append(file)
        elif len(file_names) < 1:
            file_names = open_file_dialogue(title="Open File", default_dir="K:\\InventoryExports\\Counts", default_name=f"inventory",types=[("Inventory Count File","count")])

        if file_names != None:
            if conflicts == "replace":
                reset()
            for file_name in file_names:
                data.read_file(file_name, resolve_conflicts=conflicts, options=sub_options, exclude_flags=qp_split(nflags,","))
            
            if len(data.meta_data) > 0:
                ranges = data.meta_data[0].split(",")
                for range_str in ranges:
                    range_ = range_str.split("..")
                    inv_ranges.append((range_[0],range_[-1]))
            else:
                inv_range = get_inv_range()
                if None not in inv_range:
                    inv_ranges.append(inv_range)

            return file_names
    return _wrapper(values, conflicts, **switches)
# @_global_docs
def save(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    file = ""
    if len(values) > 0:
        file = values[0]
    if "file" in switches:
        file = switches["file"]
    elif file == "":
        min,max = get_inv_range()
        file = save_file_dialogue(title = "Save file", 
                                  default_dir="K:\\InventoryExports\\Counts",
                                  default_name=f"inventory_{min}-{max}" if data.filename == "" else data.filename, 
                                  types=[("Inventory Count File","count")])
    if file == None: return False
    if len(data.meta_data) < 1:
        data.meta_data.append("")
    data.meta_data[0] = ",".join("..".join(x) for x in inv_ranges)
    def _wrapper(flags = "",nflags = "",**kwargs):
        data.save_file(file, flags = qp_split(flags,","), nflags = qp_split(nflags,","))
        return file
    return _wrapper(**switches)
# @_global_docs
def refresh_ids(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    data.refesh_serial_ids()
# @_global_docs
def get_inv(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    conflicts = None
    if hasoneof(options,"m","merge"):
        conflicts = "merge"
    elif hasoneof(options,"u","update"):
        conflicts = "update"
    elif hasoneof(options,"s","skip"):
        conflicts = "skip"
    elif hasoneof(options,"r","replace"):
        conflicts = "replace"
    
    def _wrapper(file = None, range = None, sort = None, conflicts = None, store = None, **kwargs):
        func_args = {}
        if file:
            func_args["file_name"] = file
        if sort:
            func_args["sort"] = int(sort)
        if conflicts:
            func_args["conflicts"] = conflicts
        if store:
            func_args["store_num"] = int(store)
        if range:
            ranges = range.split(",")
            func_args["inv_range_list"] = []
            for r in ranges:
                values = r.split("..")
                inv_range = (values[0],values[-1])
                func_args["inv_range_list"].append(inv_range)

        get_inventory(**func_args)

    _wrapper(conflicts = conflicts, **switches)

# @_global_docs
def send_inv(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    send_inventory()

# @_global_docs
def echo(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    rval = ""
    for value in values:
        rval += value + "\n"
    rval += "Switches:\n"
    for key,value in switches.items():
        rval += f"\t{key} = {value}\n"
    rval += "Options:\n"
    for option in options:
        rval += f"\t{option}"
    return rval
# @_global_docs
def count(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    results:ksd.KSSearchResult = find(scope,switches,options,values)
    if len(values) <= 0:
        values.append(1)

    try:
        count = float(values[0])
    except ValueError:
        count = 1.0

    for key,value in results.items():
        if not key.serialized:
            key.increase_count(count)
        else:
            this_count = count
            for sn in value:
                if not sn.active:
                    sn.restore()
                    this_count -= 1
                    if this_count <= 0:
                        break
# @_global_docs
def uncount(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    results:ksd.KSSearchResult = find(scope,switches,options,values)
    if len(values) <= 0:
        values.append(1)
        
    try:
        count = float(values[0])
    except ValueError:
        count = 1.0
    
    for key,value in reversed(results.items()):
        if not key.serialized:
            key.decrease_count(count)
        else:
            this_count = count
            for sn in reversed(value):
                if sn.active:
                    sn.remove()
                    this_count -= 1
                    if this_count <= 0:
                        break
# @_global_docs
def recount(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    search_args = parse_search_args(switches,options)
    search_args["item_only"] = True
    results = data.search(scope = scope, **search_args)
    for item in results.keys():
        if item.serialized:
            for sn in item.serial_nums:
                sn.remove()
        if len(values) > 0:
            item.reset_count(values[0])
        else:
            item.reset_count()
# @_global_docs
def run_macro(scope = None, switches:dict[str,str] = {}, options:set[str] = set(), values:list[str] = []):
    if len(values) < 1:
        return
    file_name = None
    final_path = MACRO_PATH
    
    if "path" in switches:
        final_path = switches["path"] + ';' + final_path
    
    for p in final_path.split(";"):
        file_path = f"{p}\\{values[0]}{MACRO_EXTENSION}"
        if os.path.isfile(file_path):
            file_name = file_path
            break
    else:
        raise FileNotFoundError("Invalid macro name")
    rval = ""
    try:
        commands = []
        with open(file_name, 'r') as file:
            line = file.readline()
            while len(line) > 0:
                formatted = re.sub(r'\$(\d+)',lambda exp: (lambda idx: f'"{values[idx]}"' if idx < len(values) else "")(int(exp.group(1))),line)
                print(f"formatted command: {formatted}")
                commands.append(formatted)
                line = file.readline()

        for command in commands:
            print(f"Running {command}...")
            result = run_command(refresh(), command)
            if result[0]:
                print(":Success:")
                rval += f"{result[1]} ... Success\n"
            else:
                print(":Failed:")
                rval += f"{result[1]} ... Failed\n"
    except FileNotFoundError:
        raise FileNotFoundError("Macro file not found.")
    return rval
    
command_list = {
    "filter":filter,
    "refresh":refresh,
    "find":find,
    "list":find,
    "add":add,
    "remove":remove,
    "restore":restore,
    "forget":forget,
    "flag":flag,
    "unflag":unflag,
    "reflag":reflag,
    "reset":reset,
    "import":import_file,
    "export":export_file,
    "open":load,
    "save":save,
    "refresh_ids":refresh_ids,
    "vexport":export_variance,
    "echo":echo,
    "get_inv":get_inv,
    # "send_inv":send_inv,
    "count":count,
    "uncount":uncount,
    "recount":recount,
    "macro":run_macro
}

def run_command(scope:ksd.KSSearchResult|None,command_str:str):
    try:
        parsed_commands = parse_command_str(command_str)
        results = []
        for cmd_group in parsed_commands:
            command,switches,options,values = cmd_group
            print(f"[{command},{switches},{options},{values}]")
            rval = command_list[command](scope,switches,options,values)
            results.append((command,rval))
        return (True,results)
    except Exception as e:
        return (False,command,repr(e))
    
if __name__ == "__main__":
    result = (True,"",None)
    while result[1] != "quit":
        line = input(">").lower()
        result = run_command(None, line)
        if result[1] == "quit":
            continue

        if result[0]:
            print(f'Succesfully run command "{result[1]}"...')
            
            if type(result[2]) == ksd.KSSearchResult:
                print(repr(result[2]))
            elif type(result[2]) == dict:
                for key,value in result[2].items():
                    print(key)
                    for v in value:
                        print(f"\t{v}")
            elif type(result[2]) == list:
                for value in result[2]:
                    print(repr(value))
            else:
                print(result[2])
        else:
            print(f'Error running command "{result[1]}": {result[2]}')