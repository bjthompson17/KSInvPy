import tkinter as tk
from tkinter import simpledialog
from tkinter import messagebox
import win32print as wp
import win32gui,win32con,winsound,os
import regex as re
from datetime import date, datetime
from typing import Literal,Any
import KSData as KSD
import KSInventoryApp as KSIA

# TODO: add import options screen to allow customization of importing hidden and service numbers

# TODO: add option to export either only on screen or all items

# TODO: add cursor selection support. Relevant example code below:
# text_widget.mark_set(tk.INSERT, "1.0")  # Set cursor position to line 1, column 0
# position = text_widget.index(tk.INSERT) # Get cursor position. Returns line.column
#
# More info on text indicies: https://anzeljg.github.io/rin2/book2/2405/docs/tkinter/text-index.html

# TODO: Make manual fixes easier (half done)

# TODO: make 'forget' command update what inventory gets imported. Add import filter to ignore numbers in range for faster importing

# TODO: Add checkbox to enable/disable RegEx searching

# TODO: Add UI configuration for product code or stock # 
# in export variance.

filter_results:KSD.KSSearchResult = KSD.KSSearchResult()

error_frequency = 1500  # frequency in Hertz
error_duration = 500  # duration in ms

window = tk.Tk()
window.title("Inventory Bikes")
window.geometry("900x500")

menu_bar = tk.Menu(window, tearoff = 0)
control_frame = tk.Frame(window)
text_display = tk.Text(window, state = tk.DISABLED, wrap = 'none', font=("Consolas",12))
vsb = tk.Scrollbar(window, orient='vertical',command = text_display.yview)
text_display["yscrollcommand"] = vsb.set
text_display.tag_config("current_line",background="#797979",foreground='white')


class PrintWindow(simpledialog.Dialog):
    def __init__(self, parent, title = None):
        self.printer_names = []
        self.printers_offline:list[int] = []
        idx = 0
        for x in wp.EnumPrinters(wp.PRINTER_ENUM_LOCAL,None,2): 
            pName = x['pPrinterName']
            driver_name = x["pDriverName"]
            online = (x["Attributes"] & wp.PRINTER_ATTRIBUTE_WORK_OFFLINE) == 0
            if driver_name != None and ("ZDesigner" in driver_name or "Generic / Text Only" in driver_name):
                self.printer_names.append(pName)
                if not online:
                    self.printers_offline.append(idx)
            idx += 1
            

        # datatype of menu text
        self.selected_printer = tk.StringVar()
        if len(self.printer_names) <= 0:
            print("No Valid printers")
            self.printer_names.append("None")
            self.selected_printer.set("None")
        elif wp.GetDefaultPrinter() in self.printer_names:
            self.selected_printer.set( wp.GetDefaultPrinter() )
        else :
            self.selected_printer.set( self.printer_names[0] )
            
        self.dpi = tk.IntVar()
        self.language = tk.StringVar()

        def on_printer_change(a,b,c):
            select_str = self.selected_printer.get().lower()
            if("300dpi" in select_str):
                self.dpi.set(300)
            elif("203dpi" in select_str):
                self.dpi.set(203)
            if("zpl" in select_str):
                self.language.set("ZPL")
            elif("epl" in select_str):
                self.language.set("EPL")
        self.selected_printer.trace_add("write",on_printer_change)
        on_printer_change("","","")
        
        self.print_items:KSD.KSSearchResult = KSIA.find(filter_results)
        count = 0
        extra = 0
        self.item_list = ""
        for key,value in self.print_items.items():
            for sn in value:
                count += 1
                if count < 10:
                    self.item_list += sn.serial_num + "\n"
                else:
                    extra += 1
        if extra > 0:
            self.item_list += f"...[{extra} more items]"
        self.item_list = self.item_list.strip()
        simpledialog.Dialog.__init__(self,parent,title)
    
    def body(self,master:tk.Frame):
        tk.Label(master,text = "Select Printer:")
        self.printer_menu:tk.Menu = tk.OptionMenu( master , self.selected_printer , *(self.printer_names) )
        # for i in self.printers_offline:
        #     self.printer_menu.entryconfig(i,state="disabled")
        settings_frame = tk.Frame(master)
        self.printer_dpi = tk.OptionMenu(settings_frame,self.dpi,203,300)
        self.printer_lang = tk.OptionMenu(settings_frame,self.language,"EPL","ZPL")
        self.printer_menu.pack()
        settings_frame.pack()
        self.printer_dpi.pack(side = tk.LEFT)
        self.printer_lang.pack(side = tk.LEFT)
        tk.Label(master,text=self.item_list).pack()
        return self.printer_menu
    
    def apply(self):
        confirm = messagebox.askyesno(
            "Confirm Print items", 
            "Are you sure you would like to print labels for the following items?:\n" 
            + self.item_list
        )
        
        if confirm:
            self.result = self.selected_printer.get()
            KSIA.print_labels(self.result,self.dpi.get(),self.language.get(),self.print_items)

class GetInvWindow(simpledialog.Dialog):
    # Should be Configurable
    STORES = ["MAIN","CSU","WAREHOUSE"]
    def __init__(self, parent, title = None, store_num = 0):
        self.store_name = tk.StringVar(value = self.STORES[store_num])
        self.result = None
        self.shared_presets = {}
        self.local_presets = {}
        if os.path.exists(KSIA.SHARED_CONFIG_PATH + "\\inv_range_presets.json"):
            self.shared_presets = KSIA.load_config(KSIA.SHARED_CONFIG_PATH + "\\inv_range_presets.json")
        else:
            with open(KSIA.SHARED_CONFIG_PATH + "\\inv_range_presets.json", "w") as file:
                file.write(r"{}")
        
        if os.path.exists(KSIA.LOCAL_CONFIG_PATH + "\\inv_range_presets.json"):
            self.local_presets = KSIA.load_config(KSIA.LOCAL_CONFIG_PATH + "\\inv_range_presets.json")
        else:
            with open(KSIA.LOCAL_CONFIG_PATH + "\\inv_range_presets.json", "w") as file:
                file.write(r"{}")
        self.list_options = list(self.local_presets.keys()) + list(self.shared_presets.keys())
        simpledialog.Dialog.__init__(self, parent, title)
    def body(self, master):
        store_frame = tk.Frame(master)
        tk.Label(store_frame, text="Store:").pack(side="left")
        tk.OptionMenu(store_frame,self.store_name,*self.STORES).pack(side="left")
        store_frame.pack()
        
        ranges = KSIA.inv_ranges.copy()
        if len(ranges) < 1:
            # Should be Configurable
            ranges = [("11000","19999")]

        tk.Label(master, text="Please enter the range to import, in the form: [MIN1]..[MAX1](,[MINn]..[MAXn])...").pack()
        self.range_entry = tk.Entry(master)
        self.range_entry.insert(tk.END,",".join("..".join(x) for x in ranges))
        self.range_entry.select_range(0,tk.END)
        self.range_entry.pack(expand=True, fill=tk.X)
        
        tk.Label(master, text = "Presets:").pack()
        self.preset_list = tk.Listbox(master,selectmode=tk.SINGLE,height=10,)
        self.preset_list.insert(tk.END, *self.local_presets.keys())
        self.preset_list.insert(tk.END, *self.shared_presets.keys())
        self.preset_list.pack(expand = True,fill = 'x')
        
        self.search_entry = tk.Entry(master)
        self.search_entry.pack()
        def preset_filter(event:tk.Event = None):
            regex_flags = re.BESTMATCH|re.IGNORECASE
            search_str = self.search_entry.get()
            self.preset_list.delete(0, tk.END)
            for preset in self.local_presets.keys():
                if re.search(search_str,preset,flags=regex_flags) is not None:
                    self.preset_list.insert(tk.END, preset)
            for preset in self.shared_presets.keys():
                if re.search(search_str,preset,flags=regex_flags) is not None:
                    self.preset_list.insert(tk.END, preset)
        self.search_entry.bind("<Key>", lambda e: self.search_entry.after(1,preset_filter,e))
        self.bind("<Control-f>",lambda e: self.search_entry.focus_set())
        self.bind("<Control-F>",lambda e: self.search_entry.focus_set())
        
        def on_preset_select(event):
            selection = self.preset_list.curselection()
            if len(selection) <= 0: return
            sel = selection[0]
            if self.range_entry.select_present():
                self.range_entry.delete(tk.SEL_FIRST, tk.SEL_LAST)
            self.range_entry.insert(tk.END,("," if len(self.range_entry.get()) > 0 and self.range_entry.get()[-1] != ',' else "") + (self.shared_presets[event.widget.get(sel - len(self.local_presets))] if selection[0] >= len(self.local_presets) else self.local_presets[event.widget.get(sel)]))
            
        self.preset_list.bind("<<ListboxSelect>>",on_preset_select)
        return self.range_entry

    def apply(self):
        pass
    
    def validate(self):
        store_num = self.STORES.index(self.store_name.get())
        self.result = (self.range_entry.get(),store_num)
        return 1

class ItemDetailWindow(simpledialog.Dialog):
    def __init__(self, parent, title = None, result:tuple[KSD.KSItem,list[KSD.KSSerializedItem]] = None):
        if(result == None or len(result) < 2):
            self.item = KSD.KSItem(0)
            self.sns = []
        else:
            self.item = result[0] #item
            self.sns = result[1]
        simpledialog.Dialog.__init__(self, parent, title)
    def body(self, master):
        id_frame = tk.Frame(master)
        tk.Label(id_frame, text=f"Product Code: {self.item.prod_code}").pack(side='left')
        tk.Label(id_frame, text=f"ID: {self.item.id}").pack(side='right')
        id_frame.pack(expand=True,fill='x')
        
        money_frame = tk.Frame(master)
        tk.Label(money_frame, text = f"Avg. Cost: ${self.item.cost}").pack(side='left')
        tk.Label(money_frame, text = f"Base Price: ${self.item.retail}").pack(side='right')
        money_frame.pack(expand=True,fill='x')
        
        desc_frame = tk.Frame(master)
        tk.Label(desc_frame, text=f"Desc: {self.item.desc}").pack(side='left')
        desc_frame.pack(expand=True,fill='x')
        
        self.phys_count = tk.DoubleVar(value = self.item.phys)
        count_frame = tk.Frame(master)
        tk.Label(count_frame, text = f"QOH: {self.item.qoh} ").pack(side='left')
        phys_entry = tk.Entry(count_frame, textvariable=self.phys_count)
        phys_entry.pack(side="right")
        tk.Label(count_frame, text = "Physical Count: ").pack(side='right')
        count_frame.pack(expand=True,fill='x')
        self.serial_list = None
        
        
        
        if(not self.item.serialized):
            self.bind("<Alt-=>", lambda e: self.changeCount(1))
            self.bind("<Alt-minus>", lambda e: self.changeCount(-1))
            self.bind("<Alt-+>", lambda e: self.changeCount(10))
            self.bind("<Alt-_>", lambda e: self.changeCount(-10))
            self.bind("<Alt-c>", lambda e: phys_entry.focus_set())
            self.bind("<Alt-C>", lambda e: phys_entry.focus_set())
        else:
            phys_entry.config(state="disabled")
            
            self.serial_list = tk.Listbox(master,background="#999", foreground="#444", selectbackground="#ddd", 
                                          selectforeground="black", selectmode=tk.MULTIPLE)
            for sn in self.sns:
                self.serial_list.insert(tk.END,sn.serial_num)
                if(sn.active):
                    self.serial_list.selection_set(tk.END)
            self.serial_list.pack(fill='x')
            return self.serial_list
        
        return phys_entry
    
    def changeCount(self, amount):
        self.phys_count.set(self.phys_count.get()+amount)
    
    def apply(self):
        if(not self.item.serialized):
            self.item.phys = self.phys_count.get()
        elif(self.serial_list != None):
            for i,sn in enumerate(self.sns):
                if self.serial_list.select_includes(i):
                    sn.restore()
                else: 
                    sn.remove()
    
    # def validate(self):
    #     return 1
     
# Overloading commands 
def my_filter(scope,switches,options,values):
    global filter_results
    id_sb.delete(0,tk.END)
    try: id_sb.insert(tk.END,switches["id"])
    except KeyError: pass

    nid_sb.delete(0,tk.END)
    try: nid_sb.insert(tk.END,switches["nid"])
    except KeyError: pass

    prod_code_sb.delete(0,tk.END)
    try:prod_code_sb.insert(tk.END,switches["item"])
    except KeyError: pass

    nprod_code_sb.delete(0,tk.END)
    try: nprod_code_sb.insert(tk.END,switches["nitem"])
    except KeyError: pass

    desc_sb.delete(0,tk.END)
    try: desc_sb.insert(tk.END,switches["desc"])
    except KeyError: pass

    ndesc_sb.delete(0,tk.END)
    try: ndesc_sb.insert(tk.END,switches["ndesc"])
    except KeyError: pass

    serial_sb.delete(0,tk.END)
    try: serial_sb.insert(tk.END,switches["sn"])
    except KeyError: pass
    try: serial_sb.insert(tk.END,switches["serial"])
    except KeyError: pass

    nserial_sb.delete(0,tk.END)
    try: nserial_sb.insert(tk.END,switches["nsn"])
    except KeyError: pass
    try: nserial_sb.insert(tk.END,switches["nserial"])
    except KeyError: pass

    flags_sb.delete(0,tk.END)
    try: flags_sb.insert(tk.END,KSIA.qp_escape(switches["flags"]))
    except KeyError: pass

    nflags_sb.delete(0,tk.END)
    try: nflags_sb.insert(0,KSIA.qp_escape(switches["nflags"]))
    except KeyError: pass
    
    eval_sb.delete(0,tk.END)
    try: eval_sb.insert(0,switches["eval"])
    except KeyError: pass
    
    bool_new_search_op.set(not ("nnew" in options))
    bool_removed_search_op.set(not ("nremoved" in options))
    bool_counted_search_op.set(not ("ncounted" in options))
    bool_service_search_op.set(not ("nservice" in options))

    filter_results = KSIA.filter(None, switches, options, values)
    return filter_results
KSIA.command_list["filter"] = my_filter

def my_find(scope,switches,options,values):
    results:KSD.KSSearchResult = KSIA.find(scope,switches,options,values)
    current_tag = text_display.tag_names("@0,0")[0]
    current_tag_pos = text_display.tag_ranges(f"{current_tag}")[0]
    
    next_type = "IT"
    next_id = results.get_items()[0].id
    
    for item,sns in results.items():
        if not item.serialized:
            next_tag_pos = text_display.tag_ranges(f"IT{item.id}")[0]
            if next_tag_pos > current_tag_pos:
                next_type = "IT"
                next_id = item.id
                break
            continue
        for sn in sns:
            next_tag_pos = text_display.tag_ranges(f"SN{sn.id}")[0]
            if next_tag_pos > current_tag_pos:
                next_type = "SN"
                next_id = sn.id
                break
        else:
            continue
        break
    
    try:
        tag_pos = text_display.tag_ranges(f"{next_type}{next_id}")[0]
        text_display.see(tag_pos)
        text_display.yview_scroll(text_display.dlineinfo(tag_pos)[1], 'pixels' )
    except IndexError:
        print("Not found")
        pass
KSIA.command_list["find"] = my_find

def my_forget(scope,switches:dict,options:set,values:list):
    options.add("item_only")
    forget_items:KSD.KSSearchResult = KSIA.find(None,switches,options,values)
    item_list = ",".join(f"{x.prod_code:>6}" for x in forget_items.keys())

    confirm = messagebox.askyesno(
        "Confirm forget items", 
        "Are you sure you would like to forget the following items?:\n" 
        + item_list
    )
    
    if confirm:
        return KSIA.forget(forget_items)
    else:
        return []
KSIA.command_list["forget"] = my_forget

def my_get_inv(scope,switches:dict,options:set,values:list):
    current_store = KSIA.current_store
    if "store" in switches:
        if switches["store"] != current_store:
            file_clear()
    
    if "range" not in switches:
        inv_win = GetInvWindow(window,"Get Inventory",store_num=current_store)
        if inv_win.result is None:
            return
        switches["range"] = inv_win.result[0]
        if "store" not in switches:
            if inv_win.result[1] != current_store:
                file_clear()
            switches["store"] = inv_win.result[1]
    KSIA.get_inv(scope,switches,options,values)
KSIA.command_list["get_inv"] = my_get_inv

def handleArrowKeys(event:tk.Event = None):
    focused_element = window.focus_get()
    
    if focused_element is None or isinstance(focused_element,tk.Text):
        return
        
    if focused_element.master.master == None: 
        return
    
    next_element = focused_element

    if event.keysym == "Right":
        next_element = focused_element.tk_focusNext()
    elif event.keysym == "Left":
        next_element = focused_element.tk_focusPrev()
    else:
        current_frame = focused_element.master
        if focused_element.winfo_manager() == "grid":
            x = focused_element.winfo_x() + 1
            y = focused_element.winfo_y()
            grid_pos = current_frame.grid_location(x,y)
            current_input_index = current_frame.grid_slaves(grid_pos[1],grid_pos[0]).index(focused_element)
            if event.keysym == "Up":
                if grid_pos[1] > 0:
                    next_list = current_frame.grid_slaves(grid_pos[1]-1,grid_pos[0])
                    if len(next_list) > 0:
                        next_element = next_list[current_input_index]
            elif event.keysym == "Down" :
                if grid_pos[1] < current_frame.grid_size()[1] - 1:
                    next_list = current_frame.grid_slaves(grid_pos[1]+1,grid_pos[0])
                    if len(next_list) > 0:
                        next_element = next_list[current_input_index]

        if next_element == focused_element:
            if event.keysym == "Up":
                next_element = current_frame.winfo_children()[0].tk_focusPrev()
            elif event.keysym == "Down":
                next_element = current_frame.winfo_children()[-1].tk_focusNext()
            
            next_frame = next_element.master
            
            if next_element.winfo_manager() == "grid":
                x = next_element.winfo_x() + 1
                y = next_element.winfo_y()
                grid_pos = next_frame.grid_location(x,y)
                i_range = range(0,next_frame.grid_size()[1])
                for i in i_range:
                    slaves = next_frame.grid_slaves(grid_pos[1],i)
                    if len(slaves) > 0:
                        i_element = slaves[0]
                        if i_element["takefocus"] != 0:
                            next_element = i_element
                            break

            elif next_element.winfo_manager() == "pack":

                slaves = next_frame.winfo_children()
                i_range = []
                if next_element.pack_info()["side"] == 'left':
                    i_range = range(0,len(slaves))
                elif next_element.pack_info()["side"] == 'right':
                    i_range = range(len(slaves) - 1,-1,-1)
                
                for i in i_range:
                    i_element = slaves[i]
                    if i_element["takefocus"] != 0:
                        next_element = i_element
                        break
    if next_element["takefocus"] == 0 or next_element["takefocus"] == "0":
        next_element = next_element.tk_focusNext()
    next_element.focus_set()

window.bind('<Alt-Left>',handleArrowKeys)
window.bind('<Alt-Right>',handleArrowKeys)
window.bind('<KeyPress-Up>',handleArrowKeys)
window.bind('<KeyPress-Down>',handleArrowKeys)

def handleCursor(event:tk.Event = None):
    element = event.widget
    if not isinstance(element,tk.Text): return
    newCursorRow = int(element.index(tk.INSERT).split('.')[0]) # index in form `row.column`, starts with `1.0`
    element.mark_set(tk.INSERT,str(newCursorRow)+'.1')
    element.tag_remove('current_line','1.0','end')
    element.tag_add('current_line','insert linestart','insert lineend + 1 chars')


text_display.bind('<Key>', lambda e: text_display.after(1,handleCursor,e))
text_display.bind('<Button-1>', lambda e: text_display.after(1,handleCursor,e))
text_display.bind('<FocusIn>', lambda e: text_display.tag_config('current_line',background="blue"))
text_display.bind('<FocusOut>', lambda e: text_display.tag_config('current_line',background="#797979"))

def update_screen():
    current_cursor_idx = text_display.index(tk.INSERT)
    def blend_color(color1:tuple[int,int,int]|None, color2:tuple[int,int,int]|str|None = None, percent_blend = 0.5):
        if color2 == None:
            color2 = color1
        elif type(color2) == str:
            if re.fullmatch(r'#[\da-f]{3}',color2,flags=re.IGNORECASE):
                color2 = tuple(int(color2[i+1]*2,16) for i in (0,1,2))
            elif re.fullmatch(r'#[\da-f]{6}',color2,flags=re.IGNORECASE):
                color2 = tuple(int(color2[i+1:i+3],16) for i in (0,2,4))
            else:
                color2 = color1

        if color1 == None:
            color1 = color2
        elif type(color1) == str:
            if re.fullmatch(r'#[\da-f]{3}',color1,flags=re.IGNORECASE):
                color1 = tuple(int(color1[i+1]*2,16) for i in (0,1,2))
            elif re.fullmatch(r'#[\da-f]{6}',color1,flags=re.IGNORECASE):
                color1 = tuple(int(color1[i+1:i+3],16) for i in (0,2,4))
            else:
                color1 = color2
        
        if color1 == None or (color1 == (0,0,0) and color2 == (0,0,0)):
            return (0,0,0)
        pcolors = ((color1[0]/255,color2[0]/255),(color1[1]/255,color2[1]/255),(color1[2]/255,color2[2]/255))
        pcmax = [0,0,0]
        
        def psum(_pca, _pcb):
            return _pca * (1-percent_blend) + _pcb * percent_blend
        
        for pca,pcb in pcolors:
            if psum(pca,pcb) > pcmax[2]: pcmax[2] = psum(pca,pcb)
            if pca > pcmax[0]: pcmax[0] = pca
            if pcb > pcmax[1]: pcmax[1] = pcb
        bright_a = pcmax[0]
        bright_b = pcmax[1]

        return tuple(int(psum(a, b)/pcmax[2] * psum(bright_a, bright_b) * 255) for a,b in pcolors)
    def color_to_hex(color:tuple[int,int,int]):
        result = "#"
        for i in color:
            result += "%02x" % i
        return result
    
    scroll_pos = text_display.index("@0,0")
    text_display.config(state=tk.NORMAL)
    text_display.delete('0.0',tk.END)
    
    sn_count = 0
    item_count = 0
    phys_count = 0
    missing_count = 0
    
    total_cost = 0.0
    total_retail = 0.0
    
    phys_cost = 0.0
    phys_retail = 0.0
    for key,value in filter_results.items():
        item_count += 1
        phys_count += key.phys
        
        text_display.insert(tk.END,repr(key) + '\n', (f'IT{key.id}'))
        
        percent_full = 1.0
        percent_over = 0.0
        if key.qoh > 0:
            percent_full = (key.phys / key.qoh)
        if percent_full > 1: 
            percent_full = 1.0
            percent_over = (key.phys - key.qoh) / key.qoh
            if percent_over > 1:
                percent_over = 1.0
        if percent_full < 0: 
            percent_full = 0.0
        item_bg_color = None
        if percent_over <= 0:
            item_bg_color = blend_color("#A66","#6A6", percent_full)
        else:
            item_bg_color = blend_color("#6A6","#A6A", percent_over)
        
        fg_color = "#000"
        bg_color = "#555"
        if item_bg_color != None:
            if max(item_bg_color[0],item_bg_color[1],item_bg_color[2]) < 90:
                fg_color = "#ccc"
            bg_color = color_to_hex(item_bg_color)
            
        text_display.tag_config(f'IT{key.id}',background=bg_color,foreground=fg_color,selectforeground="black",selectbackground="green")

        total_cost += key.cost * key.qoh
        total_retail += key.retail * key.qoh
        
        phys_cost += key.cost * key.phys
        phys_retail += key.retail * key.phys
        
        missing_count += key.qoh - key.phys
        
        if key.serialized:
            item_sn_count = 0
            sn = ""
            v = 0
            for i,sn1 in enumerate(key.serial_nums):
                if(v >= len(value)): break
                if(sn1.id == value[v].id):
                    sn=value[v]
                    v += 1
                else:
                    continue

                text_display.insert(tk.END,str(i+1)+'-\t'+repr(sn) + '\n', (f'SN{sn.id}'))
                
                item_sn_count += 1
                sn_count += 1

                # Configurable colors?
                # "white", "black", "red", "green", "blue", "cyan", "yellow", and "magenta"
                
                bg_color = "#ddd"
                fg_color = "black"
                sbg_color = "blue"
                sfg_color = "white"
                color_blend = None

                if not sn.active:
                    color_blend = blend_color(color_blend,"#999")
                    fg_color = "#444"
                
                if len(sn.flags) > 0:
                    bg_color = "#39d"
                for flag in sn.flags:
                    if flag.lower() == "red":
                        color_blend = blend_color(color_blend,"#b33")
                    elif flag.lower() == "green":
                        color_blend = blend_color(color_blend,"#3b3")
                    elif flag.lower() == "purple":
                        color_blend = blend_color(color_blend,"#93e")
                    elif flag.lower() == "blue":
                        color_blend = blend_color(color_blend,"#33c")
                    elif flag.lower() == "cyan":
                        color_blend = blend_color(color_blend,"#3cc")
                    elif flag.lower() == "yellow":
                        color_blend = blend_color(color_blend,"#cc3")
                    elif flag.lower() == "magenta":
                        color_blend = blend_color(color_blend,"#c3c")
                    elif flag.lower() == "pink":
                        color_blend = blend_color(color_blend,"#f09")
                    elif flag.lower() == "black":
                        color_blend = blend_color(color_blend,"#000")
                    elif flag.lower() == "white":
                        color_blend = blend_color(color_blend,"#fff")
                    elif re.fullmatch(r'#(?:[\da-f]{3}){1,2}',flag,flags=re.IGNORECASE):
                        color_blend = blend_color(color_blend,flag)
                if color_blend != None:
                    if max(color_blend[0],color_blend[1],color_blend[2]) < 90:
                        fg_color = "#ccc"
                    bg_color = color_to_hex(color_blend)
                
                try:
                    text_display.tag_config(f'SN{sn.id}',background=bg_color,foreground=fg_color, selectbackground=sbg_color, selectforeground=sfg_color)
                except tk.TclError:
                    print("Color error:",bg_color)
                    text_display.tag_config(f'SN{sn.id}',background="white",foreground="black", selectbackground="blue", selectforeground="white")
    
    text_display.mark_set(tk.INSERT,current_cursor_idx)
    text_display.tag_add('current_line','insert linestart','insert lineend + 1 chars')
    text_display.tag_raise('current_line')
    text_display.config(state=tk.DISABLED)
    result_count.config(text = f"Store: {GetInvWindow.STORES[KSIA.current_store]}\n{item_count} results: {sn_count} bikes, {phys_count} total items, {-1 * missing_count} variance \n"
                        f" Expected Cost: ${total_cost:.2f} \t Count Cost: ${phys_cost:.2f} \t Diff: ${phys_cost - total_cost:.2f}\n"
                        f" Expected Retail: ${total_retail:.2f} \t Count Retail: ${phys_retail:.2f} \t Diff: ${phys_retail - total_retail:.2f}")
    text_display.see(scroll_pos)
    line_info = text_display.dlineinfo(scroll_pos)
    if line_info is not None:
        text_display.yview_scroll(line_info[1], 'pixels' )

def update_filter(options:set[str] = set()):
    global filter_results
    switches = KSIA.pack_switches(
            id = id_sb.get(),
            nid = nid_sb.get(),
            item = prod_code_sb.get(),
            nitem = nprod_code_sb.get(),
            desc = desc_sb.get(),
            ndesc = ndesc_sb.get(),
            sn = serial_sb.get(),
            nsn = nserial_sb.get(),
            flags = flags_sb.get(),
            nflags = nflags_sb.get(),
            eval = eval_sb.get())
    if match_mode_option.get() == "exact":
        switches["errors"] = 0
    else:
        switches["errors"] = -1
    
    filter_options = options.copy()
    
    if not bool_new_search_op.get():
        filter_options.add("nnew")
    if not bool_removed_search_op.get():
        filter_options.add("nremoved")
    if not bool_counted_search_op.get():
        filter_options.add("ncounted")
    if not bool_service_search_op.get():
        filter_options.add("nservice")

    filter_results = KSIA.filter(None, switches, filter_options)
    update_screen()
    
def refresh_filter():
    global filter_results
    filter_results = KSIA.refresh()
    update_screen()

# ============ Menu Cascades =====================
file_menu = tk.Menu(menu_bar, tearoff = 0)

def file_clear():
    print("File->Clear")
    if messagebox.askokcancel("Clear Items", "Are you sure you want to clear all items?"):
        KSIA.reset()
        refresh_filter()
        return True
    return False
file_menu.add_command(label="Clear Items", command=file_clear)

def file_open():
    print("File->Open")
    if file_clear():
        KSIA.load()
        refresh_filter()
        return True
    return False
file_menu.add_command(label= "Open", command= file_open)

def file_merge():
    print("File->Merge")
    KSIA.load(options = set(["merge"]))
    refresh_filter()
file_menu.add_command(label= "Merge", command= file_merge)


# save_menu = tk.Menu(menu_bar, tearoff = 0)
def file_save():
    print("File->Save")
    KSIA.save()
file_menu.add_command(label= "Save As",command=file_save)
# file_menu.add_cascade(label = "Save...", menu = save_menu)

def file_get_inv():
    current_store = KSIA.current_store
    print("File->Get Inventory")
    inv_win = GetInvWindow(window, title="Get Inventory", store_num=current_store)
    if inv_win.result is None:
        return
    if inv_win.result[1] != current_store:
        file_clear()
    text_display.config(state=tk.NORMAL)
    text_display.delete('0.0',tk.END)
    text_display.insert(tk.END,chars = f"Retrieving Inventory ...\n")
    text_display.config(state=tk.DISABLED)
    text_display.update()
    KSIA.get_inv(None,KSIA.pack_switches(range = inv_win.result[0], store = inv_win.result[1]))
    refresh_filter()
file_menu.add_command(label = "Get Inventory  [F5]", command = file_get_inv)
window.bind('<F5>',lambda e: file_get_inv())

def file_start_count():
    file_get_inv()
    KSIA.recount()
    refresh_filter()
file_menu.add_command(label = "Start Count", command = file_start_count)

export_menu = tk.Menu(menu_bar,tearoff=0)

def export_vexport():
    print("Exports->Export Variance")
    text_display.config(state=tk.NORMAL)
    text_display.delete('0.0',tk.END)
    text_display.insert(tk.END,chars = f"Exporting Variance File ...\n")
    text_display.config(state=tk.DISABLED)
    text_display.update()
    KSIA.export_variance(switches={"file": KSIA.ks_path + "\\InventoryExports\\KSVARIANCE.LOG"})
    refresh_filter()
export_menu.add_command(label = "Export Variance", command = export_vexport)

def export_import():
    print("Exports->Import")
    KSIA.import_file()
    refresh_filter()
export_menu.add_command(label = "Import...", command = export_import)

menu_bar.add_cascade(label = "File", menu = file_menu)
menu_bar.add_cascade( label = "Exports", menu=export_menu)

utilities_menu = tk.Menu(menu_bar, tearoff=0)
def open_print_window():
    pw = PrintWindow(window,title="Print Bike Labels")
    return pw.result
utilities_menu.add_command(label = "Print Bike Labels ...", command=open_print_window)
menu_bar.add_cascade(label = "Utilities", menu = utilities_menu)

# ===============================================
# ============== Build Frames ===================
# ====== Search Frame ======
flexible_frame = tk.Frame(control_frame)
buttons_frame = tk.Frame(control_frame)
buttons_frame.pack(side = 'right', fill='y')
flexible_frame.pack(fill = 'x')

search_frame = tk.Frame(flexible_frame)
def update_search(e:tk.Event = None):
    if type(e.widget) == tk.Entry:
        e.widget.select_range(0,tk.END)
    update_filter()

tk.Label(search_frame, text = "Filter IDs:").grid(column=0,row = 0,sticky="w")
id_sb = tk.Entry(search_frame)
id_sb.bind("<Return>",update_search)
id_sb.grid(column=1,row = 0,sticky="ew")

tk.Label(search_frame, text = " not ").grid(column=2,row = 0)
nid_sb = tk.Entry(search_frame)
nid_sb.bind("<Return>",update_search)
nid_sb.grid(column=3,row = 0,sticky="ew")


tk.Label(search_frame,text="Filter Product Code:").grid(column=0,row = 1,sticky="w")
prod_code_sb = tk.Entry(search_frame)
prod_code_sb.bind("<Return>",update_search)
prod_code_sb.grid(column=1,row = 1,sticky="ew")

tk.Label(search_frame,text=" not ").grid(column=2,row = 1)
nprod_code_sb = tk.Entry(search_frame)
nprod_code_sb.bind("<Return>",update_search)
nprod_code_sb.grid(column=3,row = 1,sticky="ew")


tk.Label(search_frame,text="Filter Item Description: ").grid(column=0,row = 2,sticky="w")
desc_sb = tk.Entry(search_frame)
desc_sb.bind("<Return>",update_search)
desc_sb.grid(column=1,row = 2,sticky="ew")

tk.Label(search_frame,text=" not ").grid(column=2,row = 2)
ndesc_sb = tk.Entry(search_frame)
ndesc_sb.bind("<Return>",update_search)
ndesc_sb.grid(column=3,row = 2,sticky="ew")


tk.Label(search_frame,text = "Filter Serial Number: ").grid(column=0,row = 3,sticky="w")
serial_sb = tk.Entry(search_frame)
serial_sb.bind("<Return>",update_search)
serial_sb.grid(column=1,row = 3,sticky="ew")

tk.Label(search_frame,text = " not ").grid(column=2,row = 3)
nserial_sb = tk.Entry(search_frame)
nserial_sb.bind("<Return>",update_search)
nserial_sb.grid(column=3,row = 3,sticky="ew")


tk.Label(search_frame, text = "Filter Flags: ").grid(column=0,row = 4,sticky="w")
flags_sb = tk.Entry(search_frame)
flags_sb.bind("<Return>",update_search)
flags_sb.grid(column=1,row = 4,sticky="ew")

tk.Label(search_frame, text = " not ").grid(column=2,row = 4)
nflags_sb = tk.Entry(search_frame)
nflags_sb.bind("<Return>",update_search)
nflags_sb.grid(column=3,row = 4,sticky="ew")


tk.Label(search_frame, text = "Filter Eval: ").grid(column=0,row = 5,sticky="w")
eval_sb = tk.Entry(search_frame)
eval_sb.bind("<Return>",update_search)
eval_sb.grid(column=1,row = 5, columnspan = 3,sticky="ew")

search_frame.grid_columnconfigure(1,weight=1)
search_frame.grid_columnconfigure(3,weight=1)

bool_op_sb_frame = tk.Frame(search_frame)
bool_new_search_op = tk.BooleanVar(value = True)
bool_removed_search_op = tk.BooleanVar(value = True)
bool_counted_search_op = tk.BooleanVar(value = True)
bool_service_search_op = tk.BooleanVar(value = True)
checkbox = tk.Checkbutton(bool_op_sb_frame, text = "Added", variable=bool_new_search_op)
checkbox.bind("<Return>",update_search)
checkbox.pack(side="left")
checkbox = tk.Checkbutton(bool_op_sb_frame, text = "Removed", variable=bool_removed_search_op)
checkbox.bind("<Return>",update_search)
checkbox.pack(side="left")
checkbox = tk.Checkbutton(bool_op_sb_frame, text = "Counted", variable=bool_counted_search_op)
checkbox.bind("<Return>",update_search)
checkbox.pack(side="left")
checkbox = tk.Checkbutton(bool_op_sb_frame, text = "Service", variable=bool_service_search_op)
checkbox.bind("<Return>",update_search)
checkbox.pack(side="left")
bool_op_sb_frame.grid(column=1, row=6, columnspan=4, sticky="w")

tk.Button(search_frame,text="Search",command=update_filter, takefocus=0,bg="#3b3").grid(row = 7, column=0, columnspan=4,sticky="we")

# ==== Scan Frame ====

def scanner_count_item():
    item = '^"' + prod_code_scan_sb.get() + '"$'
    if item_only_scan_op.get():
        serial = serial_scan_sb.get()
    else:
        serial = '^"' + serial_scan_sb.get() + '"$'
    flags = global_ctrl_flags.get()
    switches = KSIA.pack_switches(item = item, sn = serial)
    if match_mode_option.get() == "exact":
        switches["errors"] = 0
    else:
        switches["errors"] = -1
    results:KSD.KSSearchResult = KSIA.find(None, switches, {"all","ncounted"})
    
    if len(results.keys()) < 1:
        winsound.Beep(error_frequency,error_duration)
        return
    
    quantity = 1
    try:
        quantity = float(scan_quantity_entry.get())
    except ValueError:
        pass
    
    sn_count = 0
    for key,value in results.items():
        if not key.serialized:
            if not item_only_scan_op.get():
                continue
            tag_range = text_display.tag_ranges(f"IT{key.id}")
            if tag_range is not None and len(tag_range) > 0:
                tag_pos = tag_range[0]
                text_display.see(tag_pos)
                text_display.yview_scroll(text_display.dlineinfo(tag_pos)[1], 'pixels' )
            key.increase_count(quantity)
            break
        for sn in value:
            if sn.active:
                continue
            sn.restore()
            sn_count += 1
            try:
                tag_range = text_display.tag_ranges(f"SN{sn.id}")
                if tag_range is not None and len(tag_range) > 0:
                    tag_pos = tag_range[0]
                    text_display.see(tag_pos)
                    text_display.yview_scroll(text_display.dlineinfo(tag_pos)[1], 'pixels' )
            except IndexError:
                pass
            
            if sn_count >= quantity:
                break
        else:
            continue
        break
    else:
        winsound.Beep(error_frequency,error_duration)
    refresh_filter()

def scan_submit(mode:str):
    if scan_filter_results.get():
        my_filter(None,{"item":prod_code_scan_sb.get(),"sn":serial_scan_sb.get()},set(),[])
    if mode == "item":
        if item_only_scan_op.get():
            scanner_count_item()
            prod_code_scan_sb.focus_set()
            prod_code_scan_sb.select_range(0,tk.END)
        else:
            serial_scan_sb.focus_set()
    elif mode == "sn":
        scanner_count_item()
        prod_code_scan_sb.focus_set()
        prod_code_scan_sb.select_range(0,tk.END)
        if not item_only_scan_op.get():
            serial_scan_sb.select_range(0,tk.END)
    if not scan_quantity_keep.get():
        scan_quantity_entry.delete(0,tk.END)

scan_frame = tk.Frame(flexible_frame)
tk.Label(scan_frame, text = "Scan Item: ").grid(row = 0,column=0,sticky="w")
prod_code_scan_sb = tk.Entry(scan_frame)
prod_code_scan_sb.bind('<Return>',lambda e: scan_submit("item"))
prod_code_scan_sb.grid(row = 0,column=1,sticky="ew")

tk.Label(scan_frame, text = "Scan Serial Number: ").grid(row = 1,column=0)
serial_scan_sb = tk.Entry(scan_frame)
serial_scan_sb.bind('<Return>',lambda e: scan_submit("sn"))
serial_scan_sb.grid(row = 1, column = 1,sticky="ew")

scan_option_frame = tk.Frame(scan_frame)
item_only_scan_op = tk.BooleanVar(value=True)
tk.Checkbutton(scan_option_frame,text="Item Only Scanning", variable=item_only_scan_op).pack(side = "left", anchor="w")

tk.Label(scan_option_frame,text = "Quantity:").pack(side = "left")
scan_quantity_entry = tk.Entry(scan_option_frame, width = 5)
scan_quantity_entry.bind('<Return>',lambda e: scan_submit("sn"))
scan_quantity_entry.pack(side = "left")

scan_quantity_keep = tk.BooleanVar(value = False)
tk.Checkbutton(scan_option_frame, text = "Keep Quantity", variable=scan_quantity_keep).pack(side="left")

scan_filter_results = tk.BooleanVar(value = False)
tk.Checkbutton(scan_option_frame, text = "Filter Results", variable = scan_filter_results).pack(side="left")

scan_option_frame.grid(row = 2, column = 0, columnspan=2, sticky="ew")

scan_frame.grid_columnconfigure(1,weight=1)
# ================= Global Control Frame =====================

global_ctrl_fm = tk.Frame(control_frame)
tk.Label(global_ctrl_fm, text = "Edit Flags: ").pack(side='left')
global_ctrl_flags = tk.Entry(global_ctrl_fm, width = 20)
def search_ctrl_set_flags():
    for result in filter_results.values():
        for sn in result:
            sn.set_flags(global_ctrl_flags.get().split(",") if len(global_ctrl_flags.get()) > 0 else [])
    update_screen()
def search_ctrl_remove_flags():
    for result in filter_results.values():
        for sn in result:
            if len(global_ctrl_flags.get()) <= 0:
                sn.clear_flags()
            else:
                sn.remove_flags(global_ctrl_flags.get().split(","))
    update_screen()
global_ctrl_flags.bind("<Return>",lambda e: search_ctrl_set_flags())
global_ctrl_flags.bind("<Shift-Return>",lambda e: search_ctrl_remove_flags())
global_ctrl_flags.pack(side="left")

match_mode_option = tk.StringVar(value="exact")
tk.Radiobutton(global_ctrl_fm,text="Exact Match", variable=match_mode_option, value = "exact").pack(side = "left")
tk.Radiobutton(global_ctrl_fm,text="Closest Match" , variable=match_mode_option, value = "closest").pack(side = "left")

global_ctrl_fm.pack(fill='x')

# =================================================
def clear_search_boxes():
    for widget in flexible_frame.winfo_children():
        widget.pack_forget()
pages = (("Search Mode",search_frame, serial_sb), ("Scan Mode",scan_frame, prod_code_scan_sb))
next_page = 1
def switch_mode(e:tk.Event = None):
    global next_page
    clear_search_boxes()
    pages[next_page][1].pack(fill='x')
    pages[next_page][2].focus_set()
    next_page = (next_page + 1) % len(pages)
    mode_button.config(text = pages[next_page][0])

mode_button = tk.Button(buttons_frame, text=pages[next_page][0],command=switch_mode,takefocus=0)
mode_button.pack(anchor="w")
window.bind("<Control-Tab>", switch_mode)


def run_command(command_str):
    seperated_command_pattern = re.compile(r'(?:[^&"]|"(?:\\[\\"]|[^"])*")+')
    for unique_command_str in re.findall(seperated_command_pattern,command_str):
        unique_command_str = unique_command_str.strip()
        if match_mode_option.get() != "exact" and ("/errors" not in unique_command_str) :
            unique_command_str += " /errors x"
        result = KSIA.run_command(filter_results,unique_command_str)
    refresh_filter()

command_frame = tk.Frame(window)
command_entry = tk.Entry(command_frame)
command_entry.bind("<Return>", lambda e: [command_entry.select_range(0,tk.END), run_command(command_entry.get())])
command_entry.pack(fill = 'x',expand = True)

command_frame.pack(side = 'bottom', fill = 'x')

result_count = tk.Label(control_frame, text = "")
result_count.pack()
search_frame.pack(fill = 'x')
control_frame.pack(side="top", fill='x')

vsb.pack(side = 'right', fill = 'y')
text_display.pack(fill="both",expand=True)
window.config(menu = menu_bar)

# ========================= Shortcuts =============================

# Search Frame
window.bind("<Alt-s>", lambda e: serial_sb.focus_set())
window.bind("<Alt-S>", lambda e: serial_sb.focus_set())
window.bind("<Alt-d>", lambda e: desc_sb.focus_set())
window.bind("<Alt-D>", lambda e: desc_sb.focus_set())
window.bind("<Alt-c>", lambda e: command_entry.focus_set())
window.bind("<Alt-C>", lambda e: command_entry.focus_set())
window.bind("<Alt-i>", lambda e: id_sb.focus_set())
window.bind("<Alt-I>", lambda e: id_sb.focus_set())
window.bind("<Alt-g>", lambda e: flags_sb.focus_set())
window.bind("<Alt-G>", lambda e: flags_sb.focus_set())
window.bind("<Alt-p>", lambda e: prod_code_sb.focus_set())
window.bind("<Alt-P>", lambda e: prod_code_sb.focus_set())

# Scan Frame
window.bind("<Alt-n>", lambda e: prod_code_scan_sb.focus_set())
window.bind("<Alt-N>", lambda e: prod_code_scan_sb.focus_set())
window.bind("<Alt-q>", lambda e: scan_quantity_entry.focus_set())
window.bind("<Alt-Q>", lambda e: scan_quantity_entry.focus_set())

# Global Control Frame
window.bind("<Alt-x>", lambda e: global_ctrl_flags.focus_set())
window.bind("<Alt-X>", lambda e: global_ctrl_flags.focus_set())

# Text area
window.bind("<Alt-t>", lambda e: text_display.focus_set())
window.bind("<Alt-T>", lambda e: text_display.focus_set())

def manage_item_details(e):
    current_line_ranges = text_display.tag_ranges('current_line')
    if len(current_line_ranges) > 0:
        tags = text_display.tag_names(current_line_ranges[0])
        if len(tags) > 1:
            item_tag = tags[-2]
            if item_tag[0:2] == "IT":
                item = KSIA.data.items[int(item_tag[2:])]
                ItemDetailWindow(window,result=(item,filter_results.result[item]))
            elif item_tag[0:2] == "SN":
                search_results = KSIA.find(filter_results,{'id':item_tag[2:]})
                ItemDetailWindow(window,result=search_results.get_first_result())
        text_display.focus_set()
        update_screen()

window.bind("<F3>", manage_item_details)

def on_closing():
    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        window.destroy()

window.protocol("WM_DELETE_WINDOW", on_closing)

update_filter()
window.mainloop()
