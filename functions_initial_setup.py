import re
import os
#import json
import orjson
import sys
from pathlib import Path
from class_network_device import NetworkDevice

from functions_helpers import (
    rw_cell,
    add_xls_tag,
    get_xls_sheet,
    mod_dir_based_on_os,
    verify_path,
    next_available_row,
    get_current_time,
    load_device_type_cache,
    save_device_type_cache,
)

#moved
####Initial Setup Fucntions
def read_settings_sheet(ws_obj, start_row=5,end_row=15):
    """
    Read all the Settings from the Settings sheet and generate the appropriate Dictionary.
    """
    t_dict={}
    for i in range(start_row, end_row+1):
        prompt = rw_cell(ws_obj, i, 1).lower()
        t_dict[prompt] = rw_cell(ws_obj, i, 3)
        if isinstance(t_dict[prompt], int):
            continue
        elif t_dict[prompt] in ["Yes", "YES", "yes"]:
            t_dict[prompt] = True
        else:
            t_dict[prompt] = False
    return t_dict

def read_global_variables(ws_obj):
    t_dict = {}
    t_dict["username"] = rw_cell(ws_obj, 1, 2)
    t_dict["password"] = rw_cell(ws_obj, 2, 2)
    t_dict["secret"] = rw_cell(ws_obj, 3, 2)
    if not t_dict["secret"]:
        t_dict["secret"] = t_dict["password"]
    t_dict["output_dir"] = verify_path(rw_cell(ws_obj, 4, 2))
    t_dict["output_file"] = rw_cell(ws_obj, 5, 2)
    if t_dict["output_file"]:
        t_dict["output_file"] = add_xls_tag(t_dict["output_file"])
    else:
        print("Error:\tNo Output file value was entered.")
        print("\tPlease enter an output file name and try again.")
        sys.exit()
    return t_dict

def cell_iter_to_list(cell_iter, ignore_empty_cell):
    t_list=[]
    for cell in cell_iter:
        if not ignore_empty_cell:
            t_list.append(cell.value)
        elif cell.value:
            t_list.append(cell.value)
    return t_list

def update_with_cli_args(glbl_var, cli_arg):
    for key, val in cli_arg.items():
        if key in list(glbl_var.keys()) and val:
            if key =="output_dir":
                glbl_var[key] = verify_path(val)
            else:
                glbl_var[key] = val

def get_setup_vars(wb_obj, cli_arg):
    """
    Reads all the setup variables from the XLS document.
    """
    return_dict = {}
    t_list = []
    t_dict = {}
    # Gathers the Commands to capture
    return_dict["other_commands"] = cell_iter_to_list(wb_obj["Commands"]["A"], True)
    # Gathers the Settings, on which functions to do.
    return_dict["settings"] = read_settings_sheet(wb_obj["Settings"])
    # Adds all the global parameters
    t_glbl_dict =read_global_variables(wb_obj["Main"])
    ##Update override with any CLI Arguments
    update_with_cli_args(t_glbl_dict, cli_arg)
    return_dict["global"] = t_glbl_dict
    return return_dict

def read_network_devices(wb_obj, dflt_creds):
    """
    Takes WB and checks the Main to create NetworkDevices and return
    a list of Network_Devices.

    De-dupes based on Main!A (Host). Normalizes by strip + lowercase.
    Keeps first instance; marks duplicates as Active="Duplicate" and logs to Errors.
    """
    sheet_obj = get_xls_sheet(wb_obj, "Main")
    err_sheet = get_xls_sheet(wb_obj, "Errors")

    return_list = []
    seen = {}  # norm_host -> first_row
    
    cache = load_device_type_cache()

    for i in range(8, sheet_obj.max_row + 1):
        host = rw_cell(sheet_obj, i, 1)
        if not host:
            continue

        # normalize host for dedupe
        norm_host = str(host).strip().lower()
        if not norm_host:
            continue

        # If already seen, mark duplicate + log, then skip
        if norm_host in seen:
            # Mark row as duplicate so it won't run
            rw_cell(sheet_obj, i, 2, True, "Duplicate")

            # Log to Errors sheet
            row = next_available_row(err_sheet)
            rw_cell(err_sheet, row, 1, True, str(host).strip())
            rw_cell(err_sheet, row, 2, True, f"Duplicate host in Main (first at row {seen[norm_host]}, duplicate at row {i})")
            rw_cell(err_sheet, row, 3, True, get_current_time("dt"))
            continue

        seen[norm_host] = i

        active = rw_cell(sheet_obj, i, 2)
        if not active:
            active = "Yes"

        # If row is not meant to run, still create device object (original behavior),
        # but duplicates will never reach here.
        parse_method = rw_cell(sheet_obj, i, 3)
        if not parse_method:
            parse_method = "autodetect"

        # If autodetect, try cache (key by normalized host string)
        if str(parse_method).lower() == "autodetect":
            #cached = type_cache.get(norm_host)  # norm_host already computed
            cached = cache.get(norm_host)
            if cached:
                parse_method = cached

        protocol = rw_cell(sheet_obj, i, 4)
        port_override = rw_cell(sheet_obj, i, 5)
        user_name = rw_cell(sheet_obj, i, 6)
        user_pass = rw_cell(sheet_obj, i, 7)

        if not user_name:
            user_name = dflt_creds["username"]
        if not user_pass:
            user_pass = dflt_creds["password"]

        net_device = NetworkDevice(
            str(host).strip(),
            user_name,
            user_pass,
            dflt_creds["secret"],
            parse_method,
            protocol,
            port_override,
            active
        )
        net_device.main_col = i
        return_list.append(net_device)

    return return_list

def update_ntc_templ_path():
    """
    Adds 'NET_TEXTFSM' variable to the os.environ, required to utilize
    textfsm with Netmiko.
    """
    ntc_dir = "ntc-templates/templates"
    os.environ["NET_TEXTFSM"] = str(Path(os.getcwd())/Path(ntc_dir))

def get_other_shows(wb_obj):
    """
    Obtains the "show " commands from the WB, returns dict with Commands
    as keys and None for values.
    """
    wb_sheet = get_xls_sheet(wb_obj, "Commands")
    return_dict = {}
    for cell in wb_sheet["A"]:
        if cell.value:
            return_dict[(cell.value)] = None
    return return_dict

def get_json_data_from_file(file_name):
    """
    Reads a json file and imports all the values.
    """
    with open(file_name, "rb") as file:
        #return_value = json.load(file)
        return_value = orjson.loads(file.read())
    return return_value