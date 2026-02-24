"""###v3.0### -CK
Original from: https://github.com/InsightSSG/GetInventory

Detecting devices is what takes time...
"""
import socket
from threading import Thread
from concurrent.futures import ThreadPoolExecutor, as_completed, wait, FIRST_COMPLETED
import traceback
import os
import os.path
#import json
import orjson
import re
from datetime import datetime
import time
import getopt
import sys
from pathlib import Path
from rich.console import Console
from rich.logging import RichHandler
import logging
import platform
import netmiko
import openpyxl
import optparse

from time import perf_counter

from class_network_device import NetworkDevice

from functions_gatherers import (
    set_verbose,
    gather_version,
    gather_arp,
    gather_mac,
    gather_interface,
    gather_cdp,
    gather_lldp,
    gather_route,
    gather_bgp,
    gather_inventory,
    gather_commands,
    count_interfaces
)
from functions_initial_setup import (
    update_ntc_templ_path,
    get_json_data_from_file,
    get_setup_vars,
    read_network_devices,
    save_device_type_cache,
    load_device_type_cache,
)
from functions_helpers import (
    find_val_in_col,
    gen_spacer,
    map_headers,
    next_available_row,
    format_uptime,
    left,
    right,
    mid,
    center_string,
    rw_cell,
    add_xls_tag,
    get_xls_sheet,
    open_xls,
    save_xls,
    mod_dir_based_on_os,
    verify_path,
    print_net_dev_msg,
    get_current_time,
)
from functions_device_runner import (connect_devices, con_thread, start_connection_log, connect_single_device, get_hostname)
import functions_device_runner as fdr

from functions_device_autodetect import (autodetect_devices, finalize_after_autodetect)
import functions_device_autodetect as fda

from functions_writers import (save_device_data, save_dev_show_json_data, write_dev_vars_to_wb, save_other_shows_to_txt, init_next_row_cache, gather_results_to_wb, add_err_msgs_to_wb, post_process_lldp_sheet_with_wtp_versions)
import functions_writers as fw

INPUT_FILE_NAME = "GetInventory - Default.xlsx"

"""###VERBOSE Output###"""
VERBOSE = False
VERBOSE_MORE = False
VERBOSITY_LEVEL = 0

console = Console()
#logger = logging.getLogger("inventory")  # optional name



RAW_CLI_OUTPUT = False

TESTING = False

"""###Global Variables###"""
GLBL_KEY_MAP = {}

def setup_logging(verbose: bool, tracebacks=True) -> logging.Logger:
    # 1) Make root logger not emit anything to console
    root = logging.getLogger()
    root.handlers.clear()
    root.setLevel(logging.WARNING)  # ignore debug/info from other libs

    # 2) Create handler for *only* your logger
    if tracebacks == True:
        handler = RichHandler(console=console, rich_tracebacks=True)
    else:
        handler = RichHandler(console=console, rich_tracebacks=False)
    handler.setLevel(logging.DEBUG if verbose else logging.INFO)

    # Optional: only show the raw message (no "INFO" prefix, etc.)
    formatter = logging.Formatter("%(message)s")
    handler.setFormatter(formatter)

    # 3) Configure your app logger
    inv = logging.getLogger("inventory")
    inv.handlers.clear()
    inv.addHandler(handler)
    inv.setLevel(logging.DEBUG if verbose else logging.INFO)

    # 4) Prevent it from propagating to root (avoids duplicates / root filtering surprises)
    inv.propagate = False

    return inv

def print_top_command_offenders(network_devices, top_n=15):
    """
    Aggregates net_dev.cmd_timings across devices and prints the slowest commands.
    """
    totals = {}   # command -> total seconds
    counts = {}   # command -> count
    maxes = {}    # command -> max single-run seconds

    for dev in network_devices:
        cmd_timings = getattr(dev, "cmd_timings", None)
        if not cmd_timings:
            continue

        for cmd, dts in cmd_timings.items():
            totals[cmd] = totals.get(cmd, 0.0) + sum(dts)
            counts[cmd] = counts.get(cmd, 0) + len(dts)
            maxes[cmd] = max(maxes.get(cmd, 0.0), max(dts))

    rows = []
    for cmd in totals:
        rows.append((totals[cmd], totals[cmd] / counts[cmd], maxes[cmd], counts[cmd], cmd))

    rows.sort(reverse=True)  # sort by total time
    if VERBOSE_MORE: #was VERBOSE
        print("\n================ Top command time offenders ================")
        print("Total(s)   Avg(s)   Max(s)   Runs   Command")
        print("------------------------------------------------------------")
        for total_s, avg_s, max_s, runs, cmd in rows[:top_n]:
            print(f"{total_s:8.2f} {avg_s:8.2f} {max_s:8.2f} {runs:6d}   {cmd}")
        print("============================================================\n")

def main():
    """
    Main Functino to run everything.
    """
    global GLBL_KEY_MAP
    global INPUT_FILE_NAME
    global VERBOSE
    global VERBOSE_MORE
    global VERBOSITY_LEVEL
    global RAW_CLI_OUTPUT

    update_ntc_templ_path()
    cli_arg = cli_args()
    
    # 1. Verbosity
    # Pass the integer count (0, 1, or 2) to the imported function
    set_verbose(cli_arg["verbosity"])

    VERBOSITY_LEVEL = cli_arg["verbosity"]
    VERBOSE = cli_arg["verbosity"] >= 1
    VERBOSE_MORE = cli_arg["verbosity"] >= 2
    logger = setup_logging(VERBOSE)
    
    fda.set_runtime(verbose=VERBOSE, console_=console)
    fw.set_runtime(verbose=VERBOSE, verbose_more=VERBOSE_MORE, verbosity_level=VERBOSITY_LEVEL)

    # change the Input File from the Default based on CLI Commands
    if cli_arg["input_file"]:
        INPUT_FILE_NAME = cli_arg["input_file"]
    
    if VERBOSE_MORE:
        print("NOTE: Extra verbose mode is on\n")
    elif VERBOSE:
        print("NOTE: Verbose mode is on\n")

    ###Print Starting
    spacer = "\n\n" + gen_spacer("#", 1)
    print_current_time()
    print(spacer + "\t\tStarting GetInventory Script (ThreadPoolExecutor version)" + spacer)
    # Get all global variables
    print("(1) Getting JSON Data from File")
    GLBL_KEY_MAP = get_json_data_from_file("cmd_xls_key_map.json")
    print("(2) Opening XLS File")
    work_book = open_xls(INPUT_FILE_NAME)
    print("(3) Getting Setup variables")
    setup_vars = get_setup_vars(work_book, cli_arg)
    # Create path if raw_cli_output option was selected
    if cli_arg["raw_cli_output"]:
        RAW_CLI_OUTPUT = cli_arg["raw_cli_output"]
        #verify_path(setup_vars["global"]["output_dir"]+"raw_cli")
        verify_path(os.path.join(setup_vars["global"]["output_dir"], "raw_cli"))

    fdr.set_runtime(verbose=VERBOSE, verbose_more=VERBOSE_MORE, raw_cli_output=RAW_CLI_OUTPUT)

    print_current_time()

    print("(4) Reading Network Devices")
    network_devices = read_network_devices(work_book, setup_vars["global"])

    # ---------- (4b) Autodetecting & caching device types ----------
    # Load cache once
    type_cache = load_device_type_cache()

    # Autodetect devices (only those still flagged for autodetect)
    print("\n(4b) Autodetecting device types (only where needed)...")
    max_threads = setup_vars["settings"]["max_threads"]
    autodetect_devices(network_devices, max_threads)

    # Finalize results for all active devices (Yes = should be processed)
    for dev in network_devices:
        if dev.active == "Yes":
            finalize_after_autodetect(dev)

    # Update cache with successful detections
    for dev in network_devices:
        # Cache only valid, supported detections
        if dev.parse_method and dev.parse_method not in ("Unknown", "autodetect"):
            if dev.parse_method in dev.supported_devices:
                type_cache[str(dev.host).strip().lower()] = dev.parse_method

    # Save cache
    save_device_type_cache(type_cache)

    print("(4b) Autodetect/caching complete.")
    # ---------- end (4b) ----------
    
    print_current_time()

    print('(5) Connecting to Devices and capturing commands')
    if VERBOSE:
        print("Excel Row | Host            | Message")
        print(60*"-")
    ###Connects to net_devices
    logger = setup_logging(VERBOSE, tracebacks=False)
    """If you need to add more functions or to run more commands add them to the function below"""
    if TESTING:
        testing_connection(network_devices, work_book, GLBL_KEY_MAP)
        sys.exit()
    else:
        connect_devices(network_devices, setup_vars, work_book, GLBL_KEY_MAP, INPUT_FILE_NAME, console_=console)

    # print_top_command_offenders(network_devices, top_n=20)
    
    print_current_time()

    print("(6) Saving all the Device Data")
    # Clean up Passwords before that
    remove_passwords(work_book)
    
    # Saves everything
    # """All Save features should be handled by this"""
    # save_device_data(network_devices, work_book, setup_vars, GLBL_KEY_MAP)
    
    # after all devices gathered + all results have been written to work_book:
    headers_key = map_headers(work_book)  # or reuse existing headers_key if you already have it

    post_process_lldp_sheet_with_wtp_versions(work_book, headers_key, network_devices)

    print("(7) Saving XLSX file")
    save_xls(work_book, INPUT_FILE_NAME, setup_vars["global"]["output_file"], setup_vars["global"]["output_dir"], VERBOSITY_LEVEL)
    print(spacer + "(8) DONE with the script" + spacer)
    
    print_current_time()

def print_current_time():
    # Get the current time
    current_time = get_current_time("t")
    print("Current Time:", current_time)


###Helper Functions###

def cli_args():
    """Reads the CLI options provided and returns them using the OptionParser
    Will return the Values as a dictionary"""

    def set_both_verbose(option, opt_str, value, parser):
        # 1. Set the specific destination for this flag
        setattr(parser.values, 'verbose_more', True)
        # 2. Also set the standard verbose flag
        setattr(parser.values, 'verbose', True)
    
    parser = optparse.OptionParser()
    parser.add_option('-v','--verbose',
                      dest="verbosity",
                      default=0,
                      action="count",
                      help="Enable Verbose Output. -vv Enables even more verbose outputs"
                      )
    parser.add_option('-r','--raw_cli_output',
                      dest="raw_cli_output",
                      default=False,
                      action="store_true",
                      help="Capture the raw CLI output"
                      )
    parser.add_option('-i','--input_file',
                      dest="input_file",
                      default="GetInventory - Default.xlsx",
                      action="store",
                      help="Input file name of excel sheet"
                      )
    parser.add_option('-o','--output_file',
                      dest="output_file",
                      action="store",
                      help="Output file name of excel sheet"
                      )
    parser.add_option('-d','--output_directory',
                      dest="output_dir",
                      action="store",
                      help="Output Directory of excel sheet"
                      )
    parser.add_option('-u','--username',
                      dest="username",
                      action="store",
                      help="Global Username"
                      )
    parser.add_option('-p','--password',
                      dest="password",
                      action="store",
                      help="Global Password"
                      )
    parser.add_option('-s','--secret',
                      dest="secret",
                      action="store",
                      help="Global Secret"
                      )


    options, remainder = parser.parse_args()
    
    # Utilizing the vars() method we can return the options as a dictionary
    return vars(options)

###For testing Only
def testing_connection(net_devices, wb_obj, key_map):
    """
    This is only for Testing the new functinos without multithreading.
    """
    for n, device in enumerate(net_devices):
        if device.active == "Yes":
            device.collection_time = get_current_time()
            start_time = time.time()
            connection = connect_single_device(device, n)
            # Add any command function captures here
            if connection != None:
                gather_interface(connection, device, n)
                device.interface_count = count_interfaces(device.show_output_json["show interface"])
                device.active = "Completed"
                connection.disconnect()
            else:
                device.active = "Error"
            device.elapsed_time = time.time() - start_time

"""NEED TO ASK HOW WE WILL HANDLE THIS?"""

def remove_passwords(wb_obj):
    """
    Removes the passwords from the xls workbook that will be created.
    Will not remove the password from the original input xls workbook.
    """
    rw_cell(wb_obj["Main"], 1, 2, True, "")
    rw_cell(wb_obj["Main"], 2, 2, True, "")
    rw_cell(wb_obj["Main"], 3, 2, True, "")
    for n, val in enumerate(wb_obj["Main"]["F"]):
        if n > 7: # Starts at 8
            rw_cell(wb_obj["Main"], n, 6, True, "")
            rw_cell(wb_obj["Main"], n, 7, True, "")

if __name__ == "__main__":
    main()
