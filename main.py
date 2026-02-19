"""###v3.0### -CK
Original from: https://github.com/InsightSSG/GetInventory

Detecting devices is what takes time...
"""
import socket
from threading import Thread
from concurrent.futures import ThreadPoolExecutor, as_completed
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

INPUT_FILE_NAME = "GetInventory - Default.xlsx"

"""###VERBOSE Output###"""
VERBOSE = False

RAW_CLI_OUTPUT = False

TESTING = False

"""###Global Variables###"""
GLBL_KEY_MAP = {}

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
    if VERBOSE:
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
    global RAW_CLI_OUTPUT

    update_ntc_templ_path()
    cli_arg = cli_args()

    # change the Input File from the Default based on CLI Commands
    if cli_arg["input_file"]:
        INPUT_FILE_NAME = cli_arg["input_file"]
    VERBOSE = cli_arg["verbose"]
    set_verbose(VERBOSE)

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
    """If you need to add more functions or to run more commands add them to the function below"""
    if TESTING:
        testing_connection(network_devices, work_book, GLBL_KEY_MAP)
        sys.exit()
    else:
        connect_devices(network_devices, setup_vars, work_book, GLBL_KEY_MAP, INPUT_FILE_NAME)

    print_top_command_offenders(network_devices, top_n=20)

    print("(6) Saving all the Device Data")
    # Clean up Passwords before that
    remove_passwords(work_book)
    # Saves everything
    """All Save features should be handled by this"""
    #save_device_data(network_devices, work_book, setup_vars, GLBL_KEY_MAP)
    save_xls(work_book, INPUT_FILE_NAME, setup_vars["global"]["output_file"], setup_vars["global"]["output_dir"])
    print(spacer + "(7) DONE with the script" + spacer)

def print_current_time():
    # Get the current time
    current_time = get_current_time("t")
    print("Current Time:", current_time)

def autodetect_devices(net_devices, max_workers=20):
    """
    Runs Netmiko SSHDetect autodetection concurrently for devices
    that were created with parse_method == 'autodetect'.
    """
    def _detect_one(dev):
        if getattr(dev, "needs_autodetect", False) and dev.active == "Yes":
            dev.detect_device_type()
        return dev

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = [ex.submit(_detect_one, d) for d in net_devices]
        for f in as_completed(futures):
            _ = f.result()  # raise exceptions if any

def finalize_after_autodetect(dev):
    """
    Validate detected type and mark device as Error if unsupported/unknown.
    Note: detect_device_type() should already have set dev.parse_method and dev.connection["device_type"].
    """
    if dev.parse_method in ("Unknown", "autodetect") or dev.parse_method not in dev.supported_devices:
        dev.add_error_msg(
            "Unable to detect a supported device type. Detected as: " + str(dev.parse_method)
        )
        dev.active = "Error"

def finalize_after_autodetect_orig(dev):
    """
    After SSHDetect sets dev.parse_method and dev.connection['device_type'],
    validate support and ensure correct netmiko device_type suffix + port defaults.
    """
    # If autodetect failed or returned an unsupported platform, mark Error
    if dev.parse_method in ("Unknown", "autodetect") or dev.parse_method not in dev.supported_devices:
        dev.add_error_msg(
            "Unable to detect a supported device type. Detected as: " + str(dev.parse_method)
        )
        dev.active = "Error"
        return

    # Re-apply protocol-specific netmiko driver suffix + default ports,
    # matching your original __init__ behavior.
    if dev.protocol == "telnet":
        dev.connection["device_type"] = dev.parse_method + "_telnet"
        if not dev.port:
            dev.port = "23"
    else:
        dev.connection["device_type"] = dev.parse_method
        dev.protocol = "ssh"
        if not dev.port:
            dev.port = "22"


"""###Connection and logging of the devices###"""

def connect_devices(net_devices, setup_vars, wb_obj, key_map, input_file_name):
    """
    ThreadPoolExecutor version:
    - Submits all eligible devices up front (up to max_workers active at a time)
    - Processes devices as they complete (rolling completion)
    - Saves results in the MAIN thread (Excel writes/saves are not thread-safe)
    """
    max_workers = setup_vars["settings"]["max_threads"]
    headers_key = map_headers(wb_obj)
    next_row_cache = init_next_row_cache(wb_obj, headers_key.keys())

    # Only submit devices that are active
    devices_to_run = [(n, d) for n, d in enumerate(net_devices) if d.active == "Yes"]
    if not devices_to_run:
        return

    completed_batch = []
    # Tune this: how often to flush results to XLSX (every N completed devices)
    SAVE_EVERY = max(1, min(10, max_workers))  # e.g., 1..10 depending on max_workers

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        future_to_dev = {
            ex.submit(con_thread, dev, setup_vars, n): dev
            for (n, dev) in devices_to_run
        }

        for fut in as_completed(future_to_dev):
            dev = future_to_dev[fut]

            # Make worker exceptions visible (optional but recommended)
            try:
                fut.result()
            except Exception as e:
                dev.add_error_msg(f"Worker exception: {e!r}")
                dev.active = "Error"

            # Main thread: write results for devices that just finished
            completed_batch.append(dev)

            if len(completed_batch) >= SAVE_EVERY:
                save_device_data(completed_batch, wb_obj, setup_vars, key_map, headers_key, input_file_name, next_row_cache)
                completed_batch = []

    # Flush any remaining completed devices
    if completed_batch:
        save_device_data(completed_batch, wb_obj, setup_vars, key_map, headers_key, input_file_name, next_row_cache)

def connect_devices_orig(net_devices, setup_vars, wb_obj, key_map, input_file_name):
    thread_list = []
    batch_devices = []

    max_threads = setup_vars["settings"]["max_threads"]

    for n, device in enumerate(net_devices):
        thread_list.append(Thread(target=con_thread, args=(device, setup_vars, n)))
        batch_devices.append(device)

        if len(thread_list) == max_threads or n == len(net_devices) - 1:
            for thread in thread_list:
                thread.start()
            for thread in thread_list:
                thread.join()

            # save after this batch finishes
            save_device_data(batch_devices, wb_obj, setup_vars, key_map, input_file_name)

            thread_list = []
            batch_devices = []

def con_thread(net_dev, setup_vars, n):
    """
    Function that contains all the functions to be Completed while
    multithreading, all the functions that gather data fromt the device.
    """
    other_shows = setup_vars["other_commands"]
    settings = setup_vars["settings"]
    if net_dev.active == "Yes":
        net_dev.collection_time = get_current_time()
        start_time = time.time()
        conn = connect_single_device(net_dev, n)
        # Add any command function captures here
        if conn != None:
            try:
                # Start CLI Log if True
                if RAW_CLI_OUTPUT:
                    start_connection_log(conn, net_dev, setup_vars["global"]["output_dir"])
                # Sections are executed based on Settings
                if settings["gather_version"]:
                    try:
                        gather_version(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather Version", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_arp"]:
                    try:
                        gather_arp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather ARP", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_mac"]:
                    try:
                        gather_mac(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather MAC", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_interface"]:
                    try:
                        gather_interface(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather Interface", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_cdp"]:
                    try:
                        gather_cdp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather CDP", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_lldp"]:
                    try:
                        gather_lldp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather LLDP", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_route"]:
                    try:
                        gather_route(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather Route", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_bgp"]:
                    try:
                        gather_bgp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather BGP", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_inventory"]:
                    try:
                        gather_inventory(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather Inventory", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                if settings["gather_commands"]:
                    try:
                        gather_commands(conn, net_dev, other_shows, n)
                    except Exception as e:
                        if VERBOSE:
                            print(60*"*"+"\n",net_dev.host, "| Issue with Gather Commands", "\n"+60*"*")
                        net_dev.add_detected_error(e)
                net_dev.active = "Completed"
            finally:
                try:
                    conn.disconnect()
                except Exception:
                    pass
        else:
            net_dev.active = "Error"
            if VERBOSE:
                print(22*"*", "Connection Error", 21*"*")
                print_net_dev_msg(net_dev,"Unable to establish a connection")
                print(60*"*")
        net_dev.elapsed_time = int(time.time() - start_time)
        if VERBOSE:
            print(net_dev.host," completed in ",(time.time() - start_time))

def start_connection_log(conn, net_dev, log_path):
    """Starts logging in Append Mode"""
    log_path = Path(log_path)
    log_path = (log_path/"raw_cli"/(net_dev.host+"_raw_cli.log"))
    conn.open_session_log(str(log_path), "append")
    if VERBOSE:
        print_net_dev_msg(net_dev,"Session Logging has been enabled")

def connect_single_device_orig(net_dev, count):
    """
    Attempts to Connect to Device.
    Returns False if failed to connect otherwise it will return
    ConnectionHandler variable. Will also submit "term len 0 " command.
    """
    try:
        if VERBOSE:
            print_net_dev_msg(net_dev,"Starting Connection")
        conn = netmiko.ConnectHandler(**net_dev.connection)
        conn.enable()
        if net_dev.parse_method == "extreme_exos":
            conn.send_command("dis clip")
        else:
            conn.send_command("term len 0")
        get_hostname(conn, net_dev)
        if VERBOSE:
            print_net_dev_msg(net_dev, "Hostname is: {}".format( str(net_dev.hostname)))
        return conn
    except Exception as e:
        net_dev.conn_error_detected(e)
        return None


def connect_single_device(net_dev, count):
    try:
        if VERBOSE:
            print_net_dev_msg(net_dev,"Starting Connection")

        net_dev.timing = getattr(net_dev, "timing", {})
        t0 = perf_counter()
        conn = netmiko.ConnectHandler(**net_dev.connection)
        net_dev.timing["connecthandler"] = perf_counter() - t0

        t0 = perf_counter()
        conn.enable()
        net_dev.timing["enable"] = perf_counter() - t0

        t0 = perf_counter()
        if net_dev.parse_method == "extreme_exos":
            conn.send_command("dis clip")
        else:
            conn.send_command("term len 0")
        net_dev.timing["term_len"] = perf_counter() - t0

        t0 = perf_counter()
        get_hostname(conn, net_dev)
        net_dev.timing["get_hostname"] = perf_counter() - t0

        return conn
    except Exception as e:
        net_dev.conn_error_detected(e)
        return None




####Save Functions
def save_device_data(net_devices, wb_obj, setup_vars, key_map, headers_key, input_file_name, next_row_cache):
    glbl_set = setup_vars["global"]
    settings = setup_vars["settings"]

    for net_dev in net_devices:
        start_time = perf_counter()

        net_dev.update_outdir_outfile(glbl_set["output_dir"])
        if net_dev.active == "Completed" and net_dev.elapsed_time:
            gather_results_to_wb(wb_obj, net_dev, key_map, headers_key, next_row_cache)
            if VERBOSE:
                delta_timer = perf_counter() - start_time
                #print(get_current_time("t"), " - File written: ", json_f_name)
                print(f"{delta_timer:.3f}s gather_results_to_wb for: {net_dev.host}")
            if settings["gather_commands"]:
                start_time2 = perf_counter()
                save_other_shows_to_txt(net_dev)
                delta_timer2 = perf_counter() - start_time2
                if VERBOSE:
                    #print(get_current_time("t"), " - File written: ", json_f_name)
                    print(f"{delta_timer2:.3f}s save_other_shows_to_txt for: {net_dev.host}")

        delta_timer = perf_counter() - start_time
        if VERBOSE:
            #print(get_current_time("t"), " - File written: ", json_f_name)
            print(f"{delta_timer:.3f}s gather_results_to_wb and save_other_shows_to_txt for: {net_dev.host}")

        write_dev_vars_to_wb(wb_obj, net_dev, key_map["device_info_map"])
        save_dev_show_json_data(net_dev)
        add_err_msgs_to_wb(net_dev, wb_obj)

        delta_timer = perf_counter() - start_time
        if VERBOSE:
            #print(get_current_time("t"), " - File written: ", json_f_name)
            print(f"{delta_timer:.3f}s save_device_data for: {net_dev.host}")

    save_xls(wb_obj, input_file_name, glbl_set["output_file"], glbl_set["output_dir"])

def save_dev_show_json_data(net_dev):
    """
    Same 'pretty text' format as before (spacers + ****** cmd ****** + EOF),
    but writes bytes end-to-end so we can write orjson bytes directly
    and avoid decode() overhead.
    """
    start_time = perf_counter()
    json_data = net_dev.show_output_json
    if not json_data:
        return

    json_f_name = net_dev.json_out_file
    output_dir = verify_path(net_dev.out_dir + "JSON/")
    file_name = output_dir + json_f_name

    spacer_top = gen_spacer("-", 2)
    spacer_block = gen_spacer()
    cmd_output_spacer = gen_spacer("#", 1)

    # Helper: write a text string as UTF-8 with \n normalized
    def wline(fh, s: str) -> None:
        # normalize newlines to \n, then encode
        fh.write(s.replace("\r\n", "\n").replace("\r", "\n").encode("utf-8"))

    with open(file_name, "wb") as fh:
        # Header
        wline(fh, spacer_top + center_string("Connected to " + net_dev.host) + "\n")
        wline(fh, center_string("Hostname is: " + net_dev.hostname) + "\n")
        wline(fh, spacer_top)

        # Per-command blocks
        for show_cmd, output in json_data.items():
            wline(fh, spacer_block + center_string("****** " + show_cmd + " ******") + "\n")
            wline(fh, cmd_output_spacer)
            fh.write(orjson.dumps(output, option=orjson.OPT_INDENT_2))
            wline(fh, "\n" + cmd_output_spacer)

        # EOF
        wline(fh, spacer_block + "*" * 20 + "\tEnd of File\t" + "*" * 20)
    
    delta_timer = perf_counter() - start_time
    
    if VERBOSE:
        #print(get_current_time("t"), " - File written: ", json_f_name)
        print(f"{delta_timer:.3f}s JSON write: {json_f_name}")


def write_dev_vars_to_wb(wb_obj, device, var_loc):
    """
    Writes all the data to the "Main" worksheet
    """
    wb_sheet = wb_obj["Main"]
    dev_vars = vars(device)
    row = device.main_col
    for key, col in var_loc.items():
        if key in list(dev_vars.keys()):
            if key == "interface_count" and dev_vars[key]:
                for interf, c_key in col.items():
                    values = dev_vars[key][interf]
                    rw_cell(wb_sheet, row, c_key["count"], True, str(values["count"]))
                    rw_cell(wb_sheet, row, c_key["active"], True, str(values["active"]))
            elif isinstance(col, int):
                rw_cell(wb_sheet, row, col, True, str(dev_vars[key]))

def save_other_shows_to_txt(net_dev):
    """
    Saves the results to the "other" show commands entered on the
    spreadsheet. creates a new file for every device in a subdirectory
    of the "output directory"
    """
    if net_dev.user_rqstd_show:
        file_name = verify_path(net_dev.out_dir + "show_cmd_captures/")
        spacer = gen_spacer()
        cmd_output_spacer = gen_spacer("#", 1)
        file_name += net_dev.hostname + ".txt"
        with open(file_name, "w+") as filehandle:
            write_str = spacer
            write_str += center_string("Connected to " + net_dev.host) + "\n"
            write_str += center_string("Hostname is: " + net_dev.hostname)
            write_str += "\n" + spacer
            filehandle.write(write_str)
            for show_cmd in net_dev.user_rqstd_show.keys():
                write_str = show_cmd + "\n"
                filehandle.write(write_str)
            filehandle.write(spacer)
            for show_cmd, output in net_dev.user_rqstd_show.items():
                write_str = center_string("****** " + show_cmd + " ******")
                write_str += "\n" + cmd_output_spacer + output + "\n"
                write_str += cmd_output_spacer + spacer
                filehandle.write(write_str)
            write_str = "*" * 23 + " End of File " + "*" * 24
            filehandle.write(write_str)

def init_next_row_cache(wb_obj, sheets):
    cache = {}
    for sheet_name in sheets:
        ws = wb_obj[sheet_name]
        # If column A is your “has data” column:
        cache[sheet_name] = next_available_row(ws, "A")  # one scan per sheet
    return cache

def gather_results_to_wb(wb_obj, net_dev, key_map, headers_key, next_row_cache):
    """
    Writes the Simple show commands to the work book based on the
    key_map information.
    """
    #headers_key = map_headers(wb_obj)
    show_results = net_dev.show_for_xls
    d_parse = net_dev.parse_method
    hostname = net_dev.hostname
    for setting, list_of_shows in show_results.items():
        if isinstance(list_of_shows, list):
            sheet = list(key_map[d_parse][setting].keys())[0]
            c_mapper = key_map[d_parse][setting][sheet]
            for value in list_of_shows:
                #row = next_available_row(wb_obj[sheet])
                row = next_row_cache[sheet]
                next_row_cache[sheet] += 1
                rw_cell(wb_obj[sheet], row, 1, True, hostname)
                for key, c_name in c_mapper.items():
                    if key in value.keys():
                        col = headers_key[sheet][c_name]
                        wr_val = value[key]
                        rw_cell(wb_obj[sheet], row, col, True, wr_val)
        else:
            print("Error", list_of_shows)

def add_err_msgs_to_wb(net_dev, wb_obj):
    """
    Adds all the Registered Errors to the WorkSheet.
    """
    sheet_obj = wb_obj["Errors"]
    err_msgs = net_dev.error_msgs
    row = next_available_row(sheet_obj)
    hostname = net_dev.hostname
    if not hostname:
        hostname = net_dev.host
    for i, msg in enumerate(err_msgs):
        row += i
        rw_cell(sheet_obj, row, 1, True, hostname)
        rw_cell(sheet_obj, row, 2, True, msg[1])
        rw_cell(sheet_obj, row, 3, True, msg[0])

###Helper Functions###

def get_hostname(conn, net_dev):
    """
    Gets Hostname from the Connection and saves it to the device.
    """
    t_hm = conn.find_prompt()[:-1]
    if net_dev.parse_method == "cisco_xr":
        net_dev.hostname = t_hm.split(":")[1]
    else:
        net_dev.hostname = t_hm

def cli_args():
    """Reads the CLI options provided and returns them using the OptionParser
    Will return the Values as a dictionary"""
    parser = optparse.OptionParser()
    parser.add_option('-v','--verbose',
                      dest="verbose",
                      default=False,
                      action="store_true",
                      help="Enable Verbose Output"
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
