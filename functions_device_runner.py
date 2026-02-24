# functions_device_runner.py

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
from rich.live import Live
from rich.layout import Layout
from rich.align import Align
from rich.text import Text
import logging
import platform
import netmiko
import openpyxl
import optparse

logger = logging.getLogger("inventory")

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
    gather_wtp_status,
    attach_wtp_to_lldp_and_update_remote_software,
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
from functions_writers import (init_next_row_cache, save_device_data, save_dev_show_json_data, write_dev_vars_to_wb, save_other_shows_to_txt, init_next_row_cache, gather_results_to_wb, add_err_msgs_to_wb, post_process_lldp_sheet_with_wtp_versions)

# I did option 2 for now:
"""
https://chatgpt.com/g/g-p-6995bb4b5ab08191bf6b127c5dea9d44-netbox/c/6998c001-043c-832a-9612-14b74be6bed7

functions_device_runner.py

It uses these globals:

VERBOSE, VERBOSE_MORE

RAW_CLI_OUTPUT 

functions_device_runner

Those were previously defined in main.py. So if you import and call connect_devices() now, you’ll hit NameError inside con_thread() unless those exist in this module.

Fix options (same two patterns):

Best (clean): pass them through as parameters (or a ctx dict)

connect_devices(..., *, verbose=False, verbose_more=False, raw_cli_output=False)

con_thread(..., verbose, verbose_more, raw_cli_output)

connect_single_device(..., verbose)

start_connection_log(..., verbose)

Quick/low-friction: module defaults + setter (same idea as autodetect)

VERBOSE = False
VERBOSE_MORE = False
RAW_CLI_OUTPUT = False

def set_runtime(verbose=None, verbose_more=None, raw_cli_output=None):
    global VERBOSE, VERBOSE_MORE, RAW_CLI_OUTPUT
    if verbose is not None: VERBOSE = verbose
    if verbose_more is not None: VERBOSE_MORE = verbose_more
    if raw_cli_output is not None: RAW_CLI_OUTPUT = raw_cli_output

Then in main.py after you compute VERBOSE, VERBOSE_MORE, RAW_CLI_OUTPUT, call:

functions_device_runner.set_runtime(...)

functions_device_autodetect.set_runtime(...)
"""

VERBOSE = False
VERBOSE_MORE = False
RAW_CLI_OUTPUT = False

def set_runtime(verbose=None, verbose_more=None, raw_cli_output=None):
    global VERBOSE, VERBOSE_MORE, RAW_CLI_OUTPUT
    if verbose is not None: VERBOSE = verbose
    if verbose_more is not None: VERBOSE_MORE = verbose_more
    if raw_cli_output is not None: RAW_CLI_OUTPUT = raw_cli_output

"""###Connection and logging of the devices###"""

def connect_devices(net_devices, setup_vars, wb_obj, key_map, input_file_name, console_):
    """
    ThreadPoolExecutor version:
    - Submits all eligible devices up front (up to max_workers active at a time)
    - Processes devices as they complete (rolling completion)
    - Saves results in the MAIN thread (Excel writes/saves are not thread-safe)
    """
    
    # Use existing console if you have one; otherwise create it
    if console_ is None:
        try:
            console_ = Console()
        except Exception:
            console_ = None  # fallback (Live requires a console; but should exist)
    
    max_workers = setup_vars["settings"]["max_threads"]
    headers_key = map_headers(wb_obj)
    next_row_cache = init_next_row_cache(wb_obj, headers_key.keys())

    # Only submit devices that are active
    devices_to_run = [(n, d) for n, d in enumerate(net_devices) if d.active == "Yes"]
    if not devices_to_run:
        return

    total_devices = len(devices_to_run)
    done = 0

    completed_batch = []
    # Tune this: how often to flush results to XLSX (every N completed devices)
    SAVE_EVERY = max(1, min(10, max_workers))  # e.g., 1..10 depending on max_workers
    
    # --- Build a layout with a 1-line footer split into left/right ---
    def footer_renderable():
        # Align right; height=None so it doesn't create extra blank space.
        msg = Text(f"Devices Finished: {done}/{total_devices}")
        return Align(msg, align="right")

    
    # Live will continually re-render just this one line.
    # redirect_stdout=False so your prints/logging don't get captured into Live.
    with Live(
        footer_renderable(),
        console=console_,
        refresh_per_second=10,
        transient=True,
        redirect_stdout=False,
        redirect_stderr=False,
        vertical_overflow="crop",
    ) as live:
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
                done += 1
                
                # update the single-line footer
                live.update(footer_renderable(), refresh=True)

                if len(completed_batch) >= SAVE_EVERY:
                    save_device_data(completed_batch, wb_obj, setup_vars, key_map, headers_key, input_file_name, next_row_cache, all_devices=net_devices)
                    completed_batch = []

        # Flush any remaining completed devices
        if completed_batch:
            save_device_data(completed_batch, wb_obj, setup_vars, key_map, headers_key, input_file_name, next_row_cache, all_devices=net_devices)

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
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather Version", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather MAC") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_arp"]:
                    try:
                        gather_arp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather ARP", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather ARP") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_mac"]:
                    try:
                        gather_mac(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather MAC", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather MAC") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_interface"]:
                    try:
                        gather_interface(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather Interface", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather Interface") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_cdp"]:
                    try:
                        gather_cdp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather CDP", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather CDP") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_lldp"]:
                    try:
                        gather_lldp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather LLDP", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather LLDP") #, exc_info=True)
                        net_dev.add_detected_error(e)
                    try:
                        gather_wtp_status(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather WAPs", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather WAPs") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_route"]:
                    try:
                        gather_route(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather Route", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather Route") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_bgp"]:
                    try:
                        gather_bgp(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather BGP", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather BGP") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_inventory"]:
                    try:
                        gather_inventory(conn, net_dev, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather Inventory", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather Inventory") #, exc_info=True)
                        net_dev.add_detected_error(e)
                if settings["gather_commands"]:
                    try:
                        gather_commands(conn, net_dev, other_shows, n)
                    except Exception as e:
                        if VERBOSE:
                            #print(60*"*"+"\n",net_dev.host, "| Issue with Gather Commands", "\n"+60*"*")
                            logger.warning(f"{net_dev.host} | Issue with Gather Commands") #, exc_info=True)
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
                #print(22*"*", "Connection Error", 21*"*")
                logger.warning(f"{net_dev.host} | Connection Error", exc_info=True)
                print_net_dev_msg(net_dev,"Unable to establish a connection")
                #print(60*"*")
        net_dev.elapsed_time = int(time.time() - start_time)
        if VERBOSE_MORE: # Was VERBOSE
            print(net_dev.host," completed in ",(time.time() - start_time))

def start_connection_log(conn, net_dev, log_path):
    """Starts logging in Append Mode"""
    log_path = Path(log_path)
    log_path = (log_path/"raw_cli"/(net_dev.host+"_raw_cli.log"))
    conn.open_session_log(str(log_path), "append")
    if VERBOSE:
        print_net_dev_msg(net_dev,"Session Logging has been enabled")

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

def get_hostname(conn, net_dev):
    """
    Gets Hostname from the Connection and saves it to the device.
    """
    t_hm = conn.find_prompt()[:-1]
    if net_dev.parse_method == "cisco_xr":
        net_dev.hostname = t_hm.split(":")[1]
    else:
        net_dev.hostname = t_hm
