# functions_device_autodetect.py
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

# I did option 2 for now:
"""
https://chatgpt.com/g/g-p-6995bb4b5ab08191bf6b127c5dea9d44-netbox/c/6998c001-043c-832a-9612-14b74be6bed7

What you need to update (globals actually used)
functions_device_autodetect.py

It uses:

console (passed into Live(... console=console ...))

VERBOSE (passed into dev.detect_device_type(VERBOSE, console) and for logging) 

functions_device_autodetect

Fix options:

Best (clean): make them parameters

def autodetect_devices(net_devices, max_workers=20, *, verbose=False, console=None):
    if console is None:
        console = Console()

and update the ex.submit(dev.detect_device_type, verbose, console) line accordingly.

Quick/low-friction: add module defaults + a setter

VERBOSE = False
console = Console()

def set_runtime(verbose=None, console_=None):
    global VERBOSE, console
    if verbose is not None: VERBOSE = verbose
    if console_ is not None: console = console_

Then in main.py, after parsing CLI verbosity / creating console, call set_runtime(...).
"""


VERBOSE = False
console = Console()

def set_runtime(verbose=None, console_=None):
    global VERBOSE, console
    if verbose is not None: VERBOSE = verbose
    if console_ is not None: console = console_

def autodetect_devices(net_devices, max_workers=20):
    """
    Live, rolling autodetect:
    - Only max_workers devices "in progress" at any time
    - Updates a single on-screen status view
    """
    # Devices that actually need autodetect
    targets = [d for d in net_devices if getattr(d, "needs_autodetect", False) and d.active == "Yes"]
    total = len(targets)
    if total == 0:
        return

    # --- try to use rich if available ---
    try:
        from rich.live import Live
        from rich.panel import Panel
        from rich.text import Text

        def render(done_count, in_prog):
            t = Text()
            t.append(f"Detecting Device Type ({done_count}/{total}):\n", style="bold")
            if in_prog:
                # format like: "8 | 192.168.2.6; 10 | 192.168.2.11; ..."
                pieces = [f"{d.main_col} | {d.host}" for d in in_prog]
                # wrap nicely by letting Rich handle it
                t.append("; ".join(pieces) + ";")
            else:
                t.append("(none)")
            return Panel(t, title="(4b) Autodetecting device types (only where needed)...", border_style="cyan")

        # Rolling scheduler
        pending_iter = iter(targets)
        in_progress = set()
        futures = {}

        with ThreadPoolExecutor(max_workers=max_workers) as ex, Live(
            render(0, []),
            console=console,
            refresh_per_second=8,
            redirect_stdout=True,
            redirect_stderr=True,
        ) as live:
            done_count = 0

            def submit_next():
                try:
                    dev = next(pending_iter)
                except StopIteration:
                    return False
                # IMPORTANT: stop detect_device_type() from printing spam lines (see note below)
                #fut = ex.submit(dev.detect_device_type)
                fut = ex.submit(dev.detect_device_type, VERBOSE, console)
                futures[fut] = dev
                in_progress.add(dev)
                return True

            # Prime the pool
            for _ in range(min(max_workers, total)):
                submit_next()

            # Loop until all done
            while futures:
                done, _ = wait(futures.keys(), return_when=FIRST_COMPLETED)
                for fut in done:
                    dev = futures.pop(fut)
                    # surface exceptions (optional)
                    try:
                        best = fut.result()
                        if VERBOSE:
                            console.log(f"{dev.main_col} | {dev.host} | Device was detected as: {best}")
                    except Exception as e:
                        dev.add_error_msg(f"Autodetect worker exception: {e!r}")
                        dev.active = "Error"

                    in_progress.discard(dev)
                    done_count += 1

                    # Keep the pool full
                    submit_next()

                # Update display
                # Sort so it’s stable-looking
                live.update(render(done_count, sorted(in_progress, key=lambda d: d.main_col)))

    except ImportError:
        # --- fallback: simple print-based status (no extra modules) ---
        # NOTE: this won't be as pretty, but it won't spam lines.
        pending_iter = iter(targets)
        in_progress = set()
        futures = {}

        def redraw(done_count):
            # clear screen-ish (works in most terminals)
            print("\033[2J\033[H", end="")  # clear + home (ANSI)
            print("(4b) Autodetecting device types (only where needed)...")
            print(f"Detecting Device Type ({done_count}/{total}):")
            if in_progress:
                pieces = [f"{d.main_col} | {d.host}" for d in sorted(in_progress, key=lambda d: d.main_col)]
                print("; ".join(pieces) + ";")
            else:
                print("(none)")

        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            done_count = 0

            def submit_next():
                try:
                    dev = next(pending_iter)
                except StopIteration:
                    return False
                fut = ex.submit(dev.detect_device_type, VERBOSE, console)
                futures[fut] = dev
                in_progress.add(dev)
                return True

            for _ in range(min(max_workers, total)):
                submit_next()
            redraw(done_count)

            while futures:
                done, _ = wait(futures.keys(), return_when=FIRST_COMPLETED)
                for fut in done:
                    dev = futures.pop(fut)
                    try:
                        fut.result()
                    except Exception:
                        pass
                    in_progress.discard(dev)
                    done_count += 1
                    submit_next()
                redraw(done_count)

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
