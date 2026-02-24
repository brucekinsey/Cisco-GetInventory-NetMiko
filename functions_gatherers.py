# gatherers.py

import re
import os

from time import perf_counter
from collections import defaultdict
import traceback

from functions_helpers import left, mod_dir_based_on_os, print_net_dev_msg

# NOTE: These gather functions rely on a bunch of helpers that currently live in main.py.
# We import only what we need, or we re-home the helpers here.
#
# Easiest/cleanest path:
# - Move the helper functions used by gather_* into this file too (recommended).
# - Keep main.py importing gatherers.* and the rest of the app unchanged.

"""######## Gather Functions that run based on Settings Tab ########"""

## XR Ready
def gather_version(connection, net_dev, count):
    """
    Captures the show version command and saves the textfsm outcome to
    the NetworkDevice. No additional parsing needed at the moment.

    Captures version-ish info and updates the NetworkDevice vars.
    """
    if net_dev.parse_method in ["fortinet"]:
        command1= "get system status"
        command2= "get system performance status"
        txt_tmpl1 = r"ntc-templates\templates\fortinet_get_system_status.textfsm"
        output1 = log_cmd_textfsm(connection, net_dev, command1, count, txt_tmpl1)

        # read_vers_info() expects show_output_json["show version"][0]
        if isinstance(output1, list) and output1:
            net_dev.show_output_json["show version"] = output1
        else:
            # keep the expected key present to avoid KeyError later
            net_dev.show_output_json["show version"] = [{
                "hardware": "",
                "uptime": "",
                "version": "",
                "serial": "",
                "running_image": ""
            }]
    else:
        command = "show version"
        log_cmd_textfsm(connection, net_dev, command, count)

    net_dev.read_vers_info()
    show_proc_cpu(connection, net_dev, count)

## XR Ready, Need to test more
def gather_arp_orig(connection, net_dev, count):
    """
    Captures arp information and utilizing the vrf data it parses the
    output to prepare it for extraction to WB.
    """
    vrf_list = get_vrf_names(net_dev, connection, count)
    arp_list = []
    for vrf in vrf_list:
        vrf_string, output = "", ""
        if vrf != "global":
            vrf_string = " vrf " + vrf
        if net_dev.parse_method == "cisco_xr":
            command = "show arp" + vrf_string
            txt_tmpl = "ntc-templates/test_tmpl/cisco_xr_show_arp.textfsm"
            txt_tmpl = mod_dir_based_on_os(txt_tmpl)
        else:
            if net_dev.parse_method in ["fortinet"]:
                command = "get system arp"
                txt_tmpl = (r"ntc-templates\templates\fortinet_get_system_arp.textfsm")
            else:
                command = "show ip arp" + vrf_string
                txt_tmpl = None
        output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
        if isinstance(output, list):
            for arp in output:
                arp["vrf"] = vrf
                if net_dev.parse_method in ["cisco_nxos","fortinet"]:
                    arp["type"] = "ARPA"
                arp_list.append(arp)
        else:
            arp = {}
            arp['vrf'] = vrf
            arp['address'] = "No ARP Data Found"
            arp_list.append(arp)
    net_dev.show_for_xls["gather_arp"] = arp_list

def gather_arp(connection, net_dev, count):
    """
    Captures arp information and utilizing the vrf data it parses the
    output to prepare it for extraction to WB.
    """
    arp_list = []

    # -------------------------
    # Fortinet path: no VRF loop; just run get system arp once
    # -------------------------
    if net_dev.parse_method in ["fortinet"]:
        command = "get system arp"
        txt_tmpl = r"ntc-templates\templates\fortinet_get_system_arp.textfsm"
        output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)

        if isinstance(output, list):
            for arp in output:
                arp["vrf"] = "global"
                arp["type"] = "ARPA"   # keep consistent with NXOS handling
                arp_list.append(arp)
        else:
            arp_list.append({
                "vrf": "global",
                "address": "No ARP Data Found"
            })

        net_dev.show_for_xls["gather_arp"] = arp_list
        return

    # -------------------------
    # Non-Fortinet path (existing behavior)
    # -------------------------
    vrf_list = get_vrf_names(net_dev, connection, count)

    for vrf in vrf_list:
        vrf_string = ""
        if vrf != "global":
            vrf_string = " vrf " + vrf

        if net_dev.parse_method == "cisco_xr":
            command = "show arp" + vrf_string
            txt_tmpl = "ntc-templates/test_tmpl/cisco_xr_show_arp.textfsm"
            txt_tmpl = mod_dir_based_on_os(txt_tmpl)
        else:
            command = "show ip arp" + vrf_string
            txt_tmpl = None

        output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)

        if isinstance(output, list):
            for arp in output:
                arp["vrf"] = vrf
                if net_dev.parse_method in ["cisco_nxos"]:
                    arp["type"] = "ARPA"
                arp_list.append(arp)
        else:
            arp_list.append({
                "vrf": vrf,
                "address": "No ARP Data Found"
            })

    net_dev.show_for_xls["gather_arp"] = arp_list

## XR Ready, Need to test more
def gather_mac(connection, net_dev, count):
    """
    All the parsing of the
    """
    command = "show mac address-table"
    if net_dev.parse_method in ["fortinet", "fortinet_fortios", "fortinet_fw"]:
        return
    
    if net_dev.parse_method == "cisco_xr":
        output = []
        cmd = "show l2vpn forwarding bridge-domain {} mac-address location {}"
        locations = get_xr_locations(net_dev, connection, count)
        bg_grp_dmns = get_xr_bg_grp_dmns(net_dev, connection, count)
        txtfsm = "ntc-templates/test_tmpl/cisco_xr_show_l2vpn_bridge-domain_mac.textfsm"
        txtfsm = mod_dir_based_on_os(txtfsm)
        for dmn in bg_grp_dmns:
            for lctn in locations:
                command = cmd.format(dmn, lctn)
                output += log_cmd_textfsm(connection, net_dev, command, count, txtfsm)
        if not output:
            output += [{"MAC": "No MAC Data"}]
    else:
        output = log_cmd_textfsm(connection, net_dev, command, count)
    if isinstance(output, list):
        net_dev.show_for_xls["gather_mac"] = output.copy()
    elif isinstance(output, str):
        output = {}
        output["type"] = "No MAC Address Table results found for device"
        net_dev.show_for_xls["gather_mac"] = [output]

## XR Ready, Need to test more
def gather_interface(connection, net_dev, count):
    """
    Gather Interface information, the function was modified from the
    original script.
    """
    # -------------------------
    # Fortinet path
    # -------------------------
    if net_dev.parse_method in ["fortinet"]:
        cmd_int = "get system interface"
        cmd_phy = "get system interface physical"

        txt_tmp1 = (r"ntc-templates\test_tmpl\fortinet_get_system_interface.textfsm")
        txt_tmp2 = (r"ntc-templates\test_tmpl\fortinet_get_system_interface_physical.textfsm")
        #txt_tmp1 = _dbg_template(r"ntc-templates\test_tmpl\fortinet_get_system_interface.textfsm")
        #txt_tmp2 = _dbg_template(r"ntc-templates\test_tmpl\fortinet_get_system_interface_physical.textfsm")

        out_int = log_cmd_textfsm(connection, net_dev, cmd_int, count, txt_tmp1)
        out_phy = log_cmd_textfsm(connection, net_dev, cmd_phy, count, txt_tmp2)
        # try:
            # out_int = log_cmd_textfsm(connection, net_dev, cmd_int, count, txt_tmp1)
            # print("out_int type:", type(out_int))
        # except Exception:
            # print("EXCEPTION in get system interface")
            # traceback.print_exc()
            # raise

        # try:
            # out_phy = log_cmd_textfsm(connection, net_dev, cmd_phy, count, txt_tmp2)
            # print("out_phy type:", type(out_phy))
        # except Exception:
            # print("EXCEPTION in get system interface physical")
            # traceback.print_exc()
            # raise

        # If parsing failed, log_cmd_textfsm will return a string (raw CLI output / error text)
        if not isinstance(out_int, list):
            net_dev.show_for_xls["gather_interface"] = [{
                "status": "No Interface Data (Fortinet): template/command did not parse to list"
            }]
            net_dev.interface_count = 0
            return

        # OPTIONAL: build a lookup from physical output by interface name
        phy_by_name = {}
        if isinstance(out_phy, list):
            for p in out_phy:
                # depending on your template, this might be "name" or "interface"
                name = p.get("name") or p.get("interface") or p.get("port") or ""
                if name:
                    phy_by_name[name] = p

        # Normalize fields so your existing XLS headers/mapping can populate
        for row in out_int:
            # many Fortinet templates output 'name' not 'interface'
            if "interface" not in row and "name" in row:
                row["interface"] = row["name"]

            # short_if is expected by your Interfaces sheet logic/headers
            row["short_if"] = row.get("interface", row.get("name", ""))

            # If your template already outputs ip_address + netmask, great.
            # If it outputs a combined 'ip' field, you can keep it or split later.
            row.setdefault("ip_address", row.get("ip_address", ""))
            row.setdefault("netmask", row.get("netmask", ""))
            row.setdefault("status", row.get("status", ""))

            # Try to enrich speed/duplex from physical output if available
            phy = phy_by_name.get(row.get("name", "")) or phy_by_name.get(row.get("interface", ""))
            if phy:
                # Your physical output example has speed like: "1000Mbps (Duplex: full)"
                sp = phy.get("speed", "")
                row["speed"] = sp
                # If your physical template doesn't split duplex, you can still store speed string.
                row.setdefault("duplex", phy.get("duplex", ""))

        net_dev.show_for_xls["gather_interface"] = out_int
        #net_dev.interface_count = len(out_int)
        total = len(out_int)
        active = sum(1 for r in out_int if str(r.get("status", "")).lower() == "up")

        net_dev.interface_count = {
            "Ethernet": {"count": total, "active": active},
            "FastEthernet": {"count": 0, "active": 0},
            "GigabitEthernet": {"count": 0, "active": 0},
            "TenGigEthernet": {"count": 0, "active": 0},
            "TwentyFiveGigEthernet": {"count": 0, "active": 0},
            "FortyGigEthernet": {"count": 0, "active": 0},
            "HundredGigEthernet": {"count": 0, "active": 0},
            "Serial": {"count": 0, "active": 0},
            "Subinterfaces": {"count": 0, "active": 0},
            "Tunnel": {"count": 0, "active": 0},
            "Port-channel": {"count": 0, "active": 0},
            "Loopback": {"count": 0, "active": 0},
            "VLAN": {"count": 0, "active": 0},
        }
        return

    # ---- Non-Fortinet (original behavior) ----

    ####This is from old
    command = "show interface"
    command2 = "show interface status"
    log_cmd_textfsm(connection, net_dev, command, count)
    log_cmd_textfsm(connection, net_dev, command2, count)
    trunks = {}
    vrf_info = get_vrf_interfaces_dict(net_dev, connection, count)

    output = net_dev.show_output_json[command].copy()
    output2 = net_dev.show_output_json[command2]
    dev_type = net_dev.parse_method
    switchport_data_found = False
    if isinstance(output2, list):
        switchport_data_found = True
        # If device is a switch, get trunk info into dictionary.
        #log_cmd_textfsm(connection, net_dev, "show int trunk", count)
        trunks = get_trunk_dict(net_dev, connection)

    # Write Interface data to spreadsheet 'Interfaces' tab
    if isinstance(output, list):
        for i in output:
            short_if_name = get_short_if_name(i['interface'], dev_type)
            i["short_if"] = short_if_name
            if i['ip_address'] != "":
                i['l2_l3'] = "Layer 3"
                i['trunk_access'] = "Routed"
                if isinstance(vrf_info, list):
                    vrf_name = ""
                    for vrf in vrf_info:
                        for intf in vrf["interfaces"]:
                            t_intf = get_short_if_name(intf, dev_type).lower()
                            if short_if_name.lower() == t_intf:
                                vrf_name = vrf["name"]
                                i['vrf'] = vrf_name
                    if vrf_name == "" and dev_type != "cisco_xr":
                        i['vrf'] = "default"
                    elif vrf_name == "":
                        i['vrf'] = "global"
                else:
                    i['vrf'] = "global"
            if switchport_data_found is True:
                for x in output2:
                    if short_if_name.lower() == x['port'].lower():
                        if x['vlan'].isnumeric():
                            i["vlan"] = x['vlan']
                            i["vlan"] = x['vlan']
                            i["trunk_access"] = "Access"
                            i["l2_l3"] = "Layer 2"
                        elif x['vlan'] == "trunk":
                            i["l2_l3"] = "Layer 2"
                            i["trunk_access"] = "Trunk"
                            # Parse Trunk Details
                            native = get_trunk_details(short_if_name,
                                                       trunks,
                                                       "vlans_native",
                                                       net_dev)
                            allowed = get_trunk_details(short_if_name,
                                                        trunks,
                                                        "vlans_allowed",
                                                        net_dev)
                            not_pruned = get_trunk_details(short_if_name,
                                                           trunks,
                                                           "vlans_not_pruned",
                                                           net_dev)
                            i["native"] = native
                            i["allowed"] = allowed
                            i["not_pruned"] = not_pruned

                        elif x['vlan'] == "routed":
                            i["trunk_access"] = "Routed"
        net_dev.show_for_xls["gather_interface"] = output
        net_dev.interface_count = count_interfaces(net_dev.show_output_json["show interface"])
    elif isinstance(output, str):
        output = {}
        output["status"] = "No Interface Data, have Developer check the script"
        net_dev.show_for_xls["gather_interface"] = [output]

## XR Ready, need to test
def gather_cdp(connection, net_dev, count):
    
    # Fortinet only has LLDP, CDP = Cisco Discovery Protocol
    if net_dev.parse_method in ["fortinet"]:
        return
    
    command = "show cdp neighbor detail"
    command2 = "show cdp neighbor"
    output = log_cmd_textfsm(connection, net_dev, command, count)
    output2 = log_cmd_textfsm(connection, net_dev, command2, count)
    if len(output) != len(output2):
        if VERBOSE:
            print_net_dev_msg(net_dev, "The output of the length of show cdp neigh and thelength of show cdp neighbor details is not the same, please manually gather raw command of both commands")
        net_dev.add_error_msg("The output of the length of show cdp neigh and the length of show cdp neighbor details is not the same, please manually gather raw command of both commands")
    if isinstance(output, list):
        net_dev.show_for_xls["gather_cdp"] = output.copy()
    elif isinstance(output, str):
        output = {}
        output["local_port"] = "No CDP Data"
        net_dev.show_for_xls["gather_cdp"] = [output]

## XR Ready, need to test
def gather_lldp(connection, net_dev, count):
    command = "show lldp neighbor detail"
    txt_tmpl = None

    if net_dev.parse_method == "cisco_xr":
        txt_tmpl = (r"ntc-templates\test_tmpl\cisco_xr_show_lldp_neighbors_detail.textfsm")
        #txt_tmpl = mod_dir_based_on_os(txt_tmpl)
    if net_dev.parse_method == "extreme_exos":
        txt_tmpl = (r"ntc-templates\test_tmpl\extreme_exos_show_lldp_neighbors_detail.textfsm")
    if net_dev.parse_method in ["fortinet"]:
        command = "diagnose lldprx port neighbor details"
        txt_tmpl = (r"ntc-templates\test_tmpl\fortinet_diagnose_lldprx_port_neighbor_details_port-name.textfsm")
    try:
        output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
    except Exception as e:
        # Treat LLDP failures as non-fatal (no traceback spam)
        net_dev.add_error_msg(f"LLDP gather or parse failed: {e!r}")
        net_dev.show_for_xls["gather_lldp"] = [{"chassis_id": "No LLDP Data"}]
        return
    
    if net_dev.parse_method == "fortinet" and isinstance(output, list):
        for row in output:
            v = row.get("system_desc_data")
            if isinstance(v, list):
                row["system_desc_data"] = v[0] if v else ""
            elif v is None:
                row["system_desc_data"] = ""

    if isinstance(output, list):
        net_dev.show_for_xls["gather_lldp"] = output.copy()
    elif isinstance(output, str):
        net_dev.show_for_xls["gather_lldp"] = [{"chassis_id": "No LLDP Data"}]

## XR Ready, need to test
def gather_route(connection, net_dev, count):
    if net_dev.parse_method in ["fortinet"]:
        command = "get router info routing-table all"
        #txt_tmpl = _dbg_template(r"ntc-templates\templates\fortinet_get_router_info_routing-table_all.textfsm")
        txt_tmpl = (r"ntc-templates\templates\fortinet_get_router_info_routing-table_all.textfsm")

        try:
            output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
            #print("route output type:", type(output))
        except Exception:
            print("EXCEPTION in get router info routing-table all")
            traceback.print_exc()
            raise

        route_list = []
        if isinstance(output, list):
            for route in output:
                route.setdefault("vrf", "global")
                if "cidr" not in route:
                    if route.get("network") and route.get("mask"):
                        route["cidr"] = f"{route['network']}/{route['mask']}"
                    elif route.get("destination"):
                        route["cidr"] = route["destination"]
                route_list.append(route)
        else:
            route_list.append({
                "vrf": "global",
                "protocol": "No route table parsed",
                "destination": "No Route Data Found"
            })

        net_dev.show_for_xls["gather_route"] = route_list
        return
    
    # -------------------------
    # Non-Fortinet path (existing behavior)
    # -------------------------
    vrf_list = get_vrf_names(net_dev, connection, count)
    route_list = []
    route_table_present = False
    for vrf in vrf_list:
        vrf_string = ""
        if vrf != "global":
            vrf_string = " vrf " + vrf
        command = "show ip route" + vrf_string
        output = log_cmd_textfsm(connection, net_dev, command, count)
        if isinstance(output, list):
            for route in output:
                route["cidr"] = route['network'] + "/" + route['mask']
                if "vrf" not in list(route.keys()):
                    route["vrf"] = vrf
                route_list.append(route)
                route_table_present = True
        if not route_table_present:
            route = {}
            default_gateway = re.search(r"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", output)
            route["vrf"] = vrf
            route["protocol"] = "Layer 2 only"
            if default_gateway:
                route["nexthop_ip"] = default_gateway[0]
            else:
                output2 = connection.send_command("show run | incl default-gateway")
                default_gateway = re.search(r"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", output2)
                if default_gateway:
                    route["nexthop_ip"] = default_gateway[0]
            route_list.append(route)
    net_dev.show_for_xls["gather_route"] = route_list

## XR Ready, Need to test more
def gather_bgp(connection, net_dev, count):
    if net_dev.parse_method in ["fortinet"]:
        command = "get router info bgp summary"
        #txt_tmpl = _dbg_template(r"ntc-templates\templates\fortinet_get_router_info_bgp_summary.textfsm")
        txt_tmpl = (r"ntc-templates\templates\fortinet_get_router_info_bgp_summary.textfsm")

        try:
            output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
            #print("bgp output type:", type(output))
        except Exception:
            if VERBOSE:
                print("",net_dev.host," - EXCEPTION in get router info bgp summary")
                traceback.print_exc()
            raise

        if isinstance(output, str):
            # Show raw output snippet so we can tune TextFSM (or confirm BGP is disabled)
            msg = f"{net_dev.host} - BGP summary not parsed; first 400 chars: {output[:400]!r}"
            if VERBOSE:
                print(msg)
            net_dev.add_error_msg(msg)
            net_dev.show_for_xls["gather_bgp"] = [{"status": "No BGP Data"}]
            return
        if isinstance(output, list) and output:
            for row in output:
                row.setdefault("vrf", "global")
            net_dev.show_for_xls["gather_bgp"] = output
        else:
            net_dev.show_for_xls["gather_bgp"] = [{"status": "No BGP Data"}]
        return

    # -------------------------
    # Non-Fortinet path (existing behavior)
    # -------------------------
    command = "show ip bgp"
    if net_dev.parse_method == "cisco_xr":
        vrf_names = get_vrf_names(net_dev, connection, count)
        output = []
        txt_fsm = "ntc-templates/test_tmpl/cisco_xr_show_bgp_vrf.textfsm"
        for vrf in vrf_names:
            command = "show bgp vrf {} ipv4 unicast".format(vrf)
            if isinstance(output, list):
                output += log_cmd_textfsm(connection, net_dev, command, count, txt_fsm)
            else:
                output["vrf"] = vrf
                output["status"] = "No BGP Data"
                net_dev.show_for_xls["gather_bgp"] = [output]
    else:
        output = log_cmd_textfsm(connection, net_dev, command, count)
    if isinstance(output, list):
        net_dev.show_for_xls["gather_bgp"] = output
    elif isinstance(output, str):
        output = {}
        output["status"] = "No BGP Data"
        net_dev.show_for_xls["gather_bgp"] = [output]

## XR Ready
def gather_inventory(connection, net_dev, count):
    if net_dev.parse_method == "fortinet":
        command = "get system interface transceiver"
        txt_tmpl = (r"ntc-templates\test_tmpl\fortinet_get_system_interface_transceiver.textfsm")
        #txt_tmpl = _dbg_template(r"ntc-templates\test_tmpl\fortinet_get_system_interface_transceiver.textfsm")

        output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
        #print(output)

        if isinstance(output, list):
            normalized = []
            for r in output:
                vendor = (r.get("vendor_name", "") or "").strip()
                descr  = (r.get("description", "") or "").strip()

                # Only include vendor in descr when it's meaningful
                if vendor and vendor.upper() != "OEM":
                    descr = f"{vendor} | {descr}"

                normalized.append({
                    "pid": r.get("part_id", "") or "",
                    "name": r.get("device", "") or "",
                    "descr": descr,
                    "sn": r.get("serial_number", "") or "",
                })

            # store + count
            net_dev.show_for_xls["gather_inventory"] = normalized
            update_sfp_cout(net_dev, normalized)
        else:
            net_dev.show_for_xls["gather_inventory"] = [{
                "status": "Issue parsing Fortinet transceiver inventory"
            }]

        return

    command = "show inventory"
    txt_tmpl = None
    if net_dev.parse_method == "cisco_xr":
        txt_tmpl = "ntc-templates/templates/cisco_ios_show_inventory.textfsm"
        txt_tmpl = mod_dir_based_on_os(txt_tmpl)
    output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
    
    if isinstance(output, list):
        net_dev.show_for_xls["gather_inventory"] = output.copy()
        update_sfp_cout(net_dev, output)
    elif isinstance(output, str):
        output = {}
        output["status"] = "Issue, have Developer check the script"
        net_dev.show_for_xls["gather_inventory"] = [output]

def update_sfp_cout(net_dev, inventory_list):
    sfp_count = 0
    for item in inventory_list:
        pid = (item.get("pid") or "").lower()
        descr = (item.get("descr") or "").lower()
        if "sfp" in pid or "sfp" in descr:
            sfp_count += 1
    net_dev.sfp_count = sfp_count

# Fortigate Only
def gather_wtp_status(connection, net_dev, count):
    """
    Fortinet-only: gather WTP status records from:
      get wireless-controller wtp-status
    Stores list[dict] in net_dev.wtp_status (placeholder attribute).
    """
    if net_dev.parse_method != "fortinet":
        return

    command = "get wireless-controller wtp-status"
    #txt_tmpl = _dbg_template(r"ntc-templates\test_tmpl\fortinet_get_wireless_controller_wtp_status.textfsm")
    txt_tmpl = (r"ntc-templates\test_tmpl\fortinet_get_wireless_controller_wtp_status.textfsm")
    # If you want, wrap with mod_dir_based_on_os(txt_tmpl)

    output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
    
    # try:
        # output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
        # print("output type:", type(output))
    # except Exception:
        # print("EXCEPTION in get system interface")
        # traceback.print_exc()
        # raise
    
    # print(output[:1])

    if isinstance(output, list):
        # Normalize keys to lowercase (consistent with your earlier TextFSM results)
        net_dev.fortinet_wtp_status = [{k.lower(): v for k, v in rec.items()} for rec in output]
    else:
        net_dev.fortinet_wtp_status = []

def attach_wtp_to_lldp_and_update_remote_software(net_dev, all_devices_by_host=None):
    """
    1) Create LLDP-like rows from net_dev.fortinet_wtp_status and append them to
       net_dev.s:contentReference[oaicite:16]{index=16}l:contentReference[oaicite:17]{index=17}n the LLDP sheet.
    2) If all_devices_by_host (or by hostname) is provided, try to find the neighbor
       device and append WTP software-version to its software/version string.
    """
    if net_dev.parse_method != "fortinet":
        return

    wtps = getattr(net_dev, "fortinet_wtp_status", None) or []
    if not wtps:
        return

    lldp_rows = net_dev.show_for_xls.get("gather_lldp")
    if not isinstance(lldp_rows, list):
        lldp_rows = []
        net_dev.show_for_xls["gather_lldp"] = lldp_rows

    for w in wtps:
        # normalize keys (your TextFSM returns lowercase already, but safe)
        w = {k.lower(): v for k, v in w.items()}

        wtp_name = (w.get("wtp_name") or "").strip()
        sw = (w.get("software_version") or "").strip()
        wtp_ip = (w.get("local_ipv4_addr") or w.get("mgmt_ip") or "").strip()
        board_mac = (w.get("board_mac") or "").strip()
        lldp_sys = (w.get("lldp_sys_name") or "").strip()
        lldp_port = (w.get("lldp_port_id") or "").strip()

        # ---- 1) Add synthetic LLDP row for the AP itself ----
        # Use fields consistent with your LLDP TextFSM output keys for Fortinet.
        # (Adjust key names here to match fortinet_diagnose_lldprx_port_neighbor_details_port-name.textfsm)
        lldp_rows.append({
            "local_port": w.get("local_port", "") or "",         # optional if you captured it
            "system_name": wtp_name,                              # remote host = WTP name
            "management_ip": wtp_ip,                              # MGMT IP = WTP local-ipv4-addr
            "port_id": board_mac or lldp_port,                    # remote port preference: board-mac, else lldp port id
            "port_descr": "",                                     # optional
            "system_desc_data": f"WTP software-version: {sw}" if sw else "WTP",
            "chassis_id": board_mac or "",                        # also put board mac here if your sheet uses chassis_id
        })

        # ---- 2) Update the *remote neighbor* device software/version ----
        # Matching rules you asked for:
        #  - MGMT IP match (lldp neighbor mgmt ip) OR
        #  - Remote Port = board-mac OR (short IF + switch hostname) OR
        #  - Remote Host = WTP name
        #
        # Practically: easiest reliable match is LLDP sys name -> device.hostname
        # (you have that in lldp_sys_name)
        if not all_devices_by_host:
            continue

        target = None

        # Try by LLDP sys name (strip domain)
        if lldp_sys:
            short_host = lldp_sys.split(".")[0].strip().lower()
            target = all_devices_by_host.get(short_host) or all_devices_by_host.get(lldp_sys.lower())

        # Fallback: try by MGMT IP if you keep an ip->device map
        # (optional to implement later)

        if target and sw:
            # Your “software” is effectively device.version in your class
            cur = (getattr(target, "version", "") or "").strip()
            tag = f"WTP:{wtp_name}={sw}" if wtp_name else f"WTP={sw}"
            if tag not in cur:
                setattr(target, "version", (cur + " | " + tag).strip(" |"))

def apply_wtp_software_to_lldp(all_devices):
    """
    For each FortiGate device that has fortinet_wtp_status:
      - find the referenced switch (from wtp['lldp_sys_name'])
      - find the matching LLDP neighbor row on that switch by:
          * MGMT IP match (wtp local-ipv4-addr or mgmt_ip)
          * OR remote port match (board-mac OR lldp_port_id)
          * OR remote host match (wtp_name)
      - append the WTP software-version to that LLDP row's system description field

    Runs in MAIN THREAD (called from save_device_data) to avoid thread-safety issues.
    """

    # Build a hostname lookup for all devices
    dev_by_name = {}
    for d in all_devices:
        if not getattr(d, "hostname", None):
            continue
        hn = d.hostname.strip().lower()
        dev_by_name[hn] = d
        dev_by_name[hn.split(".")[0]] = d  # allow fqdn/shortname matching

    def _norm_host(s: str) -> str:
        s = (s or "").strip().lower()
        return s.split(".")[0] if s else ""

    def _norm_mac(s: str) -> str:
        return (s or "").strip().lower()

    def _append_once(base: str, tag: str) -> str:
        base = (base or "").strip()
        if not tag:
            return base
        return base if tag in base else (base + (" | " if base else "") + tag)

    for fg in all_devices:
        if getattr(fg, "parse_method", "") != "fortinet":
            continue

        wtps = getattr(fg, "fortinet_wtp_status", None) or []
        if not wtps:
            continue

        for w in wtps:
            w = {k.lower(): v for k, v in (w or {}).items()}

            wtp_name = (w.get("wtp_name") or "").strip()
            wtp_sw = (w.get("software_version") or "").strip()
            wtp_ip = (w.get("local_ipv4_addr") or w.get("mgmt_ip") or "").strip()
            board_mac = _norm_mac(w.get("board_mac"))

            # Which switch did the FortiGate report the AP is connected to?
            sw_name = _norm_host(w.get("lldp_sys_name"))
            if not sw_name:
                continue

            sw_dev = dev_by_name.get(sw_name)
            if not sw_dev:
                continue

            lldp_rows = sw_dev.show_for_xls.get("gather_lldp")
            if not isinstance(lldp_rows, list) or not lldp_rows:
                continue

            # Build matching inputs
            wtp_port_fallback = (w.get("lldp_port_id") or "").strip()

            # Try to find matching LLDP neighbor row on the switch
            matched = False
            for row in lldp_rows:
                if not isinstance(row, dict):
                    continue

                # These key names must match your LLDP template output.
                # Adjust here if your LLDP dict uses different keys.
                row_host = (row.get("system_name") or row.get("sys_name") or "").strip()
                row_mgmt_ip = (row.get("management_ip") or row.get("ip") or row.get("mgmt_ip") or "").strip()
                row_port = (row.get("port_id") or row.get("remote_port") or "").strip()
                row_chassis = _norm_mac(row.get("chassis_id") or "")

                host_match = (wtp_name and row_host and row_host.strip().lower() == wtp_name.lower())
                ip_match = (wtp_ip and row_mgmt_ip and row_mgmt_ip == wtp_ip)
                port_match = False

                if board_mac:
                    port_match = (_norm_mac(row_port) == board_mac) or (row_chassis == board_mac)
                elif wtp_port_fallback:
                    port_match = (row_port == wtp_port_fallback)

                if host_match or ip_match or port_match:
                    tag = f"WTP_SW={wtp_sw}" if wtp_sw else "WTP"
                    # This is the Fortinet LLDP field you already normalize in gather_lldp()
                    # If your mapping uses a different field, adjust below.
                    row["system_desc_data"] = _append_once(row.get("system_desc_data", ""), tag)
                    matched = True
                    break

            # Optionally also append into the switch device's 'version' field (Main sheet "software")
            if matched and wtp_sw:
                sw_tag = f"WTP:{wtp_name}={wtp_sw}" if wtp_name else f"WTP={wtp_sw}"
                sw_dev.version = _append_once(getattr(sw_dev, "version", ""), sw_tag)

## XR Ready
def gather_commands(connection, net_dev, other_shows, count=0):
    """
    Logs the other show commands requested by the user, in "Commands" sheet
    """
    for command in other_shows:
        #if VERBOSE:
            #print_net_dev_msg(net_dev, "Capturing '{}' as raw text".format(command))
        output = connection.send_command(command)
        net_dev.user_rqstd_show[command] = output

# -------------------------
# Helpers used by gather_*
# -------------------------
# If you leave these in main.py, gatherers.py would need to import them from main,
# which risks circular imports if main also imports gatherers.
#
# So: move them here too (recommended). Use placeholders unless you want me to refactor.

#moved
def _dbg_template(p):
    try:
        rp = mod_dir_based_on_os(p)  # same function log_cmd_textfsm uses
        print("Template raw:", p)
        print("Template resolved:", rp)
        print("Exists?", os.path.exists(rp), "Abs:", os.path.abspath(rp))
        return rp
    except Exception:
        print("Template path resolution failed!")
        traceback.print_exc()
        return p

def log_cmd_textfsm(connection, net_dev, command, count, txtfsm_tmpl=None):
    """
    Logs command output with TextFSM enabled and records elapsed time.
    Stores timings on net_dev.cmd_timings[command] as a list of elapsed seconds.
    """
    #if VERBOSE:
        #print_net_dev_msg(net_dev, f"Capturing '{command}' with TextFSM Enabled")

    if txtfsm_tmpl:
        txtfsm_tmpl = mod_dir_based_on_os(txtfsm_tmpl)

    # Lazy init so you don't have to edit the class yet
    if not hasattr(net_dev, "cmd_timings") or net_dev.cmd_timings is None:
        net_dev.cmd_timings = defaultdict(list)

    t0 = perf_counter()
    output = connection.send_command(
        command,
        use_textfsm=True,
        textfsm_template=txtfsm_tmpl,
        # optionally: read_timeout=XX (later, once you identify slow cmds)
    )
    dt = perf_counter() - t0
    net_dev.cmd_timings[command].append(dt)

    # existing behavior
    if isinstance(output, list):
        net_dev.show_output_json[command] = output.copy()
    else:
        net_dev.show_output_json[command] = output

    return output

#moved
def show_proc_cpu(connection, net_dev, count):
    """
    Sends Commands "show processes cpu" and logs to variables of
    the NetworkDevice.
    """
    dev_type = net_dev.parse_method
    command = "show processes cpu"
    output = log_cmd_textfsm(connection, net_dev, command, count)
    try:
        if dev_type in ["cisco_ios", "cisco_nxos"]:
            net_dev.cpu_5_sec = output[0]['cpu_5_sec']
            net_dev.cpu_1_min = output[0]['cpu_1_min']
            net_dev.cpu_5_min = output[0]['cpu_5_min']
        elif dev_type == "cisco_xr":
            net_dev.cpu_1_min = output[0]['cpu_1_min']
            net_dev.cpu_5_min = output[0]['cpu_5_min']
            net_dev.cpu_15_min = output[0]['cpu_15_min']
    except Exception as e:
        net_dev.add_detected_error(e)

#moved
def get_vrf_names(net_dev, connection, count):
    """
    Gathers the VRF names from the connection. If the names were already
    gathered then it returns the list from the NetworkDevice variable of
    vrf_names.
    """
    #fortinet: show router vrf
    if net_dev.parse_method in ["fortinet"]:
        # Cache it so we don't re-evaluate later
        if not hasattr(net_dev, "vrf_names"):
            net_dev.vrf_names = ["global"]
        return net_dev.vrf_names
        
    try:
        vrf_names = net_dev.vrf_names
    except AttributeError:
        vrf_names = ["global"]
        txt_tmpl = None
        command = "show vrf"
        if net_dev.parse_method == "cisco_xr":
            vrf_names = ["default"]
            command += " all"
            txt_tmpl = "ntc-templates/test_tmpl/cisco_xr_show_vrf_all.textfsm"
            txt_tmpl = mod_dir_based_on_os(txt_tmpl)
        output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)

        #if VERBOSE:
            #print_net_dev_msg(net_dev, "Parsing vrfs")
        if isinstance(output, list):
            for vrf in output:
                if vrf['name'] not in vrf_names:
                    vrf_names.append(vrf['name'])
        elif isinstance(output, str):
            net_dev.add_error_msg("Issue with the Gather VRF Names, seems to be an issue with 'show vrf', check textfsm template., it is not parsing the data into a list, get a string.")
        net_dev.vrf_names = vrf_names
    return vrf_names

#moved
def get_xr_locations(net_dev, connection, count):
    """
    Gathers the locations
     names from the connection. If the names were already
    gathered then it returns the list from the NetworkDevice variable of
    vrf_names.
    """
    try:
        return net_dev.xr_locations
    except AttributeError:
        if VERBOSE:
            print_net_dev_msg(net_dev,"Parsing XR Device Locations")
        command = "show l2vpn forwarding bridge-domain : mac-address location ?"
        output = connection.send_command(command)
        connection.send_command("")
        net_dev.xr_locations = parse_locations_frm_prmpt(output)
        return net_dev.xr_locations

#moved
def get_xr_bg_grp_dmns(net_dev, connection, count):
    """
    Gathers the Bridge group and domain and returns them with a colon
    """
    try:
        bg_grp_dmns = net_dev.bg_grp_dmns
    except AttributeError:
        if VERBOSE:
            print_net_dev_msg(net_dev,"Parsing BGP Group and Domain")
        net_dev.bg_grp_dmns = []
        command = "show l2vpn bridge-domain"
        txt_tmpl = "ntc-templates/test_tmpl/cisco_xr_show_l2vpn_forwarding_bridge_info.textfsm"
        output = log_cmd_textfsm(connection, net_dev, command, count, txt_tmpl)
        if isinstance(output, list):
            for i in output:
                net_dev.bg_grp_dmns.append("{}:{}".format(i["bridge_group"], i["bridge_domain"]))
    return net_dev.bg_grp_dmns

#moved
def get_vrf_interfaces_dict(device, conn, count):
    """
    This function will parse VRF information to get interface and VRF
    info.
    THIS IS MESSY, IT NEEDS TO BE CLEANED
    """
    command = "show vrf"
    if device.parse_method == "cisco_xr":
        vrf_names = get_vrf_names(device, conn, count)
        return_list = []
        for vrf in vrf_names:
            if vrf != "default":
                vrf_dict = {}
                command = "show vrf " + vrf + " detail"
                output = conn.send_command(command)
                start_log = False
                for line in output.split("\n"):
                    if line == "Interfaces:":
                        start_log = True
                    elif "Address" in line:
                        start_log = False
                        break
                    elif start_log:
                        vrf_dict["name"] = vrf
                        vrf_dict["interface"] = line[2:]
                        return_list.append(vrf_dict)
        return return_list
    if device.parse_method == "cisco_nxos":
        command += " interface"

    output = log_cmd_textfsm(conn, device, command, count)
    if "Invalid input detected at" in output:
        rtr_str = "Invalid Input"
        return rtr_str
    if isinstance(output, str):
        device.add_error_msg("Issue with 'get_vrf_interfaces_dict', seems to be an issue with 'show vrf', check textfsm template., it is not parsing the data into a list, get a string. \nString Output:\n"+output+"\n")
    return output

#moved
def get_trunk_dict(device, connection):
    """
    Parses the 'show int trunk' response, only works on cisco_ios and cisco_nxos
    """
    trunk_all_info = connection.send_command("show int trunk").split('\n')
    vlans_native_list, vlans_allowed_list, vlans_forwarding_list = [], [], []
    vlans_not_pruned_list, vlans_err_disabled_list = [], []
    x = 0

    if device.parse_method == "cisco_ios":
        for line in trunk_all_info:
            if line not in ["", " ", "\n"]:
                if x == 0:
                    # Find the first reference to the word 'port' which will house native vlan data
                    first_word = line.split(" ")[0]
                    if first_word.lower() == "port":
                        x = x + 1
                elif x == 1:
                    first_word = line.split(" ")[0]
                    if first_word.lower() == "port":
                        x = x + 1
                    # Add lines to native_vlan_list:
                    else:
                        vlans_native_list.append(line)
                elif x == 2:
                    first_word = line.split(" ")[0]
                    # Increment counter when word 'Port' is found again
                    if first_word.lower() == "port":
                        x = x + 1
                    # Add lines to vlan_allowed_list:
                    else:
                        vlans_allowed_list.append(line)
                elif x == 3:
                    first_word = line.split(" ")[0]
                    # Increment counter when word 'Port' is found again
                    if first_word.lower() == "port":
                        x = x + 1
                    # Add lines to vlan_active_list:
                    else:
                        vlans_forwarding_list.append(line)
                elif x == 4:
                    # Add lines to vlan_active_list:
                    vlans_not_pruned_list.append(line)

    if device.parse_method == "cisco_nxos":
        for line in trunk_all_info:
            if line != "" and left(line, 1) != " " and line != "\n":
                if left(line, 3) != "---" and left(line, 7) != "Feature":
                    if x == 0:
                        # Find the first reference to the word 'port'
                        # which will house native vlan data.
                        first_word = line.split(" ")[0]
                        if first_word.lower() == "port":
                            x = x + 1
                    elif x == 1:
                        # Find second instance of the word 'port'
                        # (VLANS Allowed) and increment counter
                        first_word = line.split(" ")[0]
                        if first_word.lower() == "port":
                            x = x + 1
                        # Add lines to native_vlan_list:
                        else:
                            vlans_native_list.append(line)
                    elif x == 2:
                        # Increment counter when word 'Port'
                        # is found again (ERR Disabled)
                        first_word = line.split(" ")[0]
                        if first_word.lower() == "port":
                            x = x + 1
                        # Add lines to vlan_err_disabled_list:
                        else:
                            vlans_allowed_list.append(line)
                    elif x == 3:
                        # Increment counter when word 'Port'
                        # is found again (ERR Disabled)
                        first_word = line.split(" ")[0]
                        if first_word.lower() == "port":
                            x = x + 1
                        # Add lines to vlan_err_disabled_list:
                        else:
                            vlans_err_disabled_list.append(line)
                    elif x == 4:
                        first_word = line.split(" ")[0]
                        # Increment counter when word 'Port'
                        # is found again (STP Forwarding)
                        if first_word.lower() == "port":
                            x = x + 1
                        # Add lines to vlan_active_list:
                        else:
                            vlans_forwarding_list.append(line)
                    elif x == 5:
                        # Add lines to vlan_active_list:
                        vlans_not_pruned_list.append(line)

    return {'vlans_native': vlans_native_list,
            'vlans_allowed': vlans_allowed_list,
            'vlans_err_disabled': vlans_err_disabled_list,
            'vlans_forwarding': vlans_forwarding_list,
            'vlans_not_pruned': vlans_not_pruned_list,
            }

#moved
def get_trunk_details(if_name, trunk_dict, key_value, net_dev):
    """
    Provides the trunk details for partivular key_value:
    vlans_native, vlans_allowed, vlans_not_pruned
    """
    try:
        for x in trunk_dict[key_value]:
            interface, value = "", ""
            if key_value != "vlans_native":
                interface = x.split()[0]
                value = x.split()[1]
            else:
                if net_dev.parse_method == "cisco_ios":
                    interface = x.split()[0]
                    value = x.split()[4]
                elif net_dev.parse_method == "cisco_nxos":
                    interface = x.split()[0]
                    value = x.split()[1]
            if if_name.lower() == interface.lower():
                return value
    except Exception as e:
        net_dev.add_detected_error(e)
        if VERBOSE:
            print_net_dev_msg(net_dev, "Error getting trunk info from device: "+str(e))
        return "Error, review Errors Log"

#moved
def get_short_if_name(interface, device_type):
    """
    Returns short if name. for cisco_ios it returns first 2 char and the
    interface number. Everything else it returns the first 3 plus number.
    """
    number = re.compile(r"(\d.*)$")
    name = re.compile("([a-zA-Z]+)")
    number = number.search(interface).group(1)
    name = name.search(interface).group(1)
    short_name = ""
    if device_type == "cisco_ios":
        short_name = left(name, 2)
    elif device_type in ["cisco_nxos", "cisco_xr"]:
        port = left(name, 3).lower()
        if port == "eth":
            short_name = left(name, 3)
        elif port == "vla":
            short_name = left(name, 4)
        elif port == "mgm":
            short_name = left(name, 4)
        else:
            short_name = left(name, 2)
    if int(left(number, 1)) >= 0 or number is None:
        short_name = short_name + str(number)
    return short_name

#moved
def parse_locations_frm_prmpt(raw_str):
    """
    Parses the ? prompt to get the locations
    """
    locations_list = []
    for line in raw_str.split("\n"):
        print('"' + str(line) + '"')
        if line and ":" not in line and "WORD" not in line and "ncomplete" not in line:
            for i in line.split(" "):
                if i:
                    locations_list.append(i)
                    break
    return locations_list

#moved from main
def count_interfaces(if_dictionary):
    """
    Count Number of interfaces per device, input has to be
    "show interface" with TEXTFSM=True
    """
    return_dict = {
        "Ethernet": {"count": 0, "active": 0},
        "FastEthernet": {"count": 0, "active": 0},
        "GigabitEthernet": {"count": 0, "active": 0},
        "TenGigEthernet": {"count": 0, "active": 0},
        "TwentyFiveGigEthernet": {"count": 0, "active": 0},
        "FortyGigEthernet": {"count": 0, "active": 0},
        "HundredGigEthernet": {"count": 0, "active": 0},
        "Serial": {"count": 0, "active": 0},
        "Subinterfaces": {"count": 0, "active": 0},
        "Tunnel": {"count": 0, "active": 0},
        "Port-channel": {"count": 0, "active": 0},
        "Loopback": {"count": 0, "active": 0},
        "VLAN": {"count": 0, "active": 0}
    }
    for i in if_dictionary:
        split_if = i['interface'].split(".")
        if len(split_if) == 1:
            if left(i['interface'], 3).lower() == "eth":
                return_dict["Ethernet"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["Ethernet"]["active"] += 1
            elif left(i['interface'], 3).lower() == "fas":
                return_dict["FastEthernet"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["FastEthernet"]["active"] += 1
            elif left(i['interface'], 3).lower() == "gig":
                return_dict["GigabitEthernet"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["GigabitEthernet"]["active"] += 1
            elif left(i['interface'], 3).lower() == "ten":
                return_dict["TenGigEthernet"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["TenGigEthernet"]["active"] += 1
            elif left(i['interface'], 3).lower() == "twe":
                return_dict["TwentyFiveGigEthernet"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["TwentyFiveGigEthernet"]["active"] += 1
            elif left(i['interface'], 3).lower() == "for":
                return_dict["FortyGigEthernet"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["FortyGigEthernet"]["active"] += 1
            elif left(i['interface'], 3).lower() == "hun":
                return_dict["HundredGigEthernet"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["HundredGigEthernet"]["active"] += 1
            elif left(i['interface'], 6).lower() == "serial":
                return_dict["Serial"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["Serial"]["active"] += 1
            elif left(i['interface'], 3).lower() == "tun":
                return_dict["Tunnel"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["Tunnel"]["active"] += 1
            elif left(i['interface'], 5).lower() == "port-":
                return_dict["Port-channel"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["Port-channel"]["active"] += 1
            elif left(i['interface'], 3).lower() == "loo":
                return_dict["Loopback"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["Loopback"]["active"] += 1
            elif left(i['interface'], 3).lower() == "vla":
                return_dict["VLAN"]["count"] += 1
                if i['link_status'] == "up":
                    return_dict["VLAN"]["active"] += 1
        elif len(split_if) == 2:
            return_dict["Subinterfaces"]["count"] += 1
            if i['link_status'] == "up":
                return_dict["Subinterfaces"]["active"] += 1
    return return_dict

# -------------------------
# Small gotcha: VERBOSE
# -------------------------
# Your gather functions reference VERBOSE (a global in main.py).
# To avoid importing main.py, make VERBOSE a module variable here and have main set it.

VERBOSE = False
VERBOSE_MORE = False

def set_verbose(verbosity_count):
    """
    Updates the global verbose flags based on the verbosity count.
    0 = Off
    1 = Verbose
    2+ = Verbose More
    """
    global VERBOSE
    global VERBOSE_MORE

    if verbosity_count >= 2:
        VERBOSE = True
        VERBOSE_MORE = True
    elif verbosity_count == 1:
        VERBOSE = True
        VERBOSE_MORE = False
    else:
        VERBOSE = False
        VERBOSE_MORE = False