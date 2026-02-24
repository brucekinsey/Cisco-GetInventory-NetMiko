# network_device.py
import socket
import traceback
import sys
from datetime import datetime

import netmiko

import logging
logger = logging.getLogger("inventory")  # same name as main (or __name__)


def get_current_time(str_option="dt"):
    now = datetime.now()
    str_option = str_option.lower()
    if str_option == "dt":
        return now.strftime("%m/%d/%Y") + ", " + now.strftime("%H:%M:%S")
    if str_option == "d":
        return now.strftime("%m/%d/%Y")
    if str_option == "t":
        return now.strftime("%H:%M:%S")
    return "Invalid selection. Choose d, t, or dt."

"""########## New Network Device Class ###########"""

class NetworkDevice():
    """
    NetworkDevice class to handle and retain all output information.
    """
    supported_devices = ["cisco_xr", "cisco_ios", "cisco_nxos", "extreme_exos", "fortinet", "fortinet_fortios","fortinet_fw"]

    def __init__(
            self,
            t_host,
            t_user,
            t_pass,
            t_secret,
            t_device_type,
            t_protocol,
            port_override,
            t_active
        ):

        self.active = t_active
        self.host = t_host
        self.hostname = ""
        self.protocol = ""
        self.parse_method = ""
        self.collection_time = ""
        self.elapsed_time = 0
        self.interface_count = {}
        self.show_output = {}
        self.user_rqstd_show = {}
        self.show_output_json = {}
        self.show_for_xls = {}
        self.cpu_1_min = ""
        self.cpu_5_min = ""
        self.cpu_5_sec = ""
        self.cpu_15_min = ""
        self.model = ""
        self.serial_number = ""
        self.uptime = ""
        self.version = ""
        self.conn_error = ""
        self.running_image = ""
        self.main_col = 0
        self.json_out_file = ""
        self.cmd_out_file = ""
        self.sfp_count = ""
        self.port = 0
        self.fortinet_wtp_status = []

        # Track whether we still need to run SSHDetect
        self.needs_autodetect = (str(t_device_type).lower() == "autodetect")

        # error/comment storage (these are referenced elsewhere)
        self.comments = []
        self.error_msgs = []

        # Build base connection dict
        self.connection = {
            # IMPORTANT: SSHDetect requires device_type == 'autodetect'
            'device_type': 'autodetect' if self.needs_autodetect else str(t_device_type),
            'ip': t_host,
            'username': t_user,
            'password': t_pass,
            'secret': t_secret,
        }

        # Optional protocol/port handling
        # Only append *_telnet for telnet. For SSH, Netmiko expects base types like 'cisco_ios'
        self.protocol = (t_protocol or "ssh").lower()
        if port_override:
            self.connection["port"] = port_override
            self.port = port_override
        elif self.protocol == "telnet":
            # default telnet port
            self.connection["port"] = 23
            self.port = 23
            # If we already know the platform, Netmiko telnet types are like cisco_ios_telnet
            if not self.needs_autodetect:
                self.connection["device_type"] = f"{self.connection['device_type']}_telnet"
        else:
            # default ssh port
            self.connection["port"] = 22
            self.port = 22

        self.connection.setdefault("fast_cli", True)
        #self.connection.setdefault("global_delay_factor", 1)
        self.connection.setdefault("conn_timeout", 10)
        self.connection.setdefault("banner_timeout", 15)
        self.connection.setdefault("auth_timeout", 15)

        #self.connection["global_delay_factor"] = 2

        # If NOT autodetect, set parse_method now
        if not self.needs_autodetect:
            self.parse_method = str(t_device_type)

        # If a non-autodetect device type is unsupported, mark error
        if (not self.needs_autodetect) and (self.parse_method not in self.supported_devices):
            msg = "Unsupported device type: " + str(self.parse_method)
            self.add_error_msg(msg)
            self.active = "Error"

    def detect_device_type(self, verbose=False, console=None):
        """
        Detect device type using Netmiko SSHDetect.
        Must be run while connection['device_type'] == 'autodetect'.
        """
        guesser = None
        try:
            # Removed to show in the rich panel
            # print(self.main_col, "|", self.host, "| Detecting Device Type")

            # SSHDetect requirement
            self.connection["device_type"] = "autodetect"
            self.connection.setdefault("conn_timeout", 7)
            self.connection.setdefault("banner_timeout", 10)
            self.connection.setdefault("auth_timeout", 10)

            guesser = netmiko.SSHDetect(**self.connection)
            best_match = guesser.autodetect()

            if not best_match:
                msg = "Unable to detect the device type (SSHDetect returned no match)."
                self.add_error_msg(msg)
                self.active = "Error"
                best_match = "Unknown"

            self.parse_method = best_match
            self.needs_autodetect = False

            # Update connection for later ConnectHandler()
            if self.protocol == "telnet" and best_match not in ["Unknown", None, ""]:
                self.connection["device_type"] = f"{best_match}_telnet"
            else:
                self.connection["device_type"] = best_match
            
            return self.parse_method

        except Exception as e:
            self.add_detected_error(e)
            self.parse_method = "Unknown"
            self.active = "Error"
            return "Unknown"
        finally:
            # IMPORTANT: close SSHDetect's internal connection if present
            try:
                if guesser and getattr(guesser, "connection", None):
                    guesser.connection.disconnect()
            except Exception:
                pass
            
            
    def conn_error_detected(self, err_msg):
        """
        Add comment of connection error
        """
        self.conn_error = err_msg
        self.add_detected_error(err_msg)
        self.active = "Error"

    def is_socket_open(self, port):
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.connect((self.host, int(port)))
            sock.shutdown(2)
            return True
        except:
            return False

    def probe_port(self, port):
        if self.is_socket_open(port):
            self.port = port
            return
        elif self.is_socket_open(22):
            return "ssh"
        elif self.is_socket_open(23):
            return "telnet"
        else:
            self.add_error_msg("None of the Ports were open. Skipping this Device.")
            self.active = "Error"
            return

    def add_comment(self, t_comment):
        comment = str(len(t_comment) + 1) + " | "
        comment += str(t_comment)
        self.comments.append(comment)

    def add_error_msg(self, t_err_msg):
        comment = str(len(self.error_msgs) + 1) + " | "
        comment += str(t_err_msg)
        t_time = get_current_time()
        self.error_msgs.append([comment, t_time])

    def read_vers_info(self):
        """
        Reads the "show version" information and updates the NetworkDevices variables.
        """
        output = self.show_output_json["show version"][0]
        trantab = str.maketrans("", "", "\'\"{}[]")
        self.model = str(output["hardware"]).translate(trantab)
        self.uptime = output["uptime"]
        self.version = output["version"]
        if self.parse_method != "cisco_xr":
            self.serial_number = str(output["serial"]).translate(trantab)
            self.running_image = output["running_image"]

    def update_outdir_outfile(self, out_dir):
        """
        Updates output directory and filenames for output json / requested commands.
        """
        if self.hostname:
            self.json_out_file = self.hostname + "_" + self.host + "_JSON_cmds.json"
            self.cmd_out_file = self.hostname + "_" + self.host + "_requested_cmds.txt"
        else:
            self.json_out_file = self.host + "_JSON_cmds.json"
            self.cmd_out_file = self.host + "_requested_cmds.txt"
        self.out_dir = out_dir

    def add_detected_error(self, e):
        exc_tb = sys.exc_info()[2]
        exc_type = sys.exc_info()[0]
        exc_line = exc_tb.tb_lineno
        full_error = traceback.format_exc()
        f_name = traceback.extract_tb(exc_tb, 1)[0][2]
        t_err_msg = "{} | Exception Type: {} | At Function: {} | Line No: {} | Error Message: {}\n{}"
        t_err_msg = t_err_msg.format(self.host, exc_type, f_name, exc_line, e, full_error)
        self.add_error_msg(t_err_msg)
