#!python3
# coding: utf-8

"""
---------------------------------

Webnovel to EPUB

Version: 03

Script created to scan all connected USB-Devices to find any Amazon Kindle-device.
If found, checks if Windows identifies them as dirty; if so, scans them with chkdsk
to unset this.

---------------------------------
"""

import winreg
import sys
import contextlib
import itertools
import os
import re
import codecs
import subprocess

def run_command(command):
    """
    Executes the supplied command and returns the output
    """
    output = subprocess.getoutput(command)
    return output


def get_subkeys(path, hkey=winreg.HKEY_LOCAL_MACHINE, flags=0):
    with contextlib.suppress(WindowsError), winreg.OpenKey(hkey, path, 0, winreg.KEY_READ | flags) as k:
        for i in itertools.count():
            yield winreg.EnumKey(k, i)


print(f"Getting all known Kindle devices...")

kindle_devices = list()

root_path = r'SYSTEM\ControlSet001\Enum\SWD\WPDBUSENUM'
for subkey in get_subkeys(root_path):
    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, os.path.join(root_path, subkey), 0, winreg.KEY_READ)
    leaf_dict = dict()
    for i in range(0, winreg.QueryInfoKey(key)[1]):
        leaf_tuple = winreg.EnumValue(key, i)
        leaf_tuple_simple = leaf_tuple[:2]
        if type(leaf_tuple_simple[1]) == str:
            leaf_dict[leaf_tuple_simple[0]] = leaf_tuple_simple[1].strip()
        else:
            leaf_dict[leaf_tuple_simple[0]] = leaf_tuple_simple[1]

    if leaf_dict.get('Mfg') == "Kindle":

        regex_device_id = re.compile(r"(&Rev_[0-9]+#)\w+")
        m = regex_device_id.search(subkey)

        if m:
            device_guid = m.group().split('#')[1]
            leaf_dict['DeviceGUID'] = device_guid
            kindle_devices.append(leaf_dict)

devices_connected = False

print(f"Known kindle devices: {len(kindle_devices)}")

leaf_dict = dict()
with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\MountedDevices", 0, winreg.KEY_READ) as target_key:
    for i in range(0, winreg.QueryInfoKey(target_key)[1]):

        leaf_tuple = winreg.EnumValue(target_key, i)

        if 'DosDevices' in leaf_tuple[0]:

            decoded_value = codecs.decode(leaf_tuple[1], "ascii", "ignore")
            sentence = re.sub(r'\W+', '', decoded_value)

            for device in kindle_devices:

                is_genuine_kindle = False
                cleanup_needed = False

                if device['DeviceGUID'] in sentence:

                    drive_letter = leaf_tuple[0].split('\\')[-1]

                    if os.path.isdir(os.path.join(drive_letter, 'audible')) \
                            and os.path.isdir(os.path.join(drive_letter, 'documents')) \
                            and os.path.isdir(os.path.join(drive_letter, 'fonts')) \
                            and os.path.isdir(os.path.join(drive_letter, 'voice')):
                        is_genuine_kindle = True

                if is_genuine_kindle:

                    devices_connected = True

                    print(f"Connected Kindle detected at {drive_letter}")

                    check_dirty_status = run_command(f"fsutil dirty query {drive_letter}")
                    if check_dirty_status.strip() == f"Volume - {drive_letter} is Dirty":
                        cleanup_needed = True

                    if cleanup_needed:
                        print("Kindle detected as dirty...")
                        run_checkdisk = run_command(f"chkdsk /f {drive_letter}")

                        if 'File and folder verification is complete.' in run_checkdisk:
                            print("Device cleaned")
                        else:
                            print(run_checkdisk)
                            sys.exit("Something went wrong")
                    else:
                        print(f"Device is not dirty")

if not devices_connected:
    print(f"No Kindle device connected")

print(f"Script finished!")
