
import subprocess
from win32 import win32print
import winreg
import win32api
import shlex
import time
import os
import shutil
import ctypes
from datetime import datetime
import tempfile
import sys
import json
import logging
from logging.handlers import RotatingFileHandler
from windows_toasts import Toast,WindowsToaster


logger = logging.getLogger("Error Log")
logger.setLevel(logging.ERROR)
handler = RotatingFileHandler(os.path.join(tempfile.gettempdir(), "OWMError.log"), maxBytes=1000000, backupCount=0, encoding=None, delay=0)
logger.addHandler(handler)
IS_OFFLINE = "0"
IS_OUT_OF_INK = "1"
IS_STUCK = "2"
IS_OUT_OF_PAPER = "3"
IS_JOB_USER_INTERVENTION = "4"
IS_PRINTER_ERROR = "5"
IS_UNKNOWN = "6"
IS_READY = "7"

def checkToasterAvailable(pcp_name):
    isToastAvailable = False
    try:
        wintoaster = WindowsToaster(pcp_name)
        newToast = Toast()
        isToastAvailable = True

        return isToastAvailable,newToast,wintoaster
    except Exception as os_error:
        print("This is the OS Error",os_error)
        isToastAvailable = False
        return False,None,None
pcp_name = "PCP 2.0.6"
isToastAvailable, newToast, wintoaster = checkToasterAvailable(pcp_name)


def get_printer_status(status):
    if status == win32print.PRINTER_STATUS_PAUSED:
        print('Printer is paused')
    elif status == win32print.PRINTER_STATUS_ERROR:
        return IS_PRINTER_ERROR
    elif status == win32print.PRINTER_STATUS_PENDING_DELETION:
        print('Printer is pending deletion')
    elif status == win32print.PRINTER_STATUS_PAPER_JAM:
        return IS_STUCK
    elif status == win32print.PRINTER_STATUS_PAPER_OUT:
        return IS_OUT_OF_PAPER
    elif status == win32print.PRINTER_STATUS_MANUAL_FEED:
        print('Printer is waiting for manual feed')
    elif status == win32print.PRINTER_STATUS_OFFLINE:
        return IS_OFFLINE
    elif status == win32print.PRINTER_STATUS_BUSY:
        print('Printer is busy')
    elif status == win32print.PRINTER_STATUS_TONER_LOW:
        return IS_OUT_OF_INK
    elif status == win32print.JOB_STATUS_USER_INTERVENTION:
        return IS_JOB_USER_INTERVENTION
    else:
        return IS_UNKNOWN

temp = tempfile.gettempdir()
# Initialize printer status JSON file
printer_status_file = os.path.join(temp, 'PCP_Printer_pages_log.json')
print(printer_status_file)
printer_status_data = {}

if os.path.exists(printer_status_file):
    with open(printer_status_file, "r") as f:
        printer_status_data = json.load(f)

def check_if_folder_exists(path):
    return os.path.exists(path)

def create_folder_in_given_path(dir_name, parent_dir):
    if not check_if_folder_exists(os.path.join(parent_dir, dir_name)):
        os.mkdir(os.path.join(parent_dir, dir_name))
        FILE_ATTRIBUTE_HIDDEN = 0x02
        ctypes.windll.kernel32.SetFileAttributesW(os.path.join(parent_dir, dir_name), FILE_ATTRIBUTE_HIDDEN)
        return {"result": "folder is created"}
    else:
        print("Path exists")
        return {"result": "Folder already exists"}

# Function to print to a printer
async def printToPrinter(dirname, filename, printerName, paramName, websocket, json_data):
    si = subprocess.STARTUPINFO()
    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    directory_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    print(directory_path)
    all_folders = os.listdir(directory_path)

    if "bin" in all_folders:
        subprocess.call(["bin\gsbatchprintc.exe", "-P", printerName, "-F", filename, "-q"], shell=False,startupinfo=si)

        # Update printer status JSON file
        if printerName in printer_status_data:
            printer_status_data[printerName]["pages_printed"] += 1
        else:
            printer_status_data[printerName] = {"pages_printed": 1}

        # Write the updated data back to the JSON file
        with open(printer_status_file, "w") as f:
            json.dump(printer_status_data, f)

        await websocket.send(json.dumps({
            "statusCode": 200,
            "status": "Success",
            "message": "Printed the Document Successfully",
            "payload": {"_id": json_data['_id'], "name": json_data['name'],
                         "printerName": json_data["printerName"], 'Number of pages printed': printer_status_data[printerName]["pages_printed"]},
            "isPrinted": True
        }))
        os.remove(filename)
    else:
        if isToastAvailable:
            newToast.text_fields = ['gsbatchprintc.exe is not available']
            wintoaster.show_toast(newToast)
        await websocket.send(json.dumps({
            "statusCode": 520,
            "status": "Failure",
            "message": "gsbatchprintc.exe is not available. Please reinstall the application",
            "payload": {"_id": json_data['_id'], "name": json_data['name'],
                         "printerName": json_data["printerName"]},
            "isPrinted": False
        }))
        os.remove(filename)


