#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Lai
#
# Created:     26/10/2015
#-------------------------------------------------------------------------------
import os, sys
import subprocess
import time
import datetime
from xlwings import Application,Workbook, Sheet, Range, Chart
import sheet_post
import template_search

sheet_setup = {}
sheet_arrange = {}

def get_folder_filenames(path):
    """
    Get all file names with folder under the path
    Input:  string:file_path
    Output: dict:{folder,file_names}
    """
    result = {}
    for parent, dirnames, filenames in os.walk(path):
        if not dirnames:
            result[os.path.split(parent)[-1]] = filenames
    return result

def make_folder(path,folder_names):
    """
    Make folders under the path if the folder dosen't exist
    """
    for folder in folder_names:
        if not os.path.exists(os.path.join(path,folder)):
            os.makedirs(os.path.join(path,folder))

def get_sheet_arrange():
    sheet_names = [i.name.lower() for i in Sheet.all()]
    sheet_ref = {}
    for idx in range(len(sheet_names)):
        if "2.4ghz" in sheet_names[idx]:
            if "tx" in sheet_names[idx]:
                sheet_ref["TX2G"] = idx + 1
            elif "sensitivity" in sheet_names[idx]:
                sheet_ref["RX2G"] = idx + 1
        elif "5ghz" in sheet_names[idx]:
            if "tx" in sheet_names[idx]:
                sheet_ref["TX5G"] = idx + 1
            elif "sensitivity" in sheet_names[idx]:
                sheet_ref["RX5G"] = idx + 1
    return sheet_ref

def post(data_path, data_name):
    #initial setup
    standard_anchor = "Standard"
    channel_anchor = "Ch"
    band = find_band(data_name)
    tx_or_rx = find_tx_rx(data_name)
    if not (band and tx_or_rx):
        print data_path + " is NOT a legal data file."
        return 1

    sheet = sheet_arrange[tx_or_rx + band]
##    print data_name, band, tx_or_rx, sheet
##    print sheet_setup.keys()
    if not (tx_or_rx + band) in sheet_setup:
        sheet_setup[tx_or_rx + band] = get_template_set(tx_or_rx, sheet,standard_anchor,band)

    sheet_post.post(data_path,sheet,sheet_setup[tx_or_rx + band], channel_anchor)

def save_report(wb,report_path,date,name):
    path = os.path.join(report_path,date)
    if not os.path.exists(path):
        os.makedirs(path)
    wb.save(os.path.join(path,name))

def open_workbook(path):
    subprocess.Popen(path,shell=True).pid
    Workbook.set_mock_caller(path)

def find_band(data_name):
    data_name = data_name.lower()
    if "2g" in data_name:
        return "2G"
    elif "5g" in data_name:
        return "5G"
    return None

def find_tx_rx(data_name):
    data_name = data_name.lower()
    if "tx" in data_name:
        return "TX"
    elif "rx" in data_name:
        return "RX"
    return None

def get_template_set(tx_or_rx, sheet,standard_anchor,band):
    if tx_or_rx == "TX":
        fill_pos, all_anchor_row = template_search.get_fill_pos(sheet,standard_anchor,band,1,2,3,5,6)
    else:
        fill_pos, all_anchor_row = template_search.get_fill_pos(sheet,standard_anchor,band,1,2,3,6,7)
    return (fill_pos, all_anchor_row)

if __name__ == '__main__':
    rootdir = os.path.dirname(__file__)
    log_path = os.path.join(rootdir,"Log")
    report_path = os.path.join(rootdir,"Report")
    folder_file_names = get_folder_filenames(log_path)
    t = time.time()
    date = datetime.datetime.fromtimestamp(t).strftime(r"%Y%m%d")

    if len(sys.argv) == 1:
        sys.exit(0)

    template_path = sys.argv[1]
    open_workbook(template_path)
    Workbook.set_mock_caller(template_path)
    time.sleep(10) #wait excel to execute
    wb = Workbook.caller()
    wb_true = True

##    template_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 't1.xls'))
##    open_workbook(template_path)
##    Workbook.set_mock_caller(template_path)
##    time.sleep(6)
##    wb = Workbook.caller()

    # Initial
    sheet_arrange = get_sheet_arrange()

    print "Start !"

    for folder in folder_file_names:
        if not wb_true:
            open_workbook(template_path)
            Workbook.set_mock_caller(template_path)
            time.sleep(10)
            wb = Workbook.caller()
            wb_true = True
        for data_name in folder_file_names[folder]:
            data_path = os.path.join(log_path,folder,data_name)
            print data_path
            post(data_path,data_name)
        save_report(wb,report_path,date,folder)
        Application(wb).quit()
        wb_true = False

    print "Finish !"
##
##    a=raw_input()

