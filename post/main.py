#-------------------------------------------------------------------------------
# Name:        main
# Purpose:     Open template excel file and read folder to find all data .csv.
#              Then start to call function to post value and when it finish save the report
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

# save template sheet setup to avoid setup re-construct. This is for speed up
sheet_setup = {}
# save explicit (name:sheet_pos) reflection to which will be use in post
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

def post(data_path, data_name):
    """
    Get data path and pass to post function to post value
    """
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
    """
    Save workbook at report_path. If folder not exist, it will make
    """
    path = os.path.join(report_path,date)
    if not os.path.exists(path):
        os.makedirs(path)
    wb.save(os.path.join(path,name))

def open_workbook(path):
    """
    Open file
    """
    subprocess.Popen(path,shell=True).pid

def find_band(data_name):
    """
    Find it is "2.4G" or "5G"
    """
    data_name = data_name.lower()
    if "2g" in data_name:
        return "2G"
    elif "5g" in data_name:
        return "5G"
    return None

def find_tx_rx(data_name):
    """
    Find it is "TX" or "RX"
    """
    data_name = data_name.lower()
    if "tx" in data_name:
        return "TX"
    elif "rx" in data_name:
        return "RX"
    return None

def get_template_set(tx_or_rx, sheet,standard_anchor,band):
    """
    Recall template setup if it exist. If template setup isn't exist, call "get_fill_pos" to make.
    """
    if tx_or_rx == "TX":
        # standard_x = 1,module_x = 2,rate_x = 3, case_x = 5, start_x = 6
        fill_pos, all_anchor_row = template_search.get_fill_pos(sheet,standard_anchor,band,1,2,3,5,6)
    else:
        # standard_x = 1,module_x = 2,rate_x = 3, case_x = 6, start_x = 7
        fill_pos, all_anchor_row = template_search.get_fill_pos(sheet,standard_anchor,band,1,2,3,6,7)
    return (fill_pos, all_anchor_row)

if __name__ == '__main__':
    # the program path
    rootdir = os.path.dirname(__file__)
    # log folder path
    log_path = os.path.join(rootdir,"Log")
    # report folder path
    report_path = os.path.join(rootdir,"Report")
    # Get all files with dict (folder:file_name)
    folder_file_names = get_folder_filenames(log_path)
    t = time.time()
    # today in year/month/day
    date = datetime.datetime.fromtimestamp(t).strftime(r"%Y%m%d")

    # Need get template path
    if len(sys.argv) == 1:
        sys.exit(0)

    template_path = sys.argv[1]
    open_workbook(template_path)
##    Workbook.set_mock_caller(template_path) # set workbook hook
    time.sleep(10) #wait excel to execute
##    wb = Workbook.caller()
    wb = Workbook.active()
    wb_true = True
    Application(wb).screen_updating = False
##    template_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 't1.xls'))
##    open_workbook(template_path)
##    Workbook.set_mock_caller(template_path)
##    time.sleep(6)
##    wb = Workbook.caller()

    # Get sheets pos
    sheet_arrange = template_search.get_sheet_arrange()

    print "Start !"

    for folder in folder_file_names:
        if not wb_true:
            open_workbook(template_path)
##            Workbook.set_mock_caller(template_path)
            time.sleep(10)
##            wb = Workbook.caller()
            wb = Workbook.active()
            Application(wb).screen_updating = False
            wb_true = True
        for data_name in folder_file_names[folder]:
            data_path = os.path.join(log_path,folder,data_name)
            print data_path
            post(data_path,data_name)
        save_report(wb,report_path,date,folder)
        Application(wb).screen_updating = True
        Application(wb).quit()
        wb_true = False

    print "Finish !"
##
##    a=raw_input()

