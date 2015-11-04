#-------------------------------------------------------------------------------
# Name:        main
# Purpose:
#
# Author:      Lai
#
# Created:     04/11/2015
#-------------------------------------------------------------------------------
import os, sys
import subprocess
import time
import datetime
from xlwings import Application,Workbook, Sheet, Range, Chart

def open_workbook(path):
    """
    Open file
    """
    subprocess.Popen(path,shell=True).pid


if __name__ == '__main__':
##    # the program path
##    rootdir = os.path.dirname(__file__)
##    # log folder path
##    log_path = os.path.join(rootdir,"Log")
##    # report folder path
##    result_path = os.path.join(rootdir,"Result")
##    # Get all files with dict (folder:file_name)
##    folder_file_names = get_folder_filenames(log_path)
##    t = time.time()
##    # today in year/month/day
##    date = datetime.datetime.fromtimestamp(t).strftime(r"%Y%m%d")

    t = r"D:\python task\draw_data_and_post\data\Log"
    path = t + r"\2G_Tx_MIMO\MFGC_2G_Tx_MIMO2_HT20_combine_output.xlsx"

    open_workbook(path)
    Workbook.set_mock_caller(path) # set workbook hook
    time.sleep(6) #wait excel to execute
    wb = Workbook.caller()
    Application(wb).quit()

    open_workbook(path)
    Workbook.set_mock_caller(path) # set workbook hook
    time.sleep(6) #wait excel to execute
    wb = Workbook.caller()

