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
from xlwings import Workbook, Sheet, Range, Chart
import data_post

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

def test():
    wb = Workbook.caller()
    Range('E4').value = "Hello World"

if __name__ == '__main__':
    rootdir = os.path.dirname(__file__)
    log_path = os.path.join(rootdir,"Log")
    report_path = os.path.join(rootdir,"Report")
    folder_file_names = get_folder_filenames(log_path)
    t = time.time()
    date = datetime.datetime.fromtimestamp(t).strftime(r"%Y%m%d")

##    if len(sys.argv) == 1:
##        sys.exit(0)

##    template_path = sys.argv[1]
##    subprocess.Popen([template_path],shell=True).pid
##    Workbook.set_mock_caller(template_path)
##    time.sleep(6) #wait excel to execute
    subprocess.Popen([r"D:\python task\draw_data_and_post\post\usi.xls"],shell=True).pid
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'usi.xls'))
    Workbook.set_mock_caller(path)
    time.sleep(6)
    test()
##    a=raw_input()
##
##    make_folder(os.path.join(relog_path,date),folder_file_names.keys())
##
