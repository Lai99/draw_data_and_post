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
import sheet_post
import template_search

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

def post(data_path):
    #initial get all sheet setup in template
    standard_anchor = "Standard"
    channel_anchor = "Ch"
    sheet_setup = {}
    # for test
    sheet = 5
    band = "5G"
    #
    fill_pos, all_anchor_row = template_search.get_fill_pos(sheet,standard_anchor,band,1,2,3,6,7)
##    print fill_pos
##    print all_anchor_row

    sheet_setup["TX_2G"] = [fill_pos, all_anchor_row]
    #judge tx/rx, 2.4/5G -> pass data path to right post module
##    sheet = 3
##    band = "5G"
    sheet_post.post(data_path,sheet,sheet_setup["TX_2G"], standard_anchor,channel_anchor,band)

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
##    wb = Workbook.caller()


##    subprocess.Popen([r"D:\game\abstract\draw_data_and_post\post\t 1.xls"],shell=True).pid
    subprocess.Popen([os.path.join(os.path.dirname(__file__), 't1.xls')],shell=True).pid
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), 't1.xls'))
    Workbook.set_mock_caller(path)
    time.sleep(6)
    wb = Workbook.caller()
##    wb.screen_updating = False
##    data_path = r"D:\game\abstract\draw_data_and_post\post\t1.csv"
##    data_path = r"D:\python task\WAC740\IQFact_5G_Tx_SISO_HT20_result.csv"
    t = r"D:\game\abstract\WAC740"
    test_file_5G = [t + r"\IQFact_5G_Tx_SISO_HT40_result.csv",
                 t + r"\IQFact_5G_Tx_SISO_VHT40_result.csv",
                 t + r"\IQFact_5G_Tx_SISO_HT20_result.csv",
                 t + r"\IQFact_5G_Tx_SISO_VHT20_result.csv",
                 t + r"\IQFact_5G_Tx_SISO_VHT80_result.csv",
                 t + r"\IQFact_5G_Tx_SISO_11a_result.csv"]

    test_file_2G = [t + r"\IQFact_2G_Tx_SISO_11b_2484_result.csv",
                 t + r"\IQFact_2G_Tx_SISO_11b_result.csv",
                 t + r"\IQFact_2G_Tx_SISO_11g_result.csv",
                 t + r"\IQFact_2G_Tx_SISO_HT20_result.csv",
                 t + r"\IQFact_2G_Tx_SISO_HT40_result.csv",
                 t + r"\IQFact_2G_Tx_SISO_VHT20_result.csv",
                 t + r"\IQFact_2G_Tx_SISO_VHT40_result.csv"]

    RX_2G = [t + r"\IQFact_2G_Rx_MIMO_result.csv",
             t + r"\IQFact_2G_Rx_SIMO_result.csv",
             t + r"\IQFact_2G_Rx_SISO_result.csv"]

    RX_5G = [t + r"\IQFact_5G_Rx_MIMO_result.csv",
             t + r"\IQFact_5G_Rx_SIMO_result.csv",
             t + r"\IQFact_5G_Rx_SISO_result.csv"]

##    for data_path in RX_2G:
##        print data_path
##        post(data_path)


##    wb.screen_updating = True
##    data_path = test_file_2G[0]
    data_path = RX_5G[2]
##    data_path = t + r"\t.csv"
    post(data_path)
##    a=raw_input()
##    make_folder(os.path.join(relog_path,date),folder_file_names.keys())
##
