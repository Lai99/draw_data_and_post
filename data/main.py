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
import xlrd
import data_manage

def open_workbook(path):
    """
    Open file
    """
    subprocess.Popen(path,shell=True).pid

def draw_data(workbook, path, data_name):
    anchor = "Standard"
    start_pos = (1,1)

    # find data group number
    group_num = get_group_number(data_name)

    # draw data by data pos

    data_manage.draw_data(workbook, anchor,group_num)
##    for data in data_manage.draw_data(workbook, anchor,group_num):
##        save_data(data)


def get_group_number(line):
    if "MIMO" in line:
        return line.split("MIMO")[1][0]
    return None

def save_data():
    pass

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


    # Load all folders

        # Make folder for save data

        # Open the .csv for save data

        # Load all files by file path

            # Open file

            # Call draw data func.

            # Save draw data

            # Close file

    t = r"D:\game\abstract\draw_data_and_post\data\Log"
##    t = r"D:\python task\draw_data_and_post\data\Log"
##    path = t + r"\2G_Tx_MIMO\MFGC_2G_Tx_MIMO4_HT40_combine_output.xlsx"
    path = t + r"\2G_Tx_MIMO\t.xlsx"

    wb = xlrd.open_workbook(path)
    draw_data(wb,path,"MFGC_2G_Tx_MIMO4_HT40_combine_output")




