#-------------------------------------------------------------------------------
# Name:        main
# Purpose:
#
# Author:      Lai
#
# Created:     04/11/2015
#-------------------------------------------------------------------------------
import os, sys
import time
import datetime
import xlrd
import csv
import data_manage

def get_folder_filenames(path):
    """
    Get all file names with folder under the path
    Input:  string:file_path
    Output: dict:{folder,file_names}
    """
    result = {}
    for parent, dirnames, filenames in os.walk(path):
##        print dirnames, filenames
        if not dirnames:
            result[os.path.split(parent)[-1]] = filenames
    return result

def _get_value(data,item):
    if item in data:
        return data[item]
    else:
        return ""

def _get_item(data,item):
    if _get_value(data,item):
        return [item, _get_value(data,item)]
    return []

def _make_data_config(data):
    config = []
    seq = ["standard","channel","rate","BW","stream","antenna"]

    for item in seq:
        config.append(_get_value(data,item))
    return config

def _get_mode(data_name):
    mode = None
    if "tx" in data_name.lower():
        mode = "tx"
    elif "rx" in data_name.lower():
        mode = "rx"

    return mode

def _get_group_number(line):
    if "SISO" in line.upper() or "SIMO" in line.upper():
        return 1
    return None

def draw_data(workbook, save_path, data_name):
    anchor = "Standard"
    title = [" Number"," standard"," Freq"," Rate"," BW"," Stream"," Ant","  Item","  Vaule"]
    count = 1
    seq = ["Measured_Power","EVM","Mask","F_ER","Phase_Noise","Lo_Leakage","Flatness"]
    # get "TX or RX"
    mode = _get_mode(data_name)
    if not mode:
        print "!!!Alert: File name needs to describe \"TX or RX\". This file won't be executed"
        return 1

    # find data group number
    group_num = _get_group_number(data_name)

    # Open .csv for saving purpose
    file_path = os.path.join(save_path,data_name)
    with open(file_path+".csv", 'wb') as csvfile: # must use binary "b"
        data_writer = csv.writer(csvfile)
        data_writer.writerow(title)
        # draw data by data pos
        if mode == "tx":
            for data in data_manage.tx_draw_data(workbook, anchor, group_num):
                # Save draw data
    ##            print data
                config = _make_data_config(data)
                line = list([count]) + config + _get_item(data,"Target_Power")
                data_writer.writerow(line)

                for item in seq:
                    if _get_item(data,item):
                        line = [""] * (len(title)-2) + _get_item(data,item)
                        data_writer.writerow(line)
                data_writer.writerow("")
                count += 1

        elif mode == "rx":
            for data in data_manage.rx_draw_data(workbook, anchor, group_num):
                # Save draw data
    ##            print data
                config = _make_data_config(data)
                line = list([count]) + config + _get_item(data,"SENS")
                data_writer.writerow(line)

                data_writer.writerow("")
                count += 1

def make_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)

def main():
    # the program path
    rootdir = os.path.dirname(__file__)
##rootdir = os.path.dirname(os.path.abspath(sys.argv[0]))
    # log folder path
    log_path = os.path.join(rootdir,"Log")
    # report folder path
    result_path = os.path.join(rootdir,"Result")
    # Get all files with dict (folder:file_name)
    folder_file_names = get_folder_filenames(log_path)
    t = time.time()
    # today in year/month/day
    date = datetime.datetime.fromtimestamp(t).strftime(r"%Y%m%d")


    # Load all folders
    for folder in folder_file_names:
        # Make folder for save data
        folder_path = os.path.join(result_path, date, folder)
        make_folder(folder_path)
        # Load all files by file path
        try:
            for data_name in folder_file_names[folder]:
                # Open raw file
                if not (data_name.split(".")[-1] in "xls" or data_name.split(".")[-1] in "xlsx"):
                    print "!!!ALert: Not support this file type " + data_name
                    continue
                data_path = os.path.join(log_path,folder,data_name)
                print data_path
                wb = xlrd.open_workbook(data_path)
                # Call draw data func. and save arranged one
                draw_data(wb,folder_path,data_name.split(".")[0])
        except:
            print "Unexpected error:", sys.exc_info()[0]

    os.system("pause")

if __name__ == '__main__':
    main()
