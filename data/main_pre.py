#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Lai
#
# Created:     19/10/2015
#-------------------------------------------------------------------------------
import os
import time
import datetime
import data_mange
import draw_item

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

def save_data(path,data):
    f = open(path,'w')
    f.write(data)
    f.close

def data_to_str(data):
    """
    Get data and turn to .csv format
    """
    s = ""
    items = ""
    s += draw_item.draw_config(data)
    items = draw_item.draw_items(data)
    for i in range(len(items)):
        if i == 0:
            s += items[i] + "\n"
        else:
            t = len(title().split(",")) - 2
            s += "," * t + items[i] + "\n"
    return s + "\n"

def title():
    return " Number, standard, Freq, Rate, BW, Stream, Ant,  Item,  Vaule\n"

if __name__ == '__main__':
    rootdir = os.path.dirname(__file__)
    log_path = os.path.join(rootdir,"Log")
    relog_path = os.path.join(rootdir,"ReLog")
    folder_file_names = get_folder_filenames(log_path)
    anchor = "TX_MULTI_VERIFICATION"
    t = time.time()
    date = datetime.datetime.fromtimestamp(t).strftime(r"%Y%m%d")

    make_folder(os.path.join(relog_path,date),folder_file_names.keys())

    for folder in folder_file_names.keys():
        for f in folder_file_names[folder]:
            format_str = ""
            format_str += title()
            data = data_mange.make_modifed_data(os.path.join(log_path,folder,f),anchor)
            for d in data:
                format_str += data_to_str(d)

            save_data(os.path.join(relog_path,date,folder,(f.split(".")[0] + ".csv")),format_str)
