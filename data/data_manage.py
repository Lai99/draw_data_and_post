#-------------------------------------------------------------------------------
# Name:        data_manage
# Purpose:
#
# Author:      lai
#
# Created:     04/11/2015
#-------------------------------------------------------------------------------
import xlrd
import math

_item_name_ref = {"standard":"Standard",
                 "channel":"Channel",
                 "BW":"BW",
                 "rate":"rate",
                 "antenna":"ant",
                 "Power":"Measured Power",
                 "EVM":"EVM",
                 "Mask":"Mask",
                 "F_ER":"Frequency Error_ppm",
                 "flatness_inner":"spectralFlatness_InnerSubcarriers",
                 "flatness_outer":"spectralFlatness_OuterSubcarriers",
                 "result":"Test Result"
                }

def _get_standard(table, data, start_pos, _, items_pos):
    if "standard" in items_pos:
        s = table.row_values(start_pos)[items_pos["standard"]]
        if "n" in s:
            data["standard"] = "802.11n"
        elif "ac" in s:
            data["standard"] = "802.11ac"

def _get_channel(table, data, start_pos, _, items_pos):
    if "channel" in items_pos:
##############################################################################
############## HT40/80 channel need to add "2" for real
        if "standard" in items_pos:
            s = table.row_values(start_pos)[items_pos["standard"]]

        if not "20" in s:
            data["channel"] = str(int(table.row_values(start_pos)[items_pos["channel"]]) + 2)
##############################################################################
        else:
            data["channel"] = table.row_values(start_pos)[items_pos["channel"]]

def _get_band_width(table, data, start_pos, _, items_pos):
    if "BW" in items_pos:
        data["BW"] = "BW-" + table.row_values(start_pos)[items_pos["BW"]]

def _get_rate(table, data, start_pos, _, items_pos):
    if "rate" in items_pos:
        if "n" in table.row_values(start_pos)[items_pos["rate"]].lower():
            rate = table.row_values(start_pos)[items_pos["rate"]].lower()
            data["rate"] = rate.split("n")[0].upper()
        else:
            data["rate"] = table.row_values(start_pos)[items_pos["rate"]].upper()

def _get_antenna(table, data, start_pos, _, items_pos):
    ant = {1:"0",2:"0,1",3:"0,1,2",4:"0,1,2,3"}

    if "antenna" in items_pos:
        s = table.row_values(start_pos)[items_pos["antenna"]]
        ant_num = int(math.ceil(math.log(int(s),2)))
        data["antenna"] = ant[ant_num]

def _get_stream(table, data, start_pos, group_num, items_pos):
    data["stream"] = str(group_num)

def _get_power(table, data, start_pos, group_num, items_pos):
    if "Power" in items_pos:
        powers = ""
        for i in range(group_num):
            powers += table.row_values(start_pos)[items_pos["Power"]] + ","
        data["Power"] = powers[:-1]

def _get_EVM(table, data, start_pos, group_num, items_pos):
    if "EVM" in items_pos:
        evms = ""
        for i in range(group_num):
            evms += table.row_values(start_pos)[items_pos["EVM"]] + ","
        data["EVM"] = evms[:-1]

def _get_mask(table, data, start_pos, group_num, items_pos):
    if "Mask" in items_pos:
        evms = ""
        for i in range(group_num):
            if table.row_values(start_pos)[items_pos["Mask"]] == "Pass":
                evms += "0.00,"
            else:
                evms += "Fail,"
        data["Mask"] = evms[:-1]

def _get_f_er(table, data, start_pos, group_num, items_pos):
    if "F_ER" in items_pos:
        data["F_ER"] = table.row_values(start_pos)[items_pos["F_ER"]]

def _get_flatness(table, data,start_pos, group_num, items_pos):
    flatness = ""
    flatness_dup = ""

    if "flatness_inner" in items_pos:
        flatness += table.row_values(start_pos)[items_pos["flatness_inner"]] + ","
    if "flatness_outer" in items_pos:
        flatness += table.row_values(start_pos)[items_pos["flatness_outer"]]

    for i in range(group_num):
        flatness_dup += flatness + ":"
    data["Flatness"] = flatness_dup[:-1]

def _get_items_pos(items_list):
    items_pos = {}
    seq = ["standard","channel","rate","BW","antenna","Power","EVM","Mask","F_ER","flatness_inner","flatness_outer","result"]

    for item in seq:
        if _item_name_ref[item] in items_list:
            items_pos[item] = items_list.index(_item_name_ref[item])

    return items_pos

def get_items_value(table, start_pos, group_num, items_pos):
    data = {}
    _get_standard(table, data, start_pos, group_num, items_pos)
    _get_channel(table, data, start_pos, group_num, items_pos)
    _get_band_width(table, data, start_pos, group_num, items_pos)
    _get_rate(table, data, start_pos, group_num, items_pos)
    _get_antenna(table, data, start_pos, group_num, items_pos)
    _get_stream(table, data, start_pos, group_num, items_pos)
    _get_power(table, data, start_pos, group_num, items_pos)
    _get_EVM(table, data, start_pos, group_num, items_pos)
    _get_mask(table, data, start_pos, group_num, items_pos)
    _get_f_er(table, data, start_pos, group_num, items_pos)
    _get_flatness(table, data,start_pos, group_num, items_pos)

    return data

def draw_data(workbook, anchor, group_num):
    table = workbook.sheets()[0]
    items_pos = {}
    data = {}

    # Get all item column pos
    for row in range(table.nrows):
        if table.row_values(row)[0] == anchor:
            # Need item column pos first
            if not items_pos:
                items_pos = _get_items_pos(table.row_values(row))
##                print items_pos
            # Find no item stop search value
            if not table.row_values(row+1)[0]:
                continue
            #
            i = 1; space = row + i*group_num;pass_flag = True
            while pass_flag and table.nrows >= space and table.row_values(space)[items_pos["result"]]:
                for j in range(group_num):
                    # from bottom to top to avoid over list
                    if table.row_values(space-j)[items_pos["result"]] == "Fail":
                        pass_flag = False
##                        print space
                        break
                i += 1
                space = row + i*group_num
            # i now is fail pos, last one is needed
            if i-1 == 1:
                continue
            else:
                if pass_flag:
                    start_pos = row + (i-2)*group_num + 1
                    data = get_items_value(table,start_pos,group_num,items_pos)
                else:
                    start_pos = row + (i-3)*group_num + 1
                    data = get_items_value(table,start_pos,group_num,items_pos)
                yield data
                data = {}

if __name__ == '__main__':
    pass


