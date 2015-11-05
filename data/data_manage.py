#-------------------------------------------------------------------------------
# Name:        data_manage
# Purpose:
#
# Author:      lai
#
# Created:     04/11/2015
#-------------------------------------------------------------------------------
import xlrd

def _get_standard():
    pass

_item_name_ref = {"standard":"Standard",
                 "channel":"Channel",
                 "BW":"BW",
                 "rate":"rate",
                 "antenna":"ant",
                 "power":"Measured Power",
                 "EVM":"EVM",
                 "mask":"Mask",
                 "F_ER":"Frequency Error_ppm",
                 "flatness_inner":"spectralFlatness_InnerSubcarriers",
                 "flatness_outer":"spectralFlatness_OuterSubcarriers"
                }

_get_func = {"standard":_get_standard,
}


def _get_items_pos(items_list):
    items_pos = {}
    # get "standard" column pos
    if _item_name_ref["standard"] in items_list:
        items_pos["standard"] = items_list.index(_item_name_ref["standard"])
    # get "channel column pos
    if _item_name_ref["channel"] in items_list:
        items_pos["channel"] = items_list.index(_item_name_ref["channel"])
    # get "band width" column pos
    if _item_name_ref["BW"] in items_list:
        items_pos["BW"] = items_list.index(_item_name_ref["BW"])
    # get "rate" column pos
    if _item_name_ref["rate"] in items_list:
        items_pos["rate"] = items_list.index(_item_name_ref["rate"])
    # get "antenna" column pos
    if _item_name_ref["antenna"] in items_list:
        items_pos["antenna"] = items_list.index(_item_name_ref["antenna"])
    # get "power" column po
    if _item_name_ref["power"] in items_list:
        items_pos["power"] = items_list.index(_item_name_ref["power"])
    # get "EVM" column pos
    if _item_name_ref["EVM"] in items_list:
        items_pos["EVM"] = items_list.index(_item_name_ref["EVM"])
    #get "mask" column pos
    if _item_name_ref["mask"] in items_list:
        items_pos["mask"] = items_list.index(_item_name_ref["mask"])
    # get "frequency error" column pos
    if _item_name_ref["F_ER"] in items_list:
        items_pos["F_ER"] = items_list.index(_item_name_ref["F_ER"])
    # get "flatness inner" column pos
    if _item_name_ref["flatness_inner"] in items_list:
        items_pos["flatness_inner"] = items_list.index(_item_name_ref["flatness_inner"])
    # get "flatness outer" column pos
    if _item_name_ref["flatness_outer"] in items_list:
        items_pos["flatness_outer"] = items_list.index(_item_name_ref["flatness_outer"])

    return items_pos

def draw_data(workbook, anchor, group_num):
    table = workbook.sheets()[0]
    items_pos = {}
    r = 0
    while table.row_values(r)[0]:
        if table.row_values(r)[0] == anchor:
##            print r, table.row_values(r)
            items_pos = _get_items_pos(table.row_values(r))
            print items_pos
            break
        r+=1

if __name__ == '__main__':
##    pass
    t = r"D:\python task\draw_data_and_post\data\Log"
    path = t + r"\2G_Tx_MIMO\MFGC_2G_Tx_MIMO4_HT40_combine_output.xlsx"

    wb = xlrd.open_workbook(path)
    t = wb.sheets()[0]
    print t.nrows
    print t.ncols

    for row in range(t.nrows):
        if "Standard" in t.row_values(row):
            print row

