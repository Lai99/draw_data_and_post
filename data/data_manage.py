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
                 "flatness_outer":"spectralFlatness_OuterSubcarriers",
                 "result":"Test Result"
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

    # get "Test result" column pos
    if _item_name_ref["result"] in items_list:
        items_pos["result"] = items_list.index(_item_name_ref["result"])

    return items_pos

def draw_data(workbook, anchor, group_num):
    table = workbook.sheets()[0]
    items_pos = {}

    # Get all item column pos
    for row in range(table.nrows):
        if table.row_values(row)[0] == anchor:
            # Need item column pos first
            if not items_pos:
                items_pos = _get_items_pos(table.row_values(row))
##                print items_pos
            # Sea
            i = 1; space = row + i*int(group_num);pass_flag = True
            while pass_flag and table.nrows >= space and table.row_values(space)[items_pos["result"]]:
                for j in range(int(group_num)):
                    # from bottom to top to avoid over list
                    if table.row_values(space-j)[items_pos["result"]] == "Fail":
                        pass_flag = False
##                        print space
                        break
                i += 1
                space = row + i*int(group_num)
            # i now is fail pos, last one is needed
            if i-1 == 1:
                pass
            else:
                if pass_flag:
                    print row + (i-1)*int(group_num), row
                else:
                    print row + (i-2)*int(group_num), row

if __name__ == '__main__':
    pass


