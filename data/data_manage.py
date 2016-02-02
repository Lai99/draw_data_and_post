#-------------------------------------------------------------------------------
# Name:        data_manage
# Purpose:
#
# Author:      lai
#
# Created:     04/11/2015
#-------------------------------------------------------------------------------
import xlrd

import subprocess, time, sys

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
                 "tx_result":"Test Result",
                 "SENS":"dut_rx_power",
                 "rx_result":"result"
                }

_tx_seq = ["standard","channel","rate","BW","antenna","Power","EVM","Mask","F_ER","flatness_inner","flatness_outer","tx_result"]
_rx_seq = ["standard","channel","rate","BW","antenna","rx_result","SENS"]

def _get_standard(table, data, start_pos, _, items_pos):
    if "standard" in items_pos:
        s = table.row_values(start_pos)[items_pos["standard"]].lower()
        if "n" in s:
            data["standard"] = "802.11n"
        elif "ac" in s:
            data["standard"] = "802.11ac"
        elif "11a" in s:
            data["standard"] = "802.11ag"
        elif "11b" in s:
            data["standard"] = "802.11b"
        elif "11g" in s:
            data["standard"] = "802.11ag"

def _get_channel(table, data, start_pos, _, items_pos):
    if "channel" in items_pos:
##############################################################################
############## TX HT40 channel need to add "2" for real, TX HT80 channel need to add 2 or 6
        if "standard" in items_pos:
            s = table.row_values(start_pos)[items_pos["standard"]]
##        print start_pos, table.row_values(start_pos)[items_pos["channel"]]
        if int(table.row_values(start_pos)[items_pos["channel"]]) > 30:   #5G channel
            if "40" in s:
                data["channel"] = str(int(table.row_values(start_pos)[items_pos["channel"]]) + 2)

            elif "80" in s:
                if "120" in table.row_values(start_pos)[items_pos["channel"]]:
                    data["channel"] = str(int(table.row_values(start_pos)[items_pos["channel"]]) + 2)
                elif "157" in table.row_values(start_pos)[items_pos["channel"]]:
                    data["channel"] = str(int(table.row_values(start_pos)[items_pos["channel"]]) - 2)
                else:
                    data["channel"] = str(int(table.row_values(start_pos)[items_pos["channel"]]) + 6)
            else:
                data["channel"] = table.row_values(start_pos)[items_pos["channel"]]
##############################################################################
        else:
            if "40" in s:
                data["channel"] = str(int(table.row_values(start_pos)[items_pos["channel"]]) + 2)
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
            data["rate"] = table.row_values(start_pos)[items_pos["rate"]].replace('.',"_").upper()

def _get_antenna(table, data, start_pos, _, items_pos):
    ant = {1:"0",2:"1",3:"0,1",4:"2",5:"0,2",6:"1,2",7:"0,1,2",
           8:"3",9:"0,3",10:"1,3",11:"0,1,3",12:"2,3",13:"0,2,3",
           14:"1,2,3",15:"0,1,2,3"}

    if "antenna" in items_pos:
        s = table.row_values(start_pos)[items_pos["antenna"]]
        data["antenna"] = ant[int(s)]

def _get_stream_tx(table, data, start_pos, group_num, items_pos):
    data["stream"] = str(group_num)

def _get_stream_rx(table, data, start_pos, group_num, items_pos):
    if data["standard"] == "802.11n":
        data["stream"] = str(1)
    else:
        data["stream"] = str(group_num)

def _get_power(table, data, start_pos, group_num, items_pos):
    if "Power" in items_pos:
        powers = ""
        for i in range(group_num):
            powers += table.row_values(start_pos + i)[items_pos["Power"]] + ","
        data["Power"] = powers[:-1]

def _get_EVM(table, data, start_pos, group_num, items_pos):
    if "EVM" in items_pos:
        evms = ""
        for i in range(group_num):
            evms += table.row_values(start_pos + i)[items_pos["EVM"]] + ","
        data["EVM"] = evms[:-1]

def _get_mask(table, data, start_pos, group_num, items_pos):
    if "Mask" in items_pos:
        evms = ""
        for i in range(group_num):
##            print table.row_values(start_pos + i)[items_pos["Mask"]]
            if table.row_values(start_pos + i)[items_pos["Mask"]] == "Pass":
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
        if table.row_values(start_pos)[items_pos["flatness_inner"]]:
            flatness += table.row_values(start_pos)[items_pos["flatness_inner"]] + ","
    if "flatness_outer" in items_pos:
        if table.row_values(start_pos)[items_pos["flatness_outer"]]:
            flatness += table.row_values(start_pos)[items_pos["flatness_outer"]]
    if flatness:
        for i in range(group_num):
            flatness_dup += flatness + ":"
        data["Flatness"] = flatness_dup[:-1]

def _get_sens(table, data,start_pos, group_num, items_pos):
    if "SENS" in items_pos:
        data["SENS"] = table.row_values(start_pos)[items_pos["SENS"]]

def _get_group_number(num):
    ant = {"1":1,"2":1,"3":2,"4":1,"5":2,"6":2,"7":3,
           "8":1,"9":2,"10":2,"11":3,"12":2,"13":3,
           "14":3,"15":4}

    return ant[num]


_get_func = {"standard":_get_standard,
             "channel":_get_channel,
             "rate":_get_rate,
             "BW":_get_band_width,
             "antenna":_get_antenna,
             "stream_tx":_get_stream_tx,
             "stream_rx":_get_stream_rx,
             "Power":_get_power,
             "EVM":_get_EVM,
             "Mask":_get_mask,
             "F_ER":_get_f_er,
             "flatness":_get_flatness,
             "SENS":_get_sens
            }

def _get_tx_items_pos(items_list):
    items_pos = {}

    for item in _tx_seq:
        if _item_name_ref[item] in items_list:
            items_pos[item] = items_list.index(_item_name_ref[item])

    return items_pos

def _get_rx_items_pos(items_list):
    items_pos = {}

    for item in _rx_seq:
        if _item_name_ref[item] in items_list:
            items_pos[item] = items_list.index(_item_name_ref[item])

    return items_pos

def _get_tx_items_value(table, start_pos, group_num, items_pos):
    data = {}

    tx_items = ["standard","channel_tx","rate","BW","antenna","stream_tx","Power","EVM","Mask","F_ER","flatness"]

    for item in tx_items:
        _get_func[item](table, data, start_pos, group_num, items_pos)

    return data

def _get_rx_items_value(table, start_pos, group_num, items_pos):
    data = {}

    rx_items = ["standard","channel","rate","BW","antenna","stream_rx","SENS"]

    for item in rx_items:
        _get_func[item](table, data, start_pos, group_num, items_pos)

    return data

def tx_draw_data(workbook, anchor, group_num):
    table = workbook.sheets()[0]
    items_pos = {}
    data = {}

    # Get all item column pos
    for row in range(table.nrows):
        if table.row_values(row)[0] == anchor:
            # Find no item stop search value
            if not table.row_values(row+1)[0]:
                continue
            #
##            # Need item column pos first
##            if not items_pos:
##                items_pos = _get_items_pos(table.row_values(row))
##                print items_pos
            # Not ensure every item placments are the same, so need get every item pos
            items_pos = _get_tx_items_pos(table.row_values(row))
            # MIMO need to set "stream"
            if group_num != 1:
                ant_pos = items_pos["antenna"]
                group_num = _get_group_number(table.row_values(row+1)[ant_pos])
            else:
                group_num = int(group_num)

            i = 1; space = row + i*group_num;pass_flag = True;pass_row = 0
            while table.nrows > space and table.row_values(space)[items_pos["tx_result"]]:
                check_pass = True
                for j in range(group_num):
                    # from bottom to top to avoid over list
                    if table.row_values(space-j)[items_pos["tx_result"]] == "Fail":
                        check_pass = False
                        break
                # Find all pass save the row pos
                if check_pass:
                    pass_row = space - group_num + 1

                i += 1
                space = row + i*group_num

            if pass_row != 0:
                data = _get_tx_items_value(table,pass_row,group_num,items_pos)

                yield data
                data = {}

def rx_draw_data(workbook, anchor,group_num):
    table = workbook.sheets()[0]
    items_pos = {}
    data = {}

    # Get all item column pos
    for row in range(table.nrows):
        if table.row_values(row)[0] == anchor:
            # Find no item stop search value
##            print row
            if not table.row_values(row+1)[0]:
                continue

            # Not ensure every item placments are the same, so need get every item pos
            items_pos = _get_rx_items_pos(table.row_values(row))
            # MIMO need to set "stream"
            if group_num != 1:
                ant_pos = items_pos["antenna"]
                group_num = _get_group_number(table.row_values(row+1)[ant_pos])
            else:
                group_num = int(group_num)
##            print items_pos
            space = row + 1;pass_flag = True;pass_row = 0
            while table.nrows > space and table.row_values(space)[items_pos["rx_result"]]:
                check_pass = True

                if table.row_values(space)[items_pos["rx_result"]] == "Fail":
                    check_pass = False

                if check_pass:
                    pass_row = space

                space += 1

            if pass_row != 0:
                data = _get_rx_items_value(table,pass_row,group_num,items_pos)

                yield data
                data = {}


if __name__ == '__main__':
    a=raw_input()
    sys.exit(0)