#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Admin
#
# Created:     28/10/2015
#-------------------------------------------------------------------------------

from xlwings import Workbook, Sheet, Range, Chart
import template_search
import data_manage

item_ref = {"standard":"standard",
            "rate":"rate",
            "channel":"channel",
            "stream":"stream",
            "BW":"BW"
           }

sheet_item_ref = {"Tx Power":"power",
                  "EVM":"EVM",
                  "Mask":"Mask",
                  "Freq error":"F_ER",
                  "SC error":"CR_ER",
                  "Flatness":"Flatness",
                  "Rx Power":"SENS"
                 }

rx_item = ["SENS"]

def post_power(data):
    if data["power"]:
        return data["power"].split(",")
    return data["power"]

def post_evm(data):
    if data["EVM"]:
        return data["EVM"].split(",")
    return data["EVM"]

def post_mask(data):
    if data["Mask"]:
        return data["Mask"].split(",")
    return data["Mask"]

def post_freq_err(data):
    if data["F_ER"]:
        return list([data["F_ER"]])
    return data["F_ER"]

def post_cr_err(data):
    if data["CR_ER"]:
        return list([data["CR_ER"]])
    return data["CR_ER"]

def post_flatness(data):
    if data["Flatness"]:
        return data["Flatness"].split(":")
    return data["Flatness"]

def post_sens(data):
    if data["sens"]:
        return data["sens"].split(",")
    return data["sens"]

post_func= {# TX
            "power":post_power,
            "EVM":post_evm,
            "Mask":post_mask,
            "F_ER":post_freq_err,
            "CR_ER":post_cr_err,
            "Flatness":post_flatness,
            # RX
            "SENS":post_sens
            }

def post(data_path, sheet, sheet_setup, channel_anchor):
    fill_pos, all_anchor_row = sheet_setup[0], sheet_setup[1]
##    print fill_pos.keys()
##    print all_anchor_row
    last_data_conf = None
    need_pos = None
    case_num = 0
    ch_start = None
    ch_pos = None

    for data in data_manage.load_data(data_path):
        Sheet(sheet).activate()
        if not check_same_row(data, last_data_conf):
##            print data
##            print fill_pos.keys()
            m = meet_standard(data, fill_pos)
##            print m
            if meet_rate(data, m):
                need_pos, case_num = meet_rate(data, m)
                last_data_conf = get_data_conf(data)
            else:
                continue
            print need_pos, case_num
            ch_start = template_search.get_channel_start(sheet,need_pos,channel_anchor,all_anchor_row)
        print ch_start, "ch start"
        ch_pos = template_search.get_channel_pos(sheet,ch_start,data[item_ref["channel"]])
        print ch_pos, "ch pos"
        if ch_pos:
            post_value(sheet,data,need_pos,ch_pos,case_num)
            ch_start = ch_pos
        else:
            continue

def check_same_row(data, last_data_conf):
    if not last_data_conf:
        return False

    conf = get_data_conf(data)
    if conf == last_data_conf:
        return True
    return False

def get_data_conf(data):
    return (data["standard"],data["rate"],data["BW"],data["stream"])

def post_value(sheet,data,start,ch_pos,case_num):
    for i in range(case_num):
        case = Range(sheet,(start[0]+i,start[1]-1)).value
##        print case
        if case in sheet_item_ref.keys():
            value = post_func[sheet_item_ref[case]](data)
##            print data
##            print value
            antennas = data["antenna"].split(",")
##            print antennas, len(antennas)
            if value:
                if len(value) > 1:
                    for idx in range(int(data["stream"])):
##                    for idx in range(len(antennas)):
                        post_pos = (start[0]+i,ch_pos[1]+int(antennas[idx]))
                        Range(sheet,post_pos).value = value[idx]
                else:
                    if len(antennas) > 1:
###################### for RX 11ac MIMO and SIMO post #######################
                        if sheet_item_ref[case] in rx_item:
                            move = template_search.find_ch_sum(sheet,ch_pos)
    ##                        print move
                            post_pos = (start[0]+i,ch_pos[1]+move)
#############################################################################
                        else:
                            # TX no need move
                            post_pos = (start[0]+i,ch_pos[1])
                    else:
                        post_pos = (start[0]+i,ch_pos[1]+int(antennas[0]))
                    Range(sheet,post_pos).value = value[0]

def meet_standard(data,fill_pos):
############ RX 11n need this to let "stream" to meet really config ###################
############ MIMO will always get stream '1'
    if data[item_ref["standard"]] == "11n":
        ch = int(data[item_ref["rate"]][3:])
        if ch < 8:
            stream = '1'
        elif ch < 16:
            stream = '2'
        elif ch < 24:
            stream = '3'
        else:
            stream = '4'
        k = (data[item_ref["standard"]], data[item_ref["BW"]], stream)
#######################################################################################
    else:
        k = (data[item_ref["standard"]], data[item_ref["BW"]], data[item_ref["stream"]])

    if k in fill_pos:
        return fill_pos[k]

    if k[0] in fill_pos:
        return fill_pos[k[0]]
    print "Can't find this channel " + data[item_ref["channel"]]
    return None

def meet_rate(data,fill_pos):
##    print fill_pos.keys()
    for k in fill_pos.keys():
        if "-" in data[item_ref["rate"]]:
            # Use "in" not "==" becasue the data["rate"] will look like "OFDM-6" but key is "OFDM-6.0"
            if data[item_ref["rate"]] in k:
                return fill_pos[k]
        else:
            if "-" in k:
                if data[item_ref["rate"]] == k.split("-")[0]:
                    return fill_pos[k]
            else:
                if data[item_ref["rate"]] == k:
                    return fill_pos[k]
    print "Can't find this modulation " + data[item_ref["rate"]]
    return None

if __name__ == '__main__':
    pass
