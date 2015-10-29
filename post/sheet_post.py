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
import data_mange

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
                  "Flatness":"Flatness"
                 }

def post_power(data):
    return data["power"].split(",")

def post_evm(data):
    return data["EVM"].split(",")

def post_mask(data):
    return data["Mask"].split(",")

def post_freq_err(data):
    return list([data["F_ER"]])

def post_cr_err(data):
    return list([data["CR_ER"]])

def post_flatness(data):
    return data["Flatness"].split(":")

post_func = {"power":post_power,
            "EVM":post_evm,
            "Mask":post_mask,
            "F_ER":post_freq_err,
            "CR_ER":post_cr_err,
            "Flatness":post_flatness
            }

def post(data_path):
    #
    sheet = 2
    #
    standard_anchor = "Standard"
    channel_anchor = "Ch"
    fill_pos = template_search.get_fill_pos(sheet,standard_anchor)
    last_data_conf = None
    need_pos = None
    case_num = 0

    for data in data_mange.load_data(data_path):
        if not check_same_row(data, last_data_conf):
##            print data
##            print fill_pos.keys()
            m = meet_standard(data, fill_pos)
##            print m
            if meet_rate(data, m):
                need_pos, case_num = meet_rate(data, m)
                last_data_conf = get_data_conf(data)
            else:
                print "Can't find this modulation " + data[item_ref["rate"]]
                continue
##        print need_pos, case_num
        ch_start = template_search.get_channel_start(sheet,need_pos,channel_anchor)
##        print ch_start
        ch_pos = template_search.get_channel_pos(sheet,ch_start,data[item_ref["channel"]])
##        print ch_pos
        if ch_pos:
            post_value(sheet,data,need_pos,ch_pos,case_num)
        else:
            print "Can't find this channel " + data[item_ref["channel"]]
            continue

def check_same_row(data, last_data_conf):
    if not last_data_conf:
        return False

    conf = get_data_conf(data)
    if conf == last_data_conf:
        return True
    return False

def get_data_conf(data):
    return (data["standard"],data["rate"],data["BW"],data["stream"],data["antenna"])

def post_value(sheet,data,start,ch_pos,case_num):
    for i in range(case_num):
        case = Range(sheet,(start[0]+i,start[1]-1)).value
        if case in sheet_item_ref.keys():
            value = post_func[sheet_item_ref[case]](data)
            antennas = data["antenna"].split(",")
            if len(value) > 1:
                for s in range(int(data["stream"])):
                    post_pos = (start[0]+i,ch_pos[1]+int(antennas[s]))
                    Range(sheet,post_pos).value = value[s]
            else:
                post_pos = (start[0]+i,ch_pos[1]+int(antennas[0]))
                Range(sheet,post_pos).value = value[0]

def meet_standard(data,fill_pos):
    k = (data[item_ref["standard"]], data[item_ref["BW"]], data[item_ref["stream"]])

    if k in fill_pos:
        return fill_pos[k]

    if k[0] in fill_pos:
        return fill_pos[k]
    return None

def meet_rate(data,fill_pos):
    for k in fill_pos.keys():
        if data[item_ref["rate"]] in k:
            return fill_pos[k]
    return None

if __name__ == '__main__':
    pass
##    a = {u'11a': {u'OFDM - QPSK 18.0': ((26, 6), 6), u'OFDM - 64 QAM 54.0': ((50, 6), 6), u'OFDM - 64 QAM 48.0': ((44, 6), 6), u'OFDM - QPSK 12.0': ((20, 6), 6), u'OFDM - BPSK 9.0': ((14, 6), 6), u'OFDM - BPSK 6.0': ((8, 6), 6), u'OFDM - 16 QAM 24.0': ((32, 6), 6), u'OFDM - 16 QAM 36.0': ((38, 6), 6)}, (u'11ac', u'VHT20', u'3'): {u'MCS7 - OFDM - 64 - QAM None': ((352, 6), 6), u'MCS8 - OFDM - 256 - QAM None': ((358, 6), 6), u'MCS1 - OFDM - QPSK None': ((316, 6), 6), u'MCS6 - OFDM - 64 - QAM None': ((346, 6), 6), u'MCS4 - OFDM - 16 - QAM None': ((334, 6), 6), u'MCS9 - OFDM - 256 - QAM None': ((364, 6), 6), u'MCS3 - OFDM - 16 - QAM None': ((328, 6), 6), u'MCS5 - OFDM - 64 - QAM None': ((340, 6), 6), u'MCS0 - OFDM - BPSK None': ((310, 6), 6), u'MCS2 - OFDM - QPSK None': ((322, 6), 6)}, (u'11n', u'HT20', u'3'): {u'MCS17 - OFDM - QPSK None': ((213, 6), 6), u'MCS22 - OFDM - 64QAM None': ((243, 6), 6), u'MCS20 - OFDM - 16QAM None': ((231, 6), 6), u'MCS16 - OFDM - BPSK None': ((207, 6), 6), u'MCS19 - OFDM - 16QAM None': ((225, 6), 6), u'MCS21 - OFDM - 64QAM None': ((237, 6), 6), u'MCS18 - OFDM - QPSK None': ((219, 6), 6), u'MCS23 - OFDM - 64QAM None': ((249, 6), 6)}, (u'11ac', u'VHT80', u'3'): {u'MCS7 - OFDM - 64 - QAM None': ((860, 6), 6), u'MCS8 - OFDM - 256 - QAM None': ((866, 6), 6), u'MCS1 - OFDM - QPSK None': ((830, 6), 6), u'MCS4 - OFDM - 16 - QAM None': ((848, 6), 6), u'MCS9 - OFDM - 256 - QAM None': ((872, 6), 6), u'MCS3 - OFDM - 16 - QAM None': ((842, 6), 6), u'MCS5 - OFDM - 64 - QAM None': ((854, 6), 6), u'MCS0 - OFDM - BPSK None': ((824, 6), 6), u'MCS2 - OFDM - QPSK None': ((836, 6), 6)}, (u'11n', u'HT40', u'2'): {u'MCS8 - OFDM - BPSK None': ((483, 6), 6), u'MCS12 - OFDM - 16QAM None': ((507, 6), 6), u'MCS9 - OFDM - QPSK None': ((489, 6), 6), u'MCS10 - OFDM - QPSK None': ((495, 6), 6), u'MCS13 - OFDM - 64QAM None': ((513, 6), 6), u'MCS14 - OFDM - 64QAM None': ((519, 6), 6), u'MCS11 - OFDM - 16QAM None': ((501, 6), 6), u'MCS15 - OFDM - 64QAM None': ((525, 6), 6)}, (u'11ac', u'VHT40', u'3'): {u'MCS7 - OFDM - 64 - QAM None': ((681, 6), 6), u'MCS8 - OFDM - 256 - QAM None': ((687, 6), 6), u'MCS1 - OFDM - QPSK None': ((645, 6), 6), u'MCS6 - OFDM - 64 - QAM None': ((675, 6), 6), u'MCS4 - OFDM - 16 - QAM None': ((663, 6), 6), u'MCS9 - OFDM - 256 - QAM None': ((693, 6), 6), u'MCS3 - OFDM - 16 - QAM None': ((657, 6), 6), u'MCS5 - OFDM - 64 - QAM None': ((669, 6), 6), u'MCS0 - OFDM - BPSK None': ((639, 6), 6), u'MCS2 - OFDM - QPSK None': ((651, 6), 6)}, (u'11n', u'HT20', u'2'): {u'MCS8 - OFDM - BPSK None': ((159, 6), 6), u'MCS12 - OFDM - 16QAM None': ((183, 6), 6), u'MCS9 - OFDM - QPSK None': ((165, 6), 6), u'MCS10 - OFDM - QPSK None': ((171, 6), 6), u'MCS13 - OFDM - 64QAM None': ((189, 6), 6), u'MCS14 - OFDM - 64QAM None': ((195, 6), 6), u'MCS11 - OFDM - 16QAM None': ((177, 6), 6), u'MCS15 - OFDM - 64QAM None': ((201, 6), 6)}, (u'11ac', u'VHT20', u'2'): {u'MCS7 - OFDM - 64 - QAM None': ((297, 6), 6), u'MCS8 - OFDM - 256 - QAM None': ((303, 6), 6), u'MCS1 - OFDM - QPSK None': ((261, 6), 6), u'MCS6 - OFDM - 64 - QAM None': ((291, 6), 6), u'MCS4 - OFDM - 16 - QAM None': ((279, 6), 6), u'MCS3 - OFDM - 16 - QAM None': ((273, 6), 6), u'MCS5 - OFDM - 64 - QAM None': ((285, 6), 6), u'MCS0 - OFDM - BPSK None': ((255, 6), 6), u'MCS2 - OFDM - QPSK None': ((267, 6), 6)}, (u'11ac', u'VHT40', u'2'): {u'MCS7 - OFDM - 64 - QAM None': ((621, 6), 6), u'MCS8 - OFDM - 256 - QAM None': ((627, 6), 6), u'MCS1 - OFDM - QPSK None': ((585, 6), 6), u'MCS6 - OFDM - 64 - QAM None': ((615, 6), 6), u'MCS4 - OFDM - 16 - QAM None': ((603, 6), 6), u'MCS9 - OFDM - 256 - QAM None': ((633, 6), 6), u'MCS3 - OFDM - 16 - QAM None': ((597, 6), 6), u'MCS5 - OFDM - 64 - QAM None': ((609, 6), 6), u'MCS0 - OFDM - BPSK None': ((579, 6), 6), u'MCS2 - OFDM - QPSK None': ((591, 6), 6)}, (u'11n', u'HT20', u'1'): {u'MCS2  - OFDM - QPSK 19.5': ((68, 6), 6), u'MCS4  - OFDM - 16QAM 39.0': ((80, 6), 6), u'MCS5  - OFDM - 64QAM 52.0': ((86, 6), 6), u'MCS0  - OFDM - BPSK 6.5': ((56, 6), 6), u'MCS1  - OFDM - QPSK 13.0': ((62, 6), 6), u'MCS3  - OFDM - 16QAM 26.0': ((74, 6), 6), u'MCS6  - OFDM - 64QAM 58.5': ((92, 6), 6), u'MCS7  - OFDM - 64QAM 65.0': ((98, 6), 6)}, (u'11n', u'HT40', u'3'): {u'MCS17 - OFDM - QPSK None': ((537, 6), 6), u'MCS22 - OFDM - 64QAM None': ((567, 6), 6), u'MCS20 - OFDM - 16QAM None': ((555, 6), 6), u'MCS16 - OFDM - BPSK None': ((531, 6), 6), u'MCS19 - OFDM - 16QAM None': ((549, 6), 6), u'MCS21 - OFDM - 64QAM None': ((561, 6), 6), u'MCS18 - OFDM - QPSK None': ((543, 6), 6), u'MCS23 - OFDM - 64QAM None': ((573, 6), 6)}, (u'11n', u'HT40', u'1'): {u'MCS5  - OFDM - 64QAM 108.0': ((405, 6), 6), u'MCS6  - OFDM - 64QAM 121.5': ((411, 6), 6), u'MCS4  - OFDM - 16QAM 81.0': ((399, 6), 6), u'MCS7  - OFDM - 64QAM 135.0': ((417, 6), 6), u'MCS1  - OFDM - QPSK 27.0': ((381, 6), 6), u'MCS0  - OFDM - BPSK 13.5': ((375, 6), 6), u'MCS3  - OFDM - 16QAM 54.0': ((393, 6), 6), u'MCS2  - OFDM - QPSK 40.5': ((387, 6), 6)}, (u'11ac', u'VHT20', u'1'): {u'MCS8 - OFDM - 256 - QAM 351.0': ((152, 6), 6), u'MCS5 - OFDM - 64 - QAM 234.0': ((134, 6), 6), u'MCS7 - OFDM - 64 - QAM 292.5': ((146, 6), 6), u'MCS6 - OFDM - 64 - QAM 263.3': ((140, 6), 6), u'MCS4 - OFDM - 16 - QAM 175.5': ((128, 6), 6), u'MCS3 - OFDM - 16 - QAM 117.0': ((122, 6), 6), u'MCS1 - OFDM - QPSK 58.5': ((110, 6), 6), u'MCS2 - OFDM - QPSK 87.8': ((116, 6), 6), u'MCS0 - OFDM - BPSK 29.3': ((104, 6), 6)}, (u'11ac', u'VHT80', u'1'): {u'MCS5 - OFDM - 64 - QAM 702.0': ((734, 6), 6), u'MCS8 - OFDM - 256 - QAM 1053.0': ((752, 6), 6), u'MCS6 - OFDM - 64 - QAM 0.0': ((740, 6), 6), u'MCS0 - OFDM - BPSK 87.8': ((704, 6), 6), u'MCS1 - OFDM - QPSK 175.5': ((710, 6), 6), u'MCS4 - OFDM - 16 - QAM 526.5': ((728, 6), 6), u'MCS3 - OFDM - 16 - QAM 351.0': ((722, 6), 6), u'MCS2 - OFDM - QPSK 263.3': ((716, 6), 6), u'MCS9 - OFDM - 256 - QAM 1170.0': ((758, 6), 6), u'MCS7 - OFDM - 64 - QAM 877.5': ((746, 6), 6)}, (u'11ac', u'VHT80', u'2'): {u'MCS7 - OFDM - 64 - QAM None': ((806, 6), 6), u'MCS8 - OFDM - 256 - QAM None': ((812, 6), 6), u'MCS1 - OFDM - QPSK None': ((770, 6), 6), u'MCS6 - OFDM - 64 - QAM None': ((800, 6), 6), u'MCS4 - OFDM - 16 - QAM None': ((788, 6), 6), u'MCS9 - OFDM - 256 - QAM None': ((818, 6), 6), u'MCS3 - OFDM - 16 - QAM None': ((782, 6), 6), u'MCS5 - OFDM - 64 - QAM None': ((794, 6), 6), u'MCS0 - OFDM - BPSK None': ((764, 6), 6), u'MCS2 - OFDM - QPSK None': ((776, 6), 6)}, (u'11ac', u'VHT40', u'1'): {u'MCS4  - OFDM - 16QAM None': ((447, 6), 6), u'MCS1  - OFDM - QPSK None': ((429, 6), 6), u'MCS2  - OFDM - QPSK None': ((435, 6), 6), u'MCS0  - OFDM - BPSK None': ((423, 6), 6), u'MCS8 - OFDM - 256 - QAM None': ((471, 6), 6), u'MCS3  - OFDM - 16QAM None': ((441, 6), 6), u'MCS5  - OFDM - 64QAM None': ((453, 6), 6), u'MCS6  - OFDM - 64QAM None': ((459, 6), 6), u'MCS9 - OFDM - 256 - QAM None': ((477, 6), 6), u'MCS7  - OFDM - 64QAM None': ((465, 6), 6)}}
##    path = (r"D:\python task\draw_data_and_post\post\TX.csv")
##    post(path,a)
