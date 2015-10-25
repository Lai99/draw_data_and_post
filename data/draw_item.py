#-------------------------------------------------------------------------------
# Name:        draw_item
# Purpose:     draw item and value from format input data
#
# Author:      Lai
#
# Created:     20/10/2015
#-------------------------------------------------------------------------------
items_name={"CONFIG":"CONFIG",
            "STANDARD":"STANDARD",
            "FREQ":"CH_FREQ_MHZ",
            "DATA_RATE":"DATA_RATE",
            "BANDWIDTH":"BSS_BANDWIDTH",
            "TX":"TX",
            "POWER":"POWER_DBM_RMS_AVG_VSA",
            "EVM":"EVM_DB_AVG_S",
            "VIOLATION_PERCENT":"VIOLATION_PERCENT_VSA_",
            "FREQ_ERROR":"FREQ_ERROR_AVG_ALL",
            "CLK_ERR":"SYMBOL_CLK_ERR_ALL",
            "Flatness":["VALUE_DB_LO_A_VSA","VALUE_DB_LO_B_VSA","VALUE_DB_UP_A_VSA","VALUE_DB_UP_B_VSA"]
            }


def draw_config(content):
    """
    Get config and return in .csv format
    """
    s = ""
    #Get raw coonfig(string) and split into individual item(list)
    config = content[items_name["CONFIG"]].split()
    #Get channel number and add
    s += ((config[0].split("."))[0] + ",")
    #Get IEEE standard
    s += (content[items_name["STANDARD"]] + ",")
    #Get freq.
    s += (draw_freq(content) + ",")
    #Get MCS
    s += (config[2] + ",")
    #Get band width
    s += (content[items_name["BANDWIDTH"]] + ",")
    #Get stream
    s += (str(draw_stream(content)) + ",")
    #Get antenna number
    s += ("\"" + draw_antenna(content) + "\"" + ",")

    return s

def draw_items(content):
    """
    Get item and value in string and return a list with those
    """
    item_value = []
    stream = draw_stream(content)
    #Get power
    item_value.append("Power," +"\"" + draw_power(content,stream) + "\"")
    #Get EVM
    item_value.append("EVM," +"\"" + draw_EVM(content,stream) + "\"")
    #Get Mask
    item_value.append("Mask," +"\"" + draw_Mask(content,stream) + "\"")
    #Get freq. error
    item_value.append("F_ER," + get_value(content,items_name["FREQ_ERROR"]))
    #Get symbol clk error
    item_value.append("CR_ER," + get_value(content,items_name["CLK_ERR"]))
    #Get flatness
    item_value.append("Flatness," +"\"" + draw_flatness(content,stream) + "\"")

    return item_value

def draw_freq(content):
    """
    find freq. and turn into channel
    """
    data_rate = int(content[items_name["FREQ"]])
    if data_rate > 5000:  #5GHz
        return str((data_rate - 5000) / 5)
    else:   #2.4GHz
        return str((data_rate - 2407) / 5)

def draw_stream(content):
    """
    Return stream number
    """
    stream = 0
    ch = 1
    k = items_name["TX"] + str(ch)
    while k in content:
        stream += int(content[k])
        ch += 1
        k = items_name["TX"] + str(ch)
    return stream

def draw_antenna(content):
    """
    Return antenna number in string like "0,1,2"
    """
    antenna = ""
    ch = 1
    k = "TX1"
    while k in content:
        if int(content[k]) == 1:
            antenna += str(ch-1)
        ch += 1
        k = "TX" + str(ch)
    return ",".join(antenna)

def draw_power(content,stream):
    """
    Return every stream power in string
    """
    s = ""
    for i in range(stream):
        k = items_name["POWER"] + str(i+1)
        if get_value(content,k):
            s += (get_value(content,k) + ",")
    return s[:-1]

def draw_EVM(content,stream):
    """
    Return every stream EVM in string
    """
    s = ""
    for i in range(stream):
        k = items_name["EVM"] + str(i+1)
        if get_value(content,k):
            s += (get_value(content,k) + ",")
    return s[:-1]

def draw_Mask(content,stream):
    """
    Return every stream mask in string
    """
    s = ""
    for i in range(stream):
        k = items_name["VIOLATION_PERCENT"] + str(i+1)
        if get_value(content,k):
            s += (get_value(content,k) + ",")
    return s[:-1]

def draw_flatness(content,stream):
    """
    Return every stream flatness in string
    """
    s = ""
    for i in range(stream):
        for k in items_name["Flatness"]:
            if get_value(content,k+str(i+1)):
                s += get_value(content,(k+str(i+1))) + ","
        s = s[:-1] + ":"
    return s[:-1]

def get_value(content, k):
    if k in content:
        return content[k]
    else:
        return ""

if __name__ == '__main__':
    pass
