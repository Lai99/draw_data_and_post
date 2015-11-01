#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Admin
#
# Created:     27/10/2015
#-------------------------------------------------------------------------------
import csv

def get_standard(line):
    if len(line) > 1:
        return line[1].split(".")[1]
    return None

def get_channel(line):
    if len(line) > 2:
        return line[2]
    return None

def get_rate(line):
    if len(line) > 3:
        return line[3]
    return None

def get_bw(line):
    if len(line) > 4:
        return line[4].split("-")[1]
    return None

def get_stream(line):
    if len(line) > 5:
        return line[5]
    return None

def get_antenna(line):
    if len(line) > 6:
        return line[6]
    return None

def get_power(line):
    if len(line) > 8:
        if line[7] == "Power":
            return line[8]
    return None

def get_sens(line):
    if len(line) > 8:
        if line[7] == "SENS":
            return line[8]
    return None

get_func = {"standard":get_standard,
            "channel":get_channel,
            "rate":get_rate,
            "BW":get_bw,
            "stream":get_stream,
            "antenna":get_antenna,
            "power":get_power,
            "sens":get_sens
           }
def load_data(path):
    data = {}
    with open(path, 'rb') as f:
        reader = csv.reader(f)
        reader.next()
        for line in reader:
            line = [i for i in line if i!='']
            if line:
##                print line
                if len(line) > 2:  #line with config
                    if get_func["standard"](line):
                        data["standard"] = get_func["standard"](line)
                    if get_func["channel"](line):
                        data["channel"] = get_func["channel"](line)
                    if get_func["rate"](line):
                        data["rate"] = get_func["rate"](line)
                    if get_func["BW"](line):
                        data["BW"] = get_func["BW"](line)
                    if get_func["stream"](line):
                        data["stream"] = get_func["stream"](line)
                    if get_func["antenna"](line):
                        data["antenna"] = get_func["antenna"](line)
                    # TX power
                    if get_func["power"](line):
                        data["power"] = get_func["power"](line)

                    # RX SENS
                    if get_func["sens"](line):
                        data["sens"] = get_func["sens"](line)

                else:
                    if len(line) > 1:
                        data[line[0]] = line[1]
                    else:
                        data[line[0]] = None
            else:
                yield data
                data = {}

if __name__ == '__main__':
##    pass
    t = r"D:\game\abstract\WAC740"
    path = (t + r"\IQFact_5G_Rx_SIMO_result.csv")
##    path = r"D:\game\abstract\draw_data_and_post\post\Log\5G_MIMO_New_S1-Tx\WAC7X0-S1-5G-2X2-MIMO-n-Tx-New_Result.csv"
    for i in load_data(path):
        print i