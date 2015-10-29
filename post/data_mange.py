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
    return line[1].split(".")[1]

def get_channel(line):
    return line[2]

def get_rate(line):
    return line[3]

def get_bw(line):
    return line[4].split("-")[1]

def get_stream(line):
    return line[5]

def get_antenna(line):
    return line[6]

def get_power(line):
    return line[8]

get_func = {"standard":get_standard,
            "channel":get_channel,
            "rate":get_rate,
            "BW":get_bw,
            "stream":get_stream,
            "antenna":get_antenna,
            "power":get_power
           }
def load_data(path):
    data = {}
    with open(path, 'rb') as f:
        reader = csv.reader(f)
        reader.next()
        for line in reader:
            line = [i for i in line if i!='']
            if line:
                if len(line) > 2:  #line with config
                    data["standard"] = get_func["standard"](line)
                    data["channel"] = get_func["channel"](line)
                    data["rate"] = get_func["rate"](line)
                    data["BW"] = get_func["BW"](line)
                    data["stream"] = get_func["stream"](line)
                    data["antenna"] = get_func["antenna"](line)
                    data["power"] = get_func["power"](line)
                else:
                    data[line[0]] = line[1]
            else:
                yield data
                data = {}

if __name__ == '__main__':
##    pass
##    path = (r"D:\python task\draw_data_and_post\post\TX.csv")
    path = r"D:\game\abstract\draw_data_and_post\post\Log\5G_MIMO_New_S1-Tx\WAC7X0-S1-5G-2X2-MIMO-n-Tx-New_Result.csv"
    for i in load_data(path):
        print i