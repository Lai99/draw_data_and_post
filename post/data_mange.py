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

get_func = {"standard":get_standard,
            "channel":get_channel,
            "rate":get_rate,
            "BW":get_bw,
            "stream":get_stream,
            "antenna":get_antenna
           }

def load_data(path):
    data = {}
    with open(path, 'rb') as f:
        reader = csv.reader(f)
        reader.next()
##        line = [i for i in reader.next() if i!='']
##        while line:
##            line = [i for i in reader.next() if i!='']
##            print line

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
                else:
                    pass
            else:
                yield data
                data = {}


def test():
    y = 1
    for i in range(3):
        yield i, y
        y += 1


if __name__ == '__main__':
##    load_data(r"D:\python task\draw_data_and_post\post\TX.csv")
    for i in load_data(r"D:\python task\draw_data_and_post\post\TX.csv"):
        print i