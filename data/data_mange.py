#-------------------------------------------------------------------------------
# Name:        data_mange
# Purpose:     Turn raw data to format data dict{item:value}
#
# Author:      Lai
#
# Created:     19/10/2015
#-------------------------------------------------------------------------------
import draw_item

def pass_result(content):
    """
    If it is a PASS result return True else return False
    """
    for line in content:
        if "[Failed]" in line:
            return False
    return True

def freq_change(content1,content2):
    """
    freq. changed return True
    """
    return content1[0].split(" ")[1] != content2[0].split(" ")[1]

def strip_data(s):
    """
    Capture in a format: Item, Value
    """
    strip_str = []
    temp = s.split(":")
    return temp[0].strip(),temp[1].strip().split(" ")[0]

def draw_items_value(content):
    """
    Capture needed item/value pair, list first one is the config
    """
    draw_data = {}
    ####
    #Not want to get the last running time data. That makes fault
    if ":" in content[0]:
        return draw_data
    ####
    draw_data["CONFIG"] = content[0]
    for line in content:
        if ":" in line:
            item_and_value = strip_data(line)
            draw_data[item_and_value[0]] = item_and_value[1]
    return draw_data

def make_modifed_data(path,anchor):
    """
    Draw needed items value from raw file.
    Input: string:file_path, string:data start position
    Output: list:[string:drew_items_value]
    """
    f = open(path,"r")
    data = []
    temp_data1 = []
    temp_data2 = []
    record_start = False

    for line in f:
        if anchor in line:
            if not record_start:
                record_start = True
            else:
                if temp_data2:
                    if freq_change(temp_data2,temp_data1):
                        if draw_items_value(temp_data2):
                            data.append(draw_items_value(temp_data2))
                        temp_data2 = []
                    else:
                        if pass_result(temp_data1):
                            temp_data2 = list(temp_data1)
                else:
                    if pass_result(temp_data1):
                        temp_data2 = list(temp_data1)
                temp_data1 = []

        if record_start:
            temp_data1.append(line)

    f.close()
    return data

if __name__ == '__main__':
    pass