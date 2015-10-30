#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Admin
#
# Created:     26/10/2015
#-------------------------------------------------------------------------------
from xlwings import Workbook, Sheet, Range, Chart

def seq_to_int(string):
    pos = -1
    sum_str = 0
    for i in range(len(string)):
        sum_str += (ord(string[pos]) - 64) * (26**i)
        pos -= 1
    return sum_str

def split_range(string):
    mid = 0
    for ch in string:
        if ord(ch) >= 65:
            mid += 1
        else:
            break
    return string[:mid], string[mid:]

def range_to_cell(string):
    a, b = split_range(string)
    return seq_to_int(a), int(b)

def search_value(value, start, end):
    start_x, start_y, end_x, end_y = 0,0,0,0
    start_x, start_y = range_to_cell(start)
    end_x, end_y = range_to_cell(end)

    for row in range(1,(end_y - start_y)):
        for col in range(1,(end_x - start_x)):
            print Range((row,col)).value
            if Range((row,col)).value == value:
                return (row,col)
    return None

def manage_standard_5G(sheet,pos):
    s = Range(sheet,pos).value
    if "\n" in s:
        standard_rate, stream = s.split("\n")
        standard_rate = standard_rate.replace(" ","")
        standard, rate = standard_rate.split("-")
##        print standard_rate
        stream = stream.split(" ")[0]
        rate = rate.split("T")[-1]
##        print (standard,rate,stream)
        return (standard,rate,stream)
    else:
##        print s
        return s

def manage_standard_2G(sheet,pos):
    s = Range(sheet,pos).value
    if "\n" in s:
        standard_rate, stream = s.split("\n")
        standard_rate = standard_rate.strip()
        print standard_rate
        standard, rate = standard_rate.split(" ")
        stream = stream.split(" ")[0]
        rate = rate.split("M")[0]
##        print (standard,rate,stream)
        return (standard,rate,stream)
    else:
##        print s
        return s

def manage_modulation(sheet,pos):
    modulation = Range(sheet,pos).value.split("-")[0]
    return modulation

def mange_rate(sheet,pos):
    pass

def make_module_item_key(sheet, pos, offset):
    if Range(sheet,(pos[0],offset)).value:
        return manage_modulation(sheet,pos) + "-" + str(Range(sheet,(pos[0],offset)).value)
    else:
        return manage_modulation(sheet,pos)

standard_manage_func = {"2G":manage_standard_2G,
                        "5G":manage_standard_5G
                        }

def get_fill_pos(sheet,anchor,band,standard_x = 1,module_x = 2,rate_x = 3, case_x = 5, start_x = 6):
    """
    Get all value can be filled position in a sheet
    Input: int:specified sheet, string:anchor which used to split data block, int:spec column position, int: modulation column position,
           int:data rate column position, int:test items column position
    Output:dict:whole sheet value can be filled position, all anchors row loocation
    """
    start = 0
    all_anchor_row = []
    #Don't need sheet front content. Use anchor to go to standard start position
    for row in range(1,50):
        if Range(sheet,(row,standard_x)).value == anchor:
            start = row
            all_anchor_row.append(row)
            break
    last_standard = (0,0)
    last_module = (0,0)
    items = {}
    module_items = {}
    case_count = 0
    #Don't no when to end. Need a value as bound ex:1500
    for row in range(start+1,1500):
        #Use standard between standard to split data block, need to add last_standard one in the end
        if Range(sheet,(row,module_x)).value != None:    #Collect Modulations in a standard
            if case_count != 0:
                #Add "module and rate" with "value start position and case numbers"
                k = make_module_item_key(sheet, last_module, rate_x)
                module_items[k] = ((last_module[0], start_x),case_count)
##                Range(sheet,(last_module[0],7)).value = [str(((last_module[0], start_x),case_count)),Range(sheet,last_module).value + " " + str(Range(sheet,(last_module[0],rate_x)).value)]
                case_count = 0
            last_module = (row,module_x)

        if Range(sheet,(row,standard_x)).value != None:
            if  Range(sheet,(row,standard_x)).value != anchor:
                if module_items:   #if true means it has a modulation collection
                    items[standard_manage_func[band](sheet,last_standard)] = module_items
##                    Range(sheet,(last_standard[0],12)).value = module_items.keys()
##                    print module_items.values()
                    module_items = {}
                last_standard = (row,standard_x)  #A spec start position
            else:
                all_anchor_row.append(row)
                continue   #Not include row which has anchor

        if Range(sheet,(row,case_x)).value != None:  #Count how many test case
            case_count += 1

    #Don't forget last one have no end point
    if case_count != 0:
        #Add "module and rate" with "value start position and case numbers"
        k = make_module_item_key(sheet, last_module, rate_x)
        module_items[k] = ((last_module[0], start_x),case_count)
##        Range(sheet,(last_module[0],7)).value = [str(((last_module[0], start_x),case_count)),Range(sheet,last_module).value + " " + str(Range(sheet,(last_module[0],rate_x)).value)]

    if module_items:
        items[standard_manage_func[band](sheet,last_standard)] = module_items
##        Range(sheet,(last_standard[0],12)).value = module_items.keys()
##        print module_items.values()
    return items, all_anchor_row

def get_channel_start(sheet, pos, anchor, all_anchor_row = None):
    row, col = pos[0], pos[1]
    if all_anchor_row:
        last_row = 0
        for r in all_anchor_row:
            if r - row > 0:
                if last_row == 0:
                    break
                else:
                    return (last_row,col)
            else:
                last_row = r
    else:
        while row != 0:
            if anchor in str(Range(sheet,(row,col)).value):
                return (row,col)
            row -= 1
    print "Can't find channel form local in sheet"
    return None

def get_channel_pos(sheet, pos, ch):
    row, col = pos[0], pos[1]
    count = 16
    while count > 0:
        while Range(sheet,(row,col)).value:
            if ch in Range(sheet,(row,col)).value:
                return (row, col)
            col += 1
            count = 16
        col += 1
        count -= 1
    print "Can't find this channel in channel form"
    return None

if __name__ == '__main__':
    pass