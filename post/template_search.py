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

def manage_standard(sheet,pos):
    s = Range(sheet,pos).value
    if "\n" in s:
        standard_rate, stream = s.split("\n")
        standard_rate = standard_rate.replace(" ","")
        standard, rate = standard_rate.split("-")
        stream = stream.split(" ")[0]
        rate = rate.split("T")[-1]
##        print (standard,rate,stream)
        return (standard,rate,stream)
    else:
##        print s
        return (s,None,None)

def manage_modulation(sheet,pos):
    s = Range(sheet,pos).value
    Range(sheet,last_module).value + " " + str(Range(sheet,(last_module[0],rate_x)).value)

def get_fill_pos(sheet,anchor,standard_x = 1,module_x = 2,rate_x = 3, case_x = 5, start_x = 6):
    """
    Get all value can be filled position in a sheet
    Input: int:specified sheet, string:anchor which used to split data block, int:spec column position, int: modulation column position,
           int:data rate column position, int:test items column position
    Output:dict:whole sheet value can be filled position
    """
    start = 0
    #Don't need sheet front content. Use anchor to go to standard start position
    for row in range(1,50):
        if Range(sheet,(row,standard_x)).value == anchor:
            start = row
            break
    last_standard = (0,0)
    last_module = (0,0)
    items = {}
    module_items = {}
    case_count = 0
    #Don't no when to end. Need a value as bound ex:1000
    for row in range(start+1,1000):
        #Use standard between standard to split data block, need to add last_standard one in the end
        if Range(sheet,(row,module_x)).value != None:    #Collect Modulations in a standard
            if case_count != 0:
                #Add "module and rate" with "value start position and case numbers"
                module_items[manage_modulation(sheet,last_module)] = ((last_module[0], start_x),case_count)
##                Range(sheet,(last_module[0],7)).value = [str(((last_module[0], start_x),case_count)),Range(sheet,last_module).value + " " + str(Range(sheet,(last_module[0],rate_x)).value)]
                case_count = 0
            last_module = (row,module_x)

        if Range(sheet,(row,standard_x)).value != None:
            if  Range(sheet,(row,standard_x)).value != anchor:
                if module_items:   #if true means it has a modulation collection
                    items[manage_standard(sheet,last_standard)] = module_items
##                    Range(sheet,(last_standard[0],12)).value = module_items.keys()
##                    print module_items.values()
                    module_items = {}
                last_standard = (row,standard_x)  #A spec start position
            else:
                continue   #Don't need row which has anchor

        if Range(sheet,(row,case_x)).value != None:  #Count how many test case
            case_count += 1

    #Don't forget last one have no end point
    if case_count != 0:
        #Add "module and rate" with "value start position and case numbers"
        module_items[manage_modulation(sheet,last_module)] = ((last_module[0], start_x),case_count)
##        Range(sheet,(last_module[0],7)).value = [str(((last_module[0], start_x),case_count)),Range(sheet,last_module).value + " " + str(Range(sheet,(last_module[0],rate_x)).value)]

    if module_items:
        items[manage_standard(sheet,last_standard)] = module_items
##        Range(sheet,(last_standard[0],12)).value = module_items.keys()
##        print module_items.values()
    return items

def get_channel_start(sheet, pos, anchor):
    row, col = pos[0], pos[1]
    while row != 0:
        if anchor in str(Range(sheet,(row,col)).value):
            return (row,col)
        else:
            row -= 1
    return None

def get_channel_pos(sheet, pos, ch):
    row, col = pos[0], pos[1]
    while Range(sheet,(row,col)).value:
        if ch in Range(sheet,(row,col)).value:
            return (row, col)
        col += 1
    return None

if __name__ == '__main__':
    pass