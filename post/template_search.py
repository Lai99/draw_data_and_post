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
    return False

def get_spec_pos(sheet,anchor,spec_x = 1,module_x = 2,case_x = 5):
    items = []
    start = 0
    for row in range(1,50):
        if Range(sheet,(row,spec_x)).value == anchor:
            start = row
            break
    last = (0,0)
    module_items = []
    for row in range(start+1,1000):
        if Range(sheet,(row,spec_x)).value != None and Range(sheet,(row,spec_x)).value != anchor:
##            if module_items:
            items.append(Range(sheet,(row,spec_x)).value)
##                items[Range(sheet,last).value] = module_items
##            module_items = []
##            last = (row,spec_x)
##        if Range(sheet,(row,module_x)).value != None:
##            module_items.append(Range(sheet,(row,module_x)).value)

##    while items.pop(anchor,False):
##        pass

    return items

if __name__ == '__main__':
    print search_value('E',"A1","K10")
