#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Admin
#
# Created:     26/10/2015
#-------------------------------------------------------------------------------

def search_value(value, start, end):
    start_a, start_b, end_a, end_b = "","","",""
    start_x, start_y, end_x, end_y = 0,0,0,0
    count = 0

    for ch in start:
        if ord(ch) >= 65:
            start_a += ch
        else:
            start_b += ch
##
##    for ch in end:
##        if ord(ch) >= 65:   #ord('A') = 65
##            end_x += ord(ch) - 64
##        else:
##            end_y += int(ch)

##    return start_x,start_y,end_x,end_y
    return start_a, start_b
def main():
    pass

if __name__ == '__main__':
    print search_value(3,"AA1","Z10")
