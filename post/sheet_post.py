#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Admin
#
# Created:     28/10/2015
#-------------------------------------------------------------------------------

item_ref = {"standard":"standard",
            "rate":"rate",
           }

def post(data_path):
    wb = Workbook.caller()

    fill_pos = template_search.get_fill_pos(2,"Standard")
##    print len(fill_pos)
    for data in data_mange.load_data(date_path):
        standard = find_standard(data, fill_pos)
        need_pos = find_rate(data, standard)

def find_standard(data,fill_pos):
    for k in fill_pos.keys():
        if data[item_ref["standard"]] in k:
            #
            return fill_pos[k]
    return False

def find_rate(data,fill_pos):
    for k in fill_pos.keys():
        if data[item_ref["channel"]] in k:
            #
            return fill_pos[k]
    return False


if __name__ == '__main__':
    main()
