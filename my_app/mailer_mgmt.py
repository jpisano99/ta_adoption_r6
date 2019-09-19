import xlrd
from my_app.settings import app_cfg
from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.get_list_from_ss import get_list_from_ss
from my_app.func_lib.push_xls_to_ss import push_xls_to_ss

# push_xls_to_ss('mailer names.xlsx', 'CTUG mailer')
# jim = get_list_from_ss('Tetration Coverage Map')
# print (jim)
#
# exit()


# Feed this a list like this
# Chris Mchenry (chmchenr) <chmchenr@cisco.com>; Gordon Hirst (ghirst) <ghirst@cisco.com>;
#
wb, ws = open_wb('mailer scrub.xlsx')

raw = ws.cell_value(0, 0)
raw_len = len(raw)
names = []
name = ''
for c in raw:
    if c != ';':
        name = name + c
    else:
        if name[0] == ' ':
            name = name[1:]
        names.append(name)
        name = ''


word = ''
word_list = []
scrubbed_names = [['fname', 'lname', 'full name', 'username', 'email']]
for name in names:
    for c in name:
        if c == ' ' or c == '>':
            word = word.replace('(', '')
            word = word.replace(')', '')
            word = word.replace('<', '')
            word_list.append(word)
            word = ''
        else:
            word = word + c

    # In case we have a middle name consolidate
    if len(word_list) != 4:
        word_list = [word_list[0], word_list[1] + ' ' + word_list[2], word_list[3], word_list[4]]

    # create and insert the full name ie Pisano, Jim
    fname = word_list[0]
    lname = word_list[1]
    full_name = lname+', '+ fname
    word_list.insert(2, full_name)

    scrubbed_names.append(word_list)
    word_list = []

for user in scrubbed_names:
    print(user)




push_list_to_xls(scrubbed_names,'mailer_names.xlsx')
# push_xls_to_ss('mailer_names.xlsx', 'CTUG mailer')


