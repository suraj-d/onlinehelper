return_list = [['Tracking id', 'Status'], ['asdf', 'rto'], ['asdf', 'rto'],['asdf', 'rto'],['asdf', 'rto'],
               ['asdf', 'customer return']]
return_type = ['rto', 'customer return', 'wrong customer return', 'wrong courier return', 'cancel']
total_return = {'rto':0,
                'customer return': 0}
for list in return_list:
    if list[1] in return_type:
        a = total_return.get(list[1])
        total_return.update({list[1]: a + 1})


print(total_return)
