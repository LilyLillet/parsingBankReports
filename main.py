import tinkoff

lst = tinkoff.get_fin_operations_list('F:\\1kv.xlsx')
tinkoff.write_to_excel(lst)