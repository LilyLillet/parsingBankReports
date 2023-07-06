import report_reader
import report_writer

file_path = 'F:\\sber 1.xlsx'

l = report_reader.get_sber_operations_list(file_path)
for i in l:
    print(i)
