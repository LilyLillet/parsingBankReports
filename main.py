import report_reader
import report_writer

file_path = 'F:\\alfa 4.xlsx'
name = input()

l = report_reader.get_alfa_operations_list(file_path)
report_writer.write_to_excel(l, name)

