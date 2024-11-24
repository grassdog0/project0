from openpyxl import Workbook

title_list = ['theme','time','location','character','ddl']

str_list = [
    'Open Tennis Championships,20241109-20241110,西湖大学云谷校区一号和二号网球场,西湖大学全体在校学生和教职工,20241102',
    'Open Tennis Championships,20241109-20241110,西湖大学云谷校区一号和二号网球场,西湖大学全体在校学生和教职工,20241102',
    'Open Tennis Championships,20241109-20241110,西湖大学云谷校区一号和二号网球场,西湖大学全体在校学生和教职工,20241102',
    'Open Tennis Championships,20241109-20241110,西湖大学云谷校区一号和二号网球场,西湖大学全体在校学生和教职工,20241102']
str_l_list = []
for i in str_list:
    str_l_list.append(i.split(','))

t_i = 1  # t_i is the index of time
for j in range(len(str_l_list)):
    for i in range(len(str_l_list) - 1):
        if str_l_list[i][t_i] > str_l_list[i + 1][t_i]:
            str_l_list[i], str_l_list[i+1] = str_l_list[i+1], str_l_list[i]

str_l_list.insert(0, title_list)

wb = Workbook()
ws = wb.active

for row_index, row_data in enumerate(str_l_list):
    for col_index, cell_data in enumerate(row_data):
        ws.cell(row=row_index + 1, column=col_index + 1, value=cell_data)

wb.save('output.xlsx')