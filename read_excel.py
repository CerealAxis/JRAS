# coding='utf-8'
import openpyxl

workbook = openpyxl.load_workbook('output/cleaned_data.xlsx')
target_column_number = 3
with open('output/contents.txt', 'w', encoding='utf-8') as f:
    for row in workbook.active.iter_rows():
        content = row[target_column_number - 1].value
        f.write(content + '\n')

f.close()

print('评论数据已总结')