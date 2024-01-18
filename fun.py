import os
import xlwt
import xlrd
from xlutils.copy import copy
import decimal
import math
import time

path = "C:\\Users\\skr\\Desktop\\需求包\\txt数据库"
save_name = time.strftime("%Y%m%d_%H%M%S", time.localtime()) + ".xlsx"

files = os.listdir(path)
wb = xlwt.Workbook(encoding='utf-8')
save_arr = []

for file in files:
    sheet_name = file.split('.')[0]
    ws = wb.add_sheet(sheet_name)
    i = 0
    with open(os.path.join(path, file), "r") as f:
        started = False
        for line in f:
            if line.strip() == "END_DATA":
                break
            if line.strip() == "BEGIN_DATA":
                started = True
                continue
            if started:
                arr = line.strip().split('\t')
                for j in range(len(arr)):
                    ws.write(i, j, arr[j].strip())
                i += 1

wb.save(save_name)

# 读取excel
books = xlrd.open_workbook(save_name)
sheet = books.sheet_by_index(0)
sheet_no = len(books.sheets())

all_data = []
for j in range(sheet.nrows):
    temp = []
    for i in range(sheet_no):
        sheet = books.sheet_by_name(books[i].name)
        temp.append(sheet.row_values(j, 12))
    all_data.append(temp)

wb = copy(books)
new_worksheet = wb.add_sheet('数据分析结果')
new_worksheet.write_merge(0, 0, 0, 2, '平均值')
new_worksheet.write(1, 0, 'L')
new_worksheet.write(1, 1, 'A')
new_worksheet.write(1, 2, 'B')

for i in range(sheet_no):
    new_worksheet.write(0, i + 3, books[i].name)
    new_worksheet.write(1, i + 3, '色差')

ava_arr = []
for i in range(len(all_data)):
    sum1 = 0
    sum2 = 0
    sum3 = 0
    for j in range(len(all_data[i])):
        # 计算平均值
        sum1 = sum1 + decimal.Decimal(all_data[i][j][0])
        sum2 = sum2 + decimal.Decimal(all_data[i][j][1])
        sum3 = sum3 + decimal.Decimal(all_data[i][j][2])

    round1 = round(sum1/len(all_data[i]), 2)
    round2 = round(sum2/len(all_data[i]), 2)
    round3 = round(sum3/len(all_data[i]), 2)

    ava_arr.append([round1, round2, round3])

    # 写入平均值
    new_worksheet.write(i + 2, 0, round1)
    new_worksheet.write(i + 2, 1, round2)
    new_worksheet.write(i + 2, 2, round3)

cal_arr = []
# 计算开根号
for i in range(len(ava_arr)):
    temp = []
    for j in range(len(all_data[i])):
        # 列
        M = math.pow(abs(decimal.Decimal(all_data[i][j][0]) - decimal.Decimal(ava_arr[i][0])), 2)
        N = math.pow(abs(decimal.Decimal(all_data[i][j][1]) - decimal.Decimal(ava_arr[i][1])), 2)
        O = math.pow(abs(decimal.Decimal(all_data[i][j][2]) - decimal.Decimal(ava_arr[i][2])), 2)
        temp.append(round(math.sqrt(M + N + O), 2))
    cal_arr.append(temp)

for i in range(len(cal_arr)):
    for j in range(sheet_no):
        new_worksheet.write(i + 2, 3 + j, cal_arr[i][j])

wb.save(save_name)