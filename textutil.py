import openpyxl
import argparse
import os
from datetime import datetime, timedelta


def print_log(msg):
    print("%s\t%s" % (datetime.now().strftime("%Y-%m-%d %H:%M"), msg))


def process_file():
    parser = argparse.ArgumentParser(description='处理SAP电源车数据.')
    parser.add_argument('--f', nargs='?', help='源文件名')
    parser.add_argument('--s', nargs='?', help='工作簿名称')
    args = parser.parse_args()
    if args.f is None:
        print_log("使用默认文件名：原版.xlsx")
        file_name = "原版.xlsx"
    else:
        file_name = args.f
    if not os.path.exists(file_name):
        print_log("%s不存在。" % file_name)
        return
    if args.s is None:
        month = datetime.now().month
        if month == 1:
            month = 12
        else:
            month = month - 1
        print_log("使用默认的工作簿名称（当前月份减一）：%d" % month)
        sheet_name = str(month)
    else:
        sheet_name = args.s
        month = sheet_name
    try:
        print_log("读取%s[%s]" % (file_name, sheet_name))
        xl = openpyxl.load_workbook(file_name)
        sheet = xl[sheet_name]
        print_log("最大行数：%s" % sheet.max_row)
        print_log("开始处理数据")
        hours_by_date_car = {}
        for row in range(sheet.max_row):
            if row == 0:
                continue
            date_info = sheet.cell(row=row+1, column=1).value
            car_no = sheet.cell(row=row+1, column=5).value
            hours = sheet.cell(row=row+1, column=6).value
            if car_no is None or hours is None:
                continue
            date_car = "%s,%s" % (date_info, car_no)
            if date_car in hours_by_date_car:
                hours_by_date_car[date_car] += hours
            else:
                hours_by_date_car[date_car] = hours
        print_log("保存结果")
        new_sheet = xl.create_sheet(title='%s-修改后' % str(month))
        row_index = 1
        for key, value in hours_by_date_car.items():
            date_car = key.split(",")
            new_sheet.cell(row=row_index, column=1).value = date_car[0]
            new_sheet.cell(row=row_index, column=2).value = date_car[1]
            new_sheet.cell(row=row_index, column=3).value = value
            row_index += 1
        xl.save(file_name)
    except Exception as e:
        print_log("发生异常：%s" % str(e))
        return
    return


if __name__ == "__main__":
    process_file()

