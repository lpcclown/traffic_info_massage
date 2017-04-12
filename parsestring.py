import csv
import time
import datetime
import openpyxl

starting_time = "02:10"
duration = 68


def calculate_ending_time(start_time):
    adding_hour = int(duration / 60)
    adding_minute = int(duration % 60)
    ending_time_hour = int(str(start_time).split(':')[0]) + adding_hour
    if ending_time_hour < 10:
        ending_time_hour = "0" + str(ending_time_hour)
    ending_time_minute = int(str(start_time).split(':')[1]) + adding_minute
    if ending_time_minute < 10:
        ending_time_minute = "0" + str(ending_time_minute)
    return str(ending_time_hour) + ":" + str(ending_time_minute)


def transfrom_time(original_string):
    old_hour_minute = original_string.split(' ', 5)[4].rsplit(':', 1)[0]
    old_hour = old_hour_minute.split(':', 1)[0]
    old_date = original_string.split(' ', 5)[3]+original_string.split(' ', 5)[1]+original_string.split(' ', 5)[2]
    date = original_string[0:24]

    new_minute = int(original_string.split(' ', 5)[4].split(':', 2)[1]) + 4
    if new_minute < 10:
        new_minute = "0" + str(new_minute)
    return str(old_hour_minute) + " - " + str(old_hour) + ":" + str(new_minute)


def transfrom_date(original_string):
    old_hour_minute = original_string.split(' ', 5)[4].rsplit(':', 1)[0]
    old_hour = old_hour_minute.split(':', 1)[0]
    old_date = original_string.split(' ', 5)[3]+original_string.split(' ', 5)[1]+original_string.split(' ', 5)[2]
    date = original_string[0:24]
    return str(datetime.datetime.strptime(date, "%a %b %d %Y %H:%M:%S")).split(' ')[0]


def compare_time(input_time, interval_upper_time):
    input_time_hour = str(input_time).split(':', 0)
    input_time_minute = str(input_time).split(':', 1)
    interval_upper_time_hour = str(interval_upper_time).split(':', 0)
    interval_upper_time_minute = str(interval_upper_time).split(':', 1)
    if input_time_hour > interval_upper_time_hour:
        return 0  # no need to write
        if input_time_minute > interval_upper_time_minute:
            return 0  # no need to write
    else:
        return 1


xfile = openpyxl.load_workbook('text2.xlsx')
sheet = xfile.get_sheet_by_name('Sheet1')
reader = csv.DictReader(open("4869_ET.csv"))
result = {}
a = ""
b = ""
c = 1
d = 1

# print(compare_time("04:10", "03:18"))

for row in reader:
    for column, value in sorted(row.iteritems(), reverse= True):
        if column == "Time\"":
            if (compare_time(starting_time, transfrom_time(str(value)).split(" - ")[1]) == 1 and compare_time(transfrom_time(str(value)).split(" - ")[0], calculate_ending_time(starting_time)) == 1):
                c += 1
                cell_name = "C" + str(c)
                sheet[cell_name] = transfrom_time(str(value))
                cell_previous_name = "A" + str(int(c)-1)
                cell_name = "A" + str(c)
                # print(transfrom_date(str(value)) + " 888888 " + str(sheet[cell_previous_name].value))
                if len(str(sheet[cell_previous_name].value)) > 5:  # used to compare None
                    temp = str(sheet[cell_previous_name].value)
                if transfrom_date(str(value)) != temp:
                    sheet[cell_name] = transfrom_date(str(value))
            else:
                break
        if column == "Strength":
            cell_name = "D" + str(c)
            sheet[cell_name] = str(value)

xfile.save('text3.xlsx')