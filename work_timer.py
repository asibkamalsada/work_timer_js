import os
from calendar import month_name, different_locale
import datetime

import csv
import sys

import openpyxl

dates = {}

with open(sys.argv[2], "r", encoding="utf-8") as my_csv:
    spam_reader = csv.DictReader([x.strip() for x in my_csv.readlines()[:-3]])
    current_month = None
    for row in spam_reader:
        start: str = row["Von"]
        end: str = row["Bis"]
        if datetime.date.fromisoformat(start.split()[0]) == datetime.date.fromisoformat(end.split()[0]):
            current_date = datetime.date.fromisoformat(start.split()[0])
            if current_month:
                if current_month != current_date.month:
                    raise Exception("not the same months")
            else:
                current_month = current_date.month
        else:
            raise Exception("working over the end of a day")
        start_time = datetime.datetime.fromisoformat(start)
        end_time = datetime.datetime.fromisoformat(end)
        if not dates.get(current_date.day, None):
            dates[current_date.day] = list()
        dates[current_date.day].append((start_time, end_time))

with different_locale("de_DE"):
    current_month_name = month_name[current_month]

workbook = openpyxl.load_workbook(sys.argv[1])
sheet = workbook.active

sheet["D4"] = current_month_name

for day, times in dates.items():
    times = sorted(times, key=lambda x: x[0])
    start: datetime.datetime = times[0][0]
    end: datetime.datetime = times[-1][-1]

    sheet[f"C{6 + day}"] = start.time()#.isoformat(timespec="seconds")
    sheet[f"C{6 + day}"].number_format = "h:mm"
    sheet[f"D{6 + day}"] = end.time()#.isoformat(timespec="seconds")
    sheet[f"D{6 + day}"].number_format = "h:mm"
    
    if len(times) > 1:
        pause = datetime.timedelta()
        for i in range(len(times) - 1):
            pause += times[i + 1][0] - times[i][1]
        sheet[f"E{6 + day}"] = pause
        sheet[f"E{6 + day}"].number_format = "h:mm"

result_path = os.path.join(os.path.dirname(sys.argv[1]), "xlsxs", f"Arbeitszeiten_Asib_Kamalsada_{current_month_name}.xlsx")

workbook.save(result_path)

print(f"csv read from {os.path.abspath(sys.argv[2])}")
print(f"template taken from {os.path.abspath(sys.argv[1])}")
print(f"file saved to {result_path}")
