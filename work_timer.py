import os
from calendar import month_name, different_locale, monthrange
import datetime

import csv
import sys

import platform

import pywintypes
import win32com.client

import openpyxl
date_to_comment = dict()
date_to_times = {}

with open(sys.argv[2], "r", encoding="utf-8") as my_csv:
    spam_reader = csv.DictReader([x.strip() for x in my_csv.readlines()[:-3]])
    current_month = None
    current_year = None
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
                current_year = current_date.year
        else:
            raise Exception("working over the end of a day")
        start_time = datetime.datetime.fromisoformat(start)
        end_time = datetime.datetime.fromisoformat(end)
        if not date_to_times.get(current_date.day, None):
            date_to_times[current_date.day] = list()
        date_to_times[current_date.day].append((start_time, end_time))
        if row.get("Kommentar", "") != "":
            date_to_comment[current_date.day] = row["Kommentar"]
        del start_time, end_time, start, end
    del spam_reader

with different_locale("de_DE"):
    current_month_name = month_name[current_month]

workbook = openpyxl.load_workbook(sys.argv[1])
sheet = workbook.active

sheet["D4"] = current_month_name

for month_day in range(1, monthrange(current_year, current_month)[1] + 1):
    sheet[f"B{6 + month_day}"] = f"{month_day}."

for day, times in date_to_times.items():
    times.sort(key=lambda x: x[0])
    start: datetime.datetime = times[0][0]
    end: datetime.datetime = times[-1][-1]

    sheet[f"C{6 + day}"] = start.time()  # .isoformat(timespec="seconds")
    sheet[f"C{6 + day}"].number_format = "h:mm"
    sheet[f"D{6 + day}"] = end.time()  # .isoformat(timespec="seconds")
    sheet[f"D{6 + day}"].number_format = "h:mm"

    if len(times) > 1:
        pause = datetime.timedelta()
        for i in range(len(times) - 1):
            pause += times[i + 1][0] - times[i][1]
        sheet[f"E{6 + day}"] = pause
        sheet[f"E{6 + day}"].number_format = "h:mm"

    if date_to_comment.get(day):
        sheet[f"i{6 + day}"] = date_to_comment[day]

file_name_without_extension = f"Arbeitszeiten_Asib_Kamalsada_{current_year}_{current_month:02d}_{current_month_name}"

result_path = os.path.abspath(os.path.join(os.path.dirname(sys.argv[1]), "xlsxs", f"{file_name_without_extension}.xlsx"))

if os.path.exists(result_path):
    if input(f"{result_path} already exists. overwrite? [Y/n]") not in ["y", "Y", ""]:
        print("aborted.")
        sys.exit(0)
    os.remove(result_path)

workbook.save(result_path)

print(f"csv read from {os.path.abspath(sys.argv[2])}")
print(f"template taken from {os.path.abspath(sys.argv[1])}")
print(f"file saved to {result_path}")

workbook.close()

o = win32com.client.Dispatch("Excel.Application")

o.Visible = False

path_to_pdf = os.path.abspath(os.path.join(os.path.dirname(sys.argv[1]), "pdfs", f"{file_name_without_extension}.pdf"))
if os.path.exists(path_to_pdf):
    os.remove(path_to_pdf)

print_area = 'A1:I37'

try:
    wb = o.Workbooks.Open(result_path)
    ws = wb.Worksheets[0]
    ws.PageSetup.PaperSize = 9  # A4 laut https://docs.microsoft.com/de-de/office/vba/api/excel.xlpapersize
    ws.ExportAsFixedFormat(0, path_to_pdf)
    print(f"printed as pdf to {path_to_pdf}")
except pywintypes.com_error as err:
    print(err)
finally:
    wb.Close(False)
