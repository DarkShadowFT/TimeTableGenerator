import time
import math
from openpyxl import load_workbook
from xlwings import Range, Book
from Period import Period
import pandas as pd
import xlsxwriter


def extractSection(period):
	return str(period.section)

def get_time(start_time, minutes):
	res = ""
	temp = start_time.split(':')
	temp[1] = int(temp[1]) + int(minutes)
	hours = 0
	if temp[1] >= 60:
		hours = temp[1] / 60
		temp[1] %= 60
	elif temp[1] < 0:
		hours = math.floor(temp[1] / 60)
		temp[1] = 60 - (abs(temp[1]) % 60)
	temp[0] = int(temp[0]) + int(hours)
	if temp[0] < 10:
		temp[0] = "0" + str(temp[0])
	if temp[1] < 10:
		temp[1] = "0" + str(temp[1])
	res = str(temp[0]) + ":" + str(temp[1])
	return res

def make_timetable():
	day = ""
	venue = ""
	day_found = False
	none_count = 0
	none_length = 0
	start_time = 0
	start_time_set = False
	p_file = open("periods.txt", "w")

	for i in range(0, len(data)):
		j = 0
		while j < len(data[i]):
			if isinstance(data[i][j], str) and data[i][j].lower().find("period") != -1:
				j += 1
				time_count = 0
				times = []
				for k in range(j, len(data[i])):
					if time_count == 2:
						break
					elif data[i][k] != "None":
						times.insert(k, data[i][k])
						time_count += 1

				none_length = int(times[1]) - int(times[0])

			elif isinstance(data[i][j], str) and data[i][j].lower().find("a.m.") != -1 and not start_time_set:
				start_time_set = True
				start_time = data[i][j].split()[0]

			elif data[i][j] != "None":
				if j == 1 and day_found:
					none_count = 0
					venue = data[i][j]

				elif j > 1 and day_found:
					length = float(none_count) * float(none_length)
					period = data[i][j]
					if period != "NOT AVAILABLE":
						period = period.split('(')
						if len(period) > 1:
							section = period[1].split(')')
							cell_span = 1
							if Range((i + 1, j + 1)).merge_cells:
								cell_span = Range((i + 1, j + 1)).merge_area.count

							period_start_time = get_time(start_time, length)
							period_end_time = get_time(start_time, length + cell_span * none_length)
							p = Period(period_start_time, period_end_time, period[0], section[0], venue, day[0])
							none_count += cell_span
							j += cell_span - 1
							print(p, file = p_file)
				else:
					for k in range(0, len(weekdays)):
						if isinstance(data[i][j], str) and data[i][j].lower().find(weekdays[k]) != -1:
							none_count = 0
							day_found = True
							day = data[i][j].split()
							temp_day = day[0]
							while temp_day.lower() in weekdays:
								j += 1
								temp_day = data[i][j]

							venue = data[i][j]
							break
			else:
				none_count += 1
			j += 1
	
	p_file.close()


start_time = time.time()
path = r"C:\\Data\\SHIT-NUCES\\Semester 6\\Fast School of Computing Time Table Spring 2022 v1.2.xlsx"
wb = Book(path)

wb = wb.sheets[0]
wb_obj = load_workbook(path, True)
sheet_obj = wb_obj.active

print("Workbook loaded")
max_row = sheet_obj.max_row
max_col = sheet_obj.max_column
data = []

# Create file if not created
f1 = open("Output.txt", "a+")
f1.close()

f2 = open("Output.txt", "r")
contents = f2.read()
if len(contents) == 0:
	data = [ ["" for i in range(max_col)] for j in range(max_row) ]

	for i, row in enumerate(sheet_obj.rows):
		for j, cell in enumerate(row):
			data[i][j] = cell.value

	# Removing empty rows
	i = 0
	while i < len(data):
		for j in range(0, len(data[i])):
			if isinstance(data[i][j], str) and data[i][j].lower().find('=count') != -1:
				del data[i]
				break
		i += 1
		
	# Writing the parsed output to "Output.txt"
	with open("Output.txt", "w") as f:
		for i in range(0, len(data)):
			for j in range(0, len(data[i])):
				print(data[i][j], end = '', file = f)
				print('\t', end = '', file = f)
			print(file = f)

else:
	with open("Output.txt", "r") as w:
		data = w.read().split('\n')
		for i, words in enumerate(data):
			data[i] = words.split('\t')


	weekdays = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
	
	# Create periods.txt if not created
	p1 = open("periods.txt", "a")
	p1.close()

	periods = []
	p2 = open("periods.txt", "r")
	p2_contents = p2.read()
	if len(p2_contents) == 0:
		make_timetable()
	else:
		periods = p2_contents.split('\n')

	wb_obj.close()
	
	mon_periods, tue_periods, wed_periods, thu_periods, fri_periods, sat_periods = [], [], [], [], [], []
	sections = ["BCS-6A", "BCS-6B", "BCS-6C", "BCS-6D", "BCS-6E", "BCS-6F", "BCS-6G", "BCS-6H", "BCS-6J"]
	with open("periods.txt", "r") as w:
		f = w.read().split('\n')
		for period in f:
			if period != '':
				p = period.split('\t\t')
				name = p[0].split('(')[0].strip()
				section = p[0].split('(')[1].split(')')[0]
				startTime = p[1].split('-')[0]
				endTime = p[1].split('-')[1]
				venue = p[2]
				day = p[3]
				period = Period(startTime, endTime, name, section, venue, day)

				if period.day == "Monday" and period.section in sections:
					mon_periods.append(period)
				elif period.day == "Tuesday" and period.section in sections:
					tue_periods.append(period)
				elif period.day == "Wednesday" and period.section in sections:
					wed_periods.append(period)
				elif period.day == "Thursday" and period.section in sections:
					thu_periods.append(period)
				elif period.day == "Friday" and period.section in sections:
					fri_periods.append(period)
				elif period.day == "Saturday" and period.section in sections:
					sat_periods.append(period)

	all_day_periods = [mon_periods, tue_periods, wed_periods, thu_periods, fri_periods, sat_periods]

	section_periods = []
	for section in sections:
		section_periods.append([])

	for i, section in enumerate(sections):
		# print(f"\n\n\t\t{section}\n\n")
		for day in all_day_periods:
			# print(f"\n{day[0].day}\n")
			for period in day:
				if period.section == section:
					section_periods[i].append(period)

		# Create a Pandas Excel writer using XlsxWriter as the engine.
	writer = pd.ExcelWriter('timetable.xlsx', engine='xlsxwriter')

	row_numbers = []

	with open("tt.txt", "w") as tt:
		idx = 0
		for i, section in enumerate(section_periods):
			row_numbers.append(idx)

			section_df = pd.DataFrame(["", sections[i], ""]).transpose()
			section_df.to_excel(writer, sheet_name = 'Sheet1', startrow = idx, index = False, header = False)
			print(f"\t\t{sections[i]}\n", file = tt)

			df = pd.DataFrame([ [s.day, s.name, s.duration, s.venue] for s in section], columns = ["day", "name", "duration", "venue"])

			# Convert the dataframe to an XlsxWriter Excel object.
			df.to_excel(writer, sheet_name = 'Sheet1', startrow = idx + 1, index = False, header = False)

			idx += len(section) + 2
			print(f"{df.to_string(index = False)}\n", file = tt)

	# Close the Pandas Excel writer and output the Excel file.
	writer.save()


print("\n--- %s seconds ---" % (time.time() - start_time))
f2.close()