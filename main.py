import time
import math
from openpyxl import load_workbook
from xlwings import Range, Book
from Period import Period


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
	periods = []
	f = open("periods.txt", "w")

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
							periods.append(p)
							print(p, file = f)
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
	
	f.close()
	return periods


start_time = time.time()
path = r"C:\\Data\\SHIT-NUCES\\Semester 6\\Fast School of Computing Time Table Spring 2022 v1.2.xlsx"
wb = Book(path)

wb = wb.sheets[0]
wb_obj = load_workbook(path, True)
sheet_obj = wb_obj.active

print("Workbook loaded")
max_row = sheet_obj.max_row
max_col = sheet_obj.max_column

# Create file if not created
f = open("Output.txt", "a+")
f.close()

f = open("Output.txt", "r")
data = []
if len(f.read()) == 0:
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
	periods = make_timetable()

	# mon_periods, tue_periods, wed_periods, thu_periods, fri_periods, sat_periods = [], [], [], [], [], []
	# sections = ["BCS-6A", "BCS-6B", "BCS-6C", "BCS-6D", "BCS-6E", "BCS-6F", "BCS-6G", "BCS-6H", "BCS-6J"]
	# with open("periods.txt", "r") as w:
	# 	f = w.read().split('\n')
	# 	for period in f:
	# 		if period != '':
	# 			p = period.split(',')
	# 			for i in range(len(p)):
	# 				p[i] = p[i].split('=')
	# 			name = p[0][1].strip()
	# 			section = p[1][1].strip()
	# 			startTime = p[2][1].strip()
	# 			endTime = p[3][1].strip()
	# 			venue = p[4][1].strip()
	# 			day = p[5][1].strip()
	# 			period = Period(startTime, endTime, name, section, venue, day)

	# 			if period.day == "Monday" and period.section in sections:
	# 				mon_periods.append(period)
	# 			elif period.day == "Tuesday" and period.section in sections:
	# 				tue_periods.append(period)
	# 			elif period.day == "Wednesday" and period.section in sections:
	# 				wed_periods.append(period)
	# 			elif period.day == "Thursday" and period.section in sections:
	# 				thu_periods.append(period)
	# 			elif period.day == "Friday" and period.section in sections:
	# 				fri_periods.append(period)
	# 			elif period.day == "Saturday" and period.section in sections:
	# 				sat_periods.append(period)

	# # for section in sections:
	# # 	print(section, end = '\t\t')
	# # print()

	# # mon_periods = sorted(mon_periods, key=extractSection)
	# # sheet_obj.append(mon_periods)
	# # for m_period in mon_periods:
	# 	# print(m_period)
	# # wb_obj.close()

	# all_day_periods = [mon_periods, tue_periods, wed_periods, thu_periods, fri_periods, sat_periods]

	# for section in sections:
	# 	print(f"\n\n\t\t{section}\n\n")
	# 	for day in all_day_periods:
	# 		print(f"\n{day[0].day}\n")
	# 		for period in day:
	# 			if period.section == section:
	# 				print(period)


print("\n--- %s seconds ---" % (time.time() - start_time))
