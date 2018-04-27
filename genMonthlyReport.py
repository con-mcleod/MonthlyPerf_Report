#!/usr/bin/python3

import sys, sqlite3, os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, Font, PatternFill, Border, Side
from datetime import datetime

if (len(sys.argv) != 1):
	print ("Usage: python3 genMonthlyReport.py")
	exit(1)

DATABASE = "dataset.db"
output_day = str(datetime.now().day)
output_month = str(datetime.now().month)
output_year = str(datetime.now().year)
output = output_year + "." + output_month + "." + output_day + ".xlsx"

connection = sqlite3.connect(DATABASE)
cursor = connection.cursor()

##############################
#                            #
# Database handling          #
#                            #
##############################

# prepare sql query execution
def dbselect(query, payload):
	connection = sqlite3.connect(DATABASE)
	cursorObj = connection.cursor()
	if not payload:
		rows = cursorObj.execute(query)
	else:
		rows = cursorObj.execute(query,payload)
	results = []
	for row in rows:
		results.append(row)
	cursorObj.close()
	return results

# execute sql query
def dbexecute(query, payload):
	connection = sqlite3.connect(DATABASE)
	cursor = connection.cursor()
	if not payload:
		cursor.execute(query)
	else:
		cursor.execute(query, payload)
	connection.commit()
	connection.close()

##############################
#                            #
# Helper functions           #
#                            #
##############################

# return all SMIs from SF
def get_all_SMIs():
	query = "SELECT distinct(SMI) from MONTH_GEN"
	payload = None
	all_SMIs = dbselect(query, payload)
	return all_SMIs

# return the range of dates from encompass report
def get_all_months():
	query = """SELECT obs_month, obs_year from DAILY_GEN 
			group by obs_month, obs_year order by obs_year, obs_month"""
	payload = None
	all_dates = dbselect(query, payload)
	return all_dates

# return SMI details from SF
def get_SMI_details(SMI):
	query = """SELECT ref_no, state, installer, PVsize, export_control,
		panel_brand, site_status, supply_date from SMI_DETAILS where SMI=?"""
	payload = (SMI)
	result = dbselect(query, payload)
	return result

# return SMI's monthly forecasts
def get_SMI_forecast(SMI):
	query = "SELECT val from FORECAST where SMI=?"
	payload = (SMI)
	result = dbselect(query, payload)
	return result

# return SMI's actual generation
def get_SMI_generation(SMI):
	query = "SELECT val from MONTH_GEN where SMI=? order by year, month"
	payload = (SMI)
	result = dbselect(query, payload)
	if not result:
		for i in range(0,len(dates)):
			result.append((0,))
	return result

# return SMI's annual forecast, generation and percentage
def get_perf(SMI, period):
	fc = 0
	gen = 0
	query = "SELECT adj_val from ADJ_FORECAST where SMI=? order by year, month"
	payload = (SMI)
	fc_vals = dbselect(query, payload)
	if (period == "Annual"):
		fc_vals = fc_vals[-12:]
	elif (period == "Quarter"):
		fc_vals = fc_vals[-3:]
	elif (period == "Month"):
		fc_vals = fc_vals[-1:]

	if fc_vals:
		if (period == "Prev"):
			fc_vals = fc_vals[-2]

	query = "SELECT val from MONTH_GEN where SMI=? order by year, month"
	payload = (SMI)
	gen_vals = dbselect(query, payload)[-12:]
	if (period == "Annual"):
		gen_vals = gen_vals[-12:]
	elif (period == "Quarter"):
		gen_vals = gen_vals[-3:]
	elif (period == "Month"):
		gen_vals = gen_vals[-1:]
	
	if gen_vals:
		if (period == "Prev"):
			gen_vals = gen_vals[-2]

	if period == "Prev":
		for fc_val in fc_vals:
			fc += fc_val
		for gen_val in gen_vals:
			gen += gen_val
		if (fc != 0):
			perf = gen/fc
		else:
			perf = 0
	else:
		for fc_val in fc_vals:
			fc += fc_val[0]
		for gen_val in gen_vals:
			gen += gen_val[0]
		if (fc != 0):
			perf = gen/fc
		else:
			perf = 0
	return fc, gen, perf

# return number of off days for month
def get_off_days(SMI, dates):
	curr_date = dates[-1]
	curr_month = curr_date[0]
	curr_year = curr_date[1]
	off_days = 0
	query = """SELECT value from DAILY_GEN where SMI=? and obs_month=? 
				and obs_year=?"""
	payload = (SMI[0], curr_month, curr_year)
	result = dbselect(query, payload)
	if result:
		for val in result:
			if not val[0]:
				off_days += 1
			elif val[0] < .1:
				off_days += 1
	return off_days

##############################
#                            #
# GENERATE OUTPUT            #
#                            #
##############################

wb = load_workbook(output)
if 'Perf Report' in wb.sheetnames:
	ws = wb.get_sheet_by_name('Perf Report')
else:
	ws = wb.create_sheet('Perf Report')

redFill = PatternFill(start_color='FA5858', end_color='FA5858', fill_type='solid')
greenFill = PatternFill(start_color='9Afe2e', end_color='9Afe2e', fill_type='solid')
leftBorder = Border(left=Side(style='thin'))

SMIs = get_all_SMIs()
dates = get_all_months()

ws_headings = ["SMI","Ref No","State","Installer","System Size","Export Control",
				"Panel Make","PPA Status","Supply Date","Jan FC","Feb FC", "Mar FC",
				"Apr FC","May FC","Jun FC","Jul FC","Aug FC","Sep FC","Oct FC",
				"Nov FC","Dec FC"]
for date in dates:
	date = "gen(" + str(date).strip('()') + ")"
	ws_headings.append(date)
ws_headings.extend(["Annaul FC","Annual Gen","Annual Perf","Quarter FC","Quarter Gen",
					"Quater Perf","Month FC","Month Gen","Month Perf","Prev FC",
					"Prev Gen","Prev Perf","Outage Days"])
row_count = 1
for SMI in SMIs:
	col_count = 0
	if row_count == 1:
		for heading in ws_headings:
			ws.cell(row=row_count, column=col_count+1).value = heading
			col_count += 1
	col_count = 1
	ws.cell(row=row_count+1, column=1).value = SMI[0]
	details = get_SMI_details(SMI)
	if details:
		for detail in details[0]:
			ws.cell(row=row_count+1, column=col_count+1).value = detail
			col_count += 1
	forecasts = get_SMI_forecast(SMI)
	ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
	for forecast in forecasts:
		ws.cell(row=row_count+1, column=col_count+1).value = forecast[0]
		col_count += 1
	month_gen = get_SMI_generation(SMI)
	ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
	for gen in month_gen:
		ws.cell(row=row_count+1, column=col_count+1).value = gen[0]
		col_count += 1
	i = 0
	annual_perf = get_perf(SMI, "Annual")
	for val in annual_perf:
		ws.cell(row=row_count+1, column=col_count+1).value = val
		if i == 0:
			ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
		if i == 2:
			ws.cell(row=row_count+1, column=col_count+1).number_format = '0.00%'
			if val < .9:
				ws.cell(row=row_count+1, column=col_count+1).fill = redFill
			elif val > 1.2:
				ws.cell(row=row_count+1, column=col_count+1).fill = greenFill
		col_count += 1
		i += 1
	i = 0
	quarter_perf = get_perf(SMI, "Quarter")
	for val in quarter_perf:
		ws.cell(row=row_count+1, column=col_count+1).value = val
		if i == 0:
			ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
		if i == 2:
			ws.cell(row=row_count+1, column=col_count+1).number_format = '0.00%'
			if val < .9:
				ws.cell(row=row_count+1, column=col_count+1).fill = redFill
			elif val > 1.2:
				ws.cell(row=row_count+1, column=col_count+1).fill = greenFill
		col_count += 1
		i += 1
	i = 0
	month_perf = get_perf(SMI, "Month")
	for val in month_perf:
		ws.cell(row=row_count+1, column=col_count+1).value = val
		if i == 0:
			ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
		if i == 2:
			ws.cell(row=row_count+1, column=col_count+1).number_format = '0.00%'
			if val < .9:
				ws.cell(row=row_count+1, column=col_count+1).fill = redFill
			elif val > 1.2:
				ws.cell(row=row_count+1, column=col_count+1).fill = greenFill
		col_count += 1
		i += 1
	i = 0
	last_month_perf = get_perf(SMI, "Prev")
	for val in last_month_perf:
		ws.cell(row=row_count+1, column=col_count+1).value = val
		if i == 0:
			ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
		if i == 2:
			ws.cell(row=row_count+1, column=col_count+1).number_format = '0.00%'
			if val < .9:
				ws.cell(row=row_count+1, column=col_count+1).fill = redFill
			elif val > 1.2:
				ws.cell(row=row_count+1, column=col_count+1).fill = greenFill
		col_count += 1
		i += 1
	off_days = get_off_days(SMI, dates)
	ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
	ws.cell(row=row_count+1, column=col_count+1).value = off_days
	row_count += 1 

wb.save(output)