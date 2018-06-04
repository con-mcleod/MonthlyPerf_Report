#!/usr/bin/python3

import sys, sqlite3, os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Color, Font, PatternFill, Border, Side
from datetime import datetime

DATABASE = "dataset.db"
output_day = str(datetime.now().day)
output_month = str(datetime.now().month)
output_year = str(datetime.now().year)

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

def get_last_date():
	query = """SELECT obs_day, obs_month, obs_year from DAILY_GEN
			group by obs_day, obs_month, obs_year order by obs_year, obs_month, obs_day"""
	payload = None
	all_dates = dbselect(query, payload)
	last_date = all_dates[-1]
	return last_date

# return SMI details from SF
def get_SMI_details(SMI):
	query = """SELECT ref_no, state, installer, PVsize, export_control,
		panel_brand, site_type, site_status, supply_date, tariff from SMI_DETAILS where SMI=?"""
	payload = (SMI)
	result = dbselect(query, payload)
	if not result:
		for i in range(0,22):
			result.append((''))
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

if (len(sys.argv) != 1):
	print ("Usage: python3 genMonthlyReport.py")
	exit(1)

last_date = get_last_date()
output = str(last_date[2])+"."+str(last_date[1])+"."+str(last_date[0]) + ".xlsx"

wb = load_workbook(output)
if 'Perf Report' in wb.sheetnames:
	ws = wb.get_sheet_by_name('Perf Report')
	ws2 = wb.create_sheet('Summary')
else:
	ws = wb.create_sheet('Perf Report')


redFill = PatternFill(start_color='FA5858', end_color='FA5858', fill_type='solid')
greenFill = PatternFill(start_color='9Afe2e', end_color='9Afe2e', fill_type='solid')
leftBorder = Border(left=Side(style='thin'))
rightBorder = Border(right=Side(style='thin'))

SMIs = get_all_SMIs()
dates = get_all_months()

ws_headings = ["SMI","Ref No","State","Installer","System Size","Export Control",
				"Panel Make","System Type","PPA Status","Supply Date","Tariff","Jan FC","Feb FC", "Mar FC",
				"Apr FC","May FC","Jun FC","Jul FC","Aug FC","Sep FC","Oct FC",
				"Nov FC","Dec FC"]
for date in dates:
	date = "gen(" + str(date).strip('()') + ")"
	ws_headings.append(date)
ws_headings.extend(["Annual FC","Annual Gen","Annual Perf","Quarter FC","Quarter Gen",
					"Quater Perf","Month FC","Month Gen","Month Perf","Prev FC",
					"Prev Gen","Prev Perf","Outage Days","Annual FC $","Annual Gen $","Shortfall $",
					"Quarter FC $","Quarter Gen $","Shortfall $","CurrMonth FC $","CurrMonth Gen $",
					"Shortfall $","PrevMonth FC $","PrevMonth Gen $","Shortfall $"])

ann_bucket_0_10=0
ann_bucket_10_20=0
ann_bucket_20_30=0
ann_bucket_30_40=0
ann_bucket_40_50=0
ann_bucket_50_60=0
ann_bucket_60_70=0
ann_bucket_70_80=0
ann_bucket_80_90=0
ann_bucket_90_100=0
ann_bucket_100_110=0
ann_bucket_110_120=0
ann_bucket_120_130=0
ann_bucket_130_inf=0
quart_bucket_0_10=0
quart_bucket_10_20=0
quart_bucket_20_30=0
quart_bucket_30_40=0
quart_bucket_40_50=0
quart_bucket_50_60=0
quart_bucket_60_70=0
quart_bucket_70_80=0
quart_bucket_80_90=0
quart_bucket_90_100=0
quart_bucket_100_110=0
quart_bucket_110_120=0
quart_bucket_120_130=0
quart_bucket_130_inf=0
prev_bucket_0_10=0
prev_bucket_10_20=0
prev_bucket_20_30=0
prev_bucket_30_40=0
prev_bucket_40_50=0
prev_bucket_50_60=0
prev_bucket_60_70=0
prev_bucket_70_80=0
prev_bucket_80_90=0
prev_bucket_90_100=0
prev_bucket_100_110=0
prev_bucket_110_120=0
prev_bucket_120_130=0
prev_bucket_130_inf=0
curr_bucket_0_10=0
curr_bucket_10_20=0
curr_bucket_20_30=0
curr_bucket_30_40=0
curr_bucket_40_50=0
curr_bucket_50_60=0
curr_bucket_60_70=0
curr_bucket_70_80=0
curr_bucket_80_90=0
curr_bucket_90_100=0
curr_bucket_100_110=0
curr_bucket_110_120=0
curr_bucket_120_130=0
curr_bucket_130_inf=0

row_count = 1
for SMI in SMIs:
	print (SMI)
	col_count = 0
	if row_count == 1:
		for heading in ws_headings:
			ws.cell(row=row_count, column=col_count+1).value = heading
			col_count += 1
	col_count = 1
	
	ws.cell(row=row_count+1, column=1).value = SMI[0]
	
	details = get_SMI_details(SMI)
	if details:
		if details[0]:
			for detail in details[0]:
				ws.cell(row=row_count+1, column=col_count+1).value = detail
				col_count += 1
		else:
			for i in range(0,len(details)):
				ws.cell(row=row_count+1, column=col_count+1).value = ''
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

	if not isinstance(details[0][3], str):
		if (annual_perf[2]<=.1):
			ann_bucket_0_10 += details[0][3]
		elif (annual_perf[2]>.1 and annual_perf[2]<=.2):
			ann_bucket_10_20 += details[0][3]
		elif (annual_perf[2]>.2 and annual_perf[2]<=.3):
			ann_bucket_20_30 += details[0][3]
		elif (annual_perf[2]>.3 and annual_perf[2]<=.4):
			ann_bucket_30_40 += details[0][3]
		elif (annual_perf[2]>.4 and annual_perf[2]<=.5):
			ann_bucket_40_50 += details[0][3]
		elif (annual_perf[2]>.5 and annual_perf[2]<=.6):
			ann_bucket_50_60 += details[0][3]
		elif (annual_perf[2]>.6 and annual_perf[2]<=.7):
			ann_bucket_60_70 += details[0][3]
		elif (annual_perf[2]>.7 and annual_perf[2]<=.8):
			ann_bucket_70_80 += details[0][3]
		elif (annual_perf[2]>.8 and annual_perf[2]<=.9):
			ann_bucket_80_90 += details[0][3]
		elif (annual_perf[2]>.9 and annual_perf[2]<=1.0):
			ann_bucket_90_100 += details[0][3]
		elif (annual_perf[2]>1.0 and annual_perf[2]<=1.1):
			ann_bucket_100_110 += details[0][3]
		elif (annual_perf[2]>1.1 and annual_perf[2]<=1.2):
			ann_bucket_110_120 += details[0][3]
		elif (annual_perf[2]>1.2 and annual_perf[2]<=1.3):
			ann_bucket_120_130 += details[0][3]
		elif (annual_perf[2]>1.3):
			ann_bucket_130_inf += details[0][3]

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

	if not isinstance(details[0][3], str):
		if (quarter_perf[2]<=.1):
			quart_bucket_0_10 += details[0][3]
		elif (quarter_perf[2]>.1 and quarter_perf[2]<=.2):
			quart_bucket_10_20 += details[0][3]
		elif (quarter_perf[2]>.2 and quarter_perf[2]<=.3):
			quart_bucket_20_30 += details[0][3]
		elif (quarter_perf[2]>.3 and quarter_perf[2]<=.4):
			quart_bucket_30_40 += details[0][3]
		elif (quarter_perf[2]>.4 and quarter_perf[2]<=.5):
			quart_bucket_40_50 += details[0][3]
		elif (quarter_perf[2]>.5 and quarter_perf[2]<=.6):
			quart_bucket_50_60 += details[0][3]
		elif (quarter_perf[2]>.6 and quarter_perf[2]<=.7):
			quart_bucket_60_70 += details[0][3]
		elif (quarter_perf[2]>.7 and quarter_perf[2]<=.8):
			quart_bucket_70_80 += details[0][3]
		elif (quarter_perf[2]>.8 and quarter_perf[2]<=.9):
			quart_bucket_80_90 += details[0][3]
		elif (quarter_perf[2]>.9 and quarter_perf[2]<=1.0):
			quart_bucket_90_100 += details[0][3]
		elif (quarter_perf[2]>1.0 and quarter_perf[2]<=1.1):
			quart_bucket_100_110 += details[0][3]
		elif (quarter_perf[2]>1.1 and quarter_perf[2]<=1.2):
			quart_bucket_110_120 += details[0][3]
		elif (quarter_perf[2]>1.2 and quarter_perf[2]<=1.3):
			quart_bucket_120_130 += details[0][3]
		elif (quarter_perf[2]>1.3):
			quart_bucket_130_inf += details[0][3]
	
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

	if not isinstance(details[0][3], str):
		if (month_perf[2]<=.1):
			prev_bucket_0_10 += details[0][3]
		elif (month_perf[2]>.1 and month_perf[2]<=.2):
			prev_bucket_10_20 += details[0][3]
		elif (month_perf[2]>.2 and month_perf[2]<=.3):
			prev_bucket_20_30 += details[0][3]
		elif (month_perf[2]>.3 and month_perf[2]<=.4):
			prev_bucket_30_40 += details[0][3]
		elif (month_perf[2]>.4 and month_perf[2]<=.5):
			prev_bucket_40_50 += details[0][3]
		elif (month_perf[2]>.5 and month_perf[2]<=.6):
			prev_bucket_50_60 += details[0][3]
		elif (month_perf[2]>.6 and month_perf[2]<=.7):
			prev_bucket_60_70 += details[0][3]
		elif (month_perf[2]>.7 and month_perf[2]<=.8):
			prev_bucket_70_80 += details[0][3]
		elif (month_perf[2]>.8 and month_perf[2]<=.9):
			prev_bucket_80_90 += details[0][3]
		elif (month_perf[2]>.9 and month_perf[2]<=1.0):
			prev_bucket_90_100 += details[0][3]
		elif (month_perf[2]>1.0 and month_perf[2]<=1.1):
			prev_bucket_100_110 += details[0][3]
		elif (month_perf[2]>1.1 and month_perf[2]<=1.2):
			prev_bucket_110_120 += details[0][3]
		elif (month_perf[2]>1.2 and month_perf[2]<=1.3):
			prev_bucket_120_130 += details[0][3]
		elif (month_perf[2]>1.3):
			prev_bucket_130_inf += details[0][3]
	
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
			ws.cell(row=row_count+1, column=col_count+1).border = rightBorder
		col_count += 1
		i += 1

	if not isinstance(details[0][3], str):
		if (last_month_perf[2]<=.1):
			curr_bucket_0_10 += details[0][3]
		elif (last_month_perf[2]>.1 and last_month_perf[2]<=.2):
			curr_bucket_10_20 += details[0][3]
		elif (last_month_perf[2]>.2 and last_month_perf[2]<=.3):
			curr_bucket_20_30 += details[0][3]
		elif (last_month_perf[2]>.3 and last_month_perf[2]<=.4):
			curr_bucket_30_40 += details[0][3]
		elif (last_month_perf[2]>.4 and last_month_perf[2]<=.5):
			curr_bucket_40_50 += details[0][3]
		elif (last_month_perf[2]>.5 and last_month_perf[2]<=.6):
			curr_bucket_50_60 += details[0][3]
		elif (last_month_perf[2]>.6 and last_month_perf[2]<=.7):
			curr_bucket_60_70 += details[0][3]
		elif (last_month_perf[2]>.7 and last_month_perf[2]<=.8):
			curr_bucket_70_80 += details[0][3]
		elif (last_month_perf[2]>.8 and last_month_perf[2]<=.9):
			curr_bucket_80_90 += details[0][3]
		elif (last_month_perf[2]>.9 and last_month_perf[2]<=1.0):
			curr_bucket_90_100 += details[0][3]
		elif (last_month_perf[2]>1.0 and last_month_perf[2]<=1.1):
			curr_bucket_100_110 += details[0][3]
		elif (last_month_perf[2]>1.1 and last_month_perf[2]<=1.2):
			curr_bucket_110_120 += details[0][3]
		elif (last_month_perf[2]>1.2 and last_month_perf[2]<=1.3):
			curr_bucket_120_130 += details[0][3]
		elif (last_month_perf[2]>1.3):
			curr_bucket_130_inf += details[0][3]
	
	off_days = get_off_days(SMI, dates)
	ws.cell(row=row_count+1, column=col_count+1).border = rightBorder
	ws.cell(row=row_count+1, column=col_count+1).value = off_days
	if (off_days):
		if off_days > 0:
			ws.cell(row=row_count+1, column=col_count+1).fill = redFill

	col_count += 1

	tariff = None
	if (details[0]):
		if details[0][9]:
			tariff = float(details[0][9])
			
	if tariff:
		revenues = []
		for val in annual_perf[:2]:
			revenue = (val*tariff)/100
			ws.cell(row=row_count+1, column=col_count+1).value = revenue
			revenues.append(revenue)
			col_count += 1
		ws.cell(row=row_count+1, column=col_count+1).value = (revenues[1]-revenues[0])
		ws.cell(row=row_count+1, column=col_count+1).border = rightBorder
		col_count += 1

		revenues = []
		for val in quarter_perf[:2]:
			revenue = (val*tariff)/100
			ws.cell(row=row_count+1, column=col_count+1).value = revenue
			revenues.append(revenue)
			col_count += 1
		ws.cell(row=row_count+1, column=col_count+1).value = (revenues[1]-revenues[0])
		ws.cell(row=row_count+1, column=col_count+1).border = rightBorder
		col_count += 1

		revenues = []
		for val in month_perf[:2]:
			revenue = (val*tariff)/100
			ws.cell(row=row_count+1, column=col_count+1).value = revenue
			revenues.append(revenue)
			col_count += 1
		ws.cell(row=row_count+1, column=col_count+1).value = (revenues[1]-revenues[0])
		ws.cell(row=row_count+1, column=col_count+1).border = rightBorder
		col_count += 1

		revenues = []
		for val in last_month_perf[:2]:
			revenue = (val*tariff)/100
			ws.cell(row=row_count+1, column=col_count+1).value = revenue
			revenues.append(revenue)
			col_count += 1
		ws.cell(row=row_count+1, column=col_count+1).value = (revenues[1]-revenues[0])
		ws.cell(row=row_count+1, column=col_count+1).border = rightBorder

	row_count += 1 

#####
# WS2
#####

ann_buckets = []
quart_buckets = []
prev_buckets = []
curr_buckets = []
ann_buckets.extend((ann_bucket_0_10, ann_bucket_10_20, ann_bucket_20_30, ann_bucket_30_40, ann_bucket_40_50, 
	ann_bucket_50_60, ann_bucket_60_70, ann_bucket_70_80, ann_bucket_80_90, ann_bucket_90_100, 
	ann_bucket_100_110, ann_bucket_110_120, ann_bucket_120_130, ann_bucket_130_inf))
quart_buckets.extend((quart_bucket_0_10, quart_bucket_10_20, quart_bucket_20_30, quart_bucket_30_40, quart_bucket_40_50, 
	quart_bucket_50_60, quart_bucket_60_70, quart_bucket_70_80, quart_bucket_80_90, quart_bucket_90_100, 
	quart_bucket_100_110, quart_bucket_110_120, quart_bucket_120_130, quart_bucket_130_inf))
prev_buckets.extend((prev_bucket_0_10, prev_bucket_10_20, prev_bucket_20_30, prev_bucket_30_40, prev_bucket_40_50, 
	prev_bucket_50_60, prev_bucket_60_70, prev_bucket_70_80, prev_bucket_80_90, prev_bucket_90_100, 
	prev_bucket_100_110, prev_bucket_110_120, prev_bucket_120_130, prev_bucket_130_inf))
curr_buckets.extend((curr_bucket_0_10, curr_bucket_10_20, curr_bucket_20_30, curr_bucket_30_40, curr_bucket_40_50, 
	curr_bucket_50_60, curr_bucket_60_70, curr_bucket_70_80, curr_bucket_80_90, curr_bucket_90_100, 
	curr_bucket_100_110, curr_bucket_110_120, curr_bucket_120_130, curr_bucket_130_inf))


row_count = 1
col_count = 0
for val in ann_buckets:
	ws2.cell(row=row_count+1, column=col_count+1).value = val
	row_count += 1
col_count += 1
row_count = 1
for val in quart_buckets:
	ws2.cell(row=row_count+1, column=col_count+1).value = val
	row_count += 1
col_count += 1
row_count = 1
for val in prev_buckets:
	ws2.cell(row=row_count+1, column=col_count+1).value = val
	row_count += 1
col_count += 1
row_count = 1
for val in curr_buckets:
	ws2.cell(row=row_count+1, column=col_count+1).value = val
	row_count += 1


wb.save(output)