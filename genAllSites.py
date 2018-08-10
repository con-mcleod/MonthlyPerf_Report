#!/usr/bin/python3

import sys, sqlite3, os
from openpyxl import Workbook
from openpyxl.styles import Color, Font, PatternFill, Border, Side
from datetime import datetime


##############################
#                            #
# Database handling          #
#                            #
##############################

# prepare sql query execution
def dbselect(cxn, query, payload):
	cursor = cxn.cursor()
	if not payload:
		rows = cursor.execute(query)
	else:
		rows = cursor.execute(query,payload)
	results = []
	for row in rows:
		results.append(row)
	cursor.close()
	return results

# execute sql query
def dbexecute(cxn, query, payload):
	cursor = cxn.cursor()
	if not payload:
		cursor.execute(query)
	else:
		cursor.execute(query, payload)


##############################
#                            #
# Helper functions           #
#                            #
##############################

# return the range of dates from encompass report
def get_all_months(cxn):
	query = """SELECT obs_month, obs_year from DAILY_GEN 
			group by obs_month, obs_year order by obs_year, obs_month"""
	payload = None
	all_dates = dbselect(cxn, query, payload)
	return all_dates

# return all SMIs from encompass report
def get_all_SMIs(cxn):
	query = "SELECT distinct(SMI) from DAILY_GEN"
	payload = None
	all_SMIs = dbselect(cxn, query, payload)
	return all_SMIs

def get_last_date(cxn):
	query = """SELECT obs_day, obs_month, obs_year from DAILY_GEN
			group by obs_day, obs_month, obs_year order by obs_year, obs_month, obs_day"""
	payload = None
	all_dates = dbselect(cxn, query, payload)
	last_date = all_dates[-1]
	return last_date

# return SMI's generation for given month
def get_month_gen(cxn, SMI, date):
	month = date[0]
	year = date[1]
	query = "SELECT val from MONTH_GEN where SMI=? and month=? and year=?"
	payload = (SMI, month, year)
	gen = dbselect(cxn, query, payload)
	return gen

# return number of off days for month
def get_off_days(cxn, SMI, dates):
	curr_date = dates[-1]
	curr_month = curr_date[0]
	curr_year = curr_date[1]
	off_days = 0
	query = """SELECT value from DAILY_GEN where SMI=? and obs_month=? 
				and obs_year=?"""
	payload = (SMI[0], curr_month, curr_year)
	result = dbselect(cxn, query, payload)
	if result:
		for val in result:
			if not val[0]:
				off_days += 1
			elif val[0] == 0:
				off_days += 1
	return off_days

##############################
#                            #
# GENERATE OUTPUT            #
#                            #
##############################

if __name__ == '__main__':

	if (len(sys.argv) != 1):
		print ("Usage: python3 genAllSites.py")
		exit(1)

	DATABASE = "dataset.db"
	output_day = str(datetime.now().day)
	output_month = str(datetime.now().month)
	output_year = str(datetime.now().year)

	cxn = sqlite3.connect(DATABASE)

	last_date = get_last_date(cxn)
	output = str(last_date[2])+"."+str(last_date[1])+"."+str(last_date[0]) + ".xlsx"

	if (os.path.exists(output)):
		os.remove(output)

	wb = Workbook()
	ws = wb.active
	ws.title = "All_sites"

	leftBorder = Border(left=Side(style='thin'))
	rightBorder = Border(right=Side(style='thin'))

	dates = get_all_months(cxn)
	SMIs = get_all_SMIs(cxn)

	row_count = 1
	for SMI in SMIs:
		col_count = 1
		ws.cell(row=1, column=1).value = "SMI"
		ws.cell(row=row_count+1, column=1).value = SMI[0]
		ws.cell(row=row_count+1, column=1).border = rightBorder

		print("Formatting SMI: " + SMI[0])

		for date in dates:
			ws.cell(row=1, column=col_count+1).value = str(date[0]) + "," + str(date[1])
			month_gen = get_month_gen(cxn, SMI[0], date)
			if month_gen:
				ws.cell(row=row_count+1, column=col_count+1).value = month_gen[0][0]
			col_count += 1
		ws.cell(row=1, column=col_count+1).value = "Outage Days"
		off_days = get_off_days(cxn, SMI, dates)
		ws.cell(row=row_count+1, column=col_count+1).border = leftBorder
		ws.cell(row=row_count+1, column=col_count+1).value = off_days
		row_count += 1

	wb.save(output)	

	print ("Complete!")