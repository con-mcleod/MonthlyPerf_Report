#!/usr/bin/python3

import sys, csv, sqlite3, os, glob, re
from xlrd import open_workbook
from datetime import datetime
import time

##############################
#                            #
# PROGRAM EXECUTION          #
#                            #
##############################

if (len(sys.argv) != 2):
	print ("Usage: python3 salesforcify.py <SF file>")
	exit(1)

SF_file = sys.argv[1]
DATABASE = "dataset.db"
output_day = str(datetime.now().day)
output_month = str(datetime.now().month)
output_year = str(datetime.now().year)
output = output_year + "." + output_month + "." + output_day + ".xlsx"

##############################
#                            #
# SQL DATABASE CREATION      #
#                            #
##############################

connection = sqlite3.connect(DATABASE)
cursor = connection.cursor()

cursor.execute("DROP TABLE IF EXISTS SMI_DETAILS")
cursor.execute("DROP TABLE IF EXISTS FORECAST")
cursor.execute("DROP TABLE IF EXISTS ADJ_FORECAST")


cursor.execute("""CREATE TABLE IF NOT EXISTS SMI_DETAILS(
	SMI varchar(10),
	ref_no varchar(40),
	ECS varchar(150),
	installer varchar(60),
	PVsize float,
	panel_brand varchar(100),
	address varchar(150),
	postcode int,
	state varchar(10),
	site_status varchar(80),
	install_date date,
	supply_date date,
	tariff varchar(25),
	export_control int check(export_control in (0,1))
	)""")

cursor.execute("""CREATE TABLE IF NOT EXISTS FORECAST(
	SMI varchar(10),
	month int,
	val float
	)""")

cursor.execute("""CREATE TABLE IF NOT EXISTS ADJ_FORECAST(
	SMI varchar(10),
	month int,
	year int,
	adj_val float
	)""")

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

# return all SMIs from encompass report
def get_all_SMIs():
	query = "SELECT distinct(SMI) from DAILY_GEN"
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

# get the supply start date of given SMI
def get_supply_date(SMI):
	query = "SELECT supply_date from SMI_DETAILS where SMI=?"
	payload = (SMI)
	result = dbselect(query, payload)
	return result

# get forecast value for given SMI
def get_forecast(SMI, month):
	query = "SELECT val from FORECAST where SMI=? and month=?"
	payload = (SMI, month)
	result = dbselect(query, payload)
	return result

# return number of days in given month and year
def get_days_in_month(month, year):
	if month in [1,3,5,7,8,10,12]:
		return 31
	elif month in [4,6,9,11]:
		return 30
	else:
		if year in [16, 20, 24, 28]:
			return 29
		else:
			return 28

# assign headings to a specific column
def cols_to_nums(headings):
	heading_nums = {}
	num = 0
	for heading in headings:
		if re.search('SMI', heading):
			heading = "SMI"
		elif re.search('Reference', heading):
			heading = "ref_no"
		elif re.search('ECS', heading):
			heading = "ECS"
		elif re.search('Installer', heading):
			heading = "installer"
		elif re.search('Size', heading):
			heading = "PVsize"
		elif re.search('Brand', heading):
			heading = "panel_brand"
		elif re.search('Address', heading):
			heading = "address"
		elif re.search('Postcode', heading):
			heading = "postcode"
		elif re.search('State', heading):
			heading = "state"
		elif re.search('PPA', heading):
			heading = "site_status"
		elif re.search('Installation Date', heading):
			heading = "install_date"
		elif re.search('Supply', heading):
			heading = "supply_date"
		elif re.search('Export', heading):
			heading = "export_control"
		elif re.search('Tariff', heading):
			heading = "tariff"
		elif re.search('January', heading):
			heading = 1
		elif re.search('February', heading):
			heading = 2
		elif re.search('March', heading):
			heading = 3
		elif re.search('April', heading):
			heading = 4
		elif re.search('May', heading):
			heading = 5
		elif re.search('June', heading):
			heading = 6
		elif re.search('July', heading):
			heading = 7
		elif re.search('August', heading):
			heading = 8
		elif re.search('September', heading):
			heading = 9
		elif re.search('October', heading):
			heading = 10
		elif re.search('November', heading):
			heading = 11
		elif re.search('December', heading):
			heading = 12
		heading_nums[num] = heading
		num += 1
	return heading_nums

##############################
#                            #
# READ IN SALESFORCE REPORT  #
#                            #
##############################

wb = open_workbook(SF_file)
sheet = wb.sheet_by_index(0)

num_rows = sheet.nrows - 6
num_cols = sheet.ncols
headings = []

for row in range(0, 1):
	for col in range(0, num_cols):
		val = sheet.cell(row,col).value
		headings.append(val)

headings_dict = cols_to_nums(headings)

for row in range(1, num_rows):
	results = {}
	for col in range(0, num_cols):
		val = sheet.cell(row,col).value
		# val = re.sub(r'\"', '', val)

		for dkey, dval in headings_dict.items():
			if col == dkey:
				results[dval] = val

	SMI = results["SMI"]
	ref_no = results["ref_no"]
	ECS = results["ECS"]
	installer = results["installer"]
	PVsize = results["PVsize"]
	panel_brand = results["panel_brand"]
	address = results["address"]
	state = results["state"]
	if "postcode" in results:
		postcode = results["postcode"]
	else:
		postcode = ""
	site_status = results["site_status"]
	install_date = results["install_date"]
	if bool(install_date):
		dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(install_date) - 2)
		tt = dt.timetuple()
		install_date = time.strftime('%Y.%m.%d', tt)
	supply_date = results["supply_date"]
	if bool(supply_date):
		dt = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(supply_date) - 2)
		tt = dt.timetuple()
		supply_date = time.strftime('%Y.%m.%d', tt)
	export_control = results["export_control"]
	export_control = export_control.rstrip()
	if export_control == "Yes":
		export_control = 1
	else:
		export_control = 0
	tariff = results["tariff"]

	for key, value in results.items():
		if isinstance(key, int):
			value = re.sub(r'[^0-9\.]','',value)
			if value == "":
				value = 0
			value = float(value)
			cursor.execute("""INSERT OR IGNORE INTO forecast(SMI, month, val)
				VALUES (?,?,?)""", (SMI, key, value))

	cursor.execute("""INSERT OR IGNORE into SMI_DETAILS(SMI, ref_no, ECS, installer, 
			PVsize, panel_brand, address, postcode, state, site_status, install_date, 
			supply_date, tariff, export_control) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
			(SMI, ref_no, ECS, installer, PVsize, panel_brand, address, postcode, state, 
				site_status, install_date, supply_date, tariff, export_control))

connection.commit()
connection.close()

##############################
#                            #
# Forecast to adjusted       #
#                            #
##############################

connection = sqlite3.connect(DATABASE)
cursor = connection.cursor()

dates = get_all_months()
SMIs = get_all_SMIs()

for SMI in SMIs:
	supply_date = get_supply_date(SMI)
	if bool(supply_date) and supply_date[0][0] != '':
		supply_year = int(supply_date[0][0][2:4])
		supply_month = int(supply_date[0][0][5:7])
		supply_day = int(supply_date[0][0][8:10])
		for date in dates:
			month = date[0]
			year = date[1]
			adj_forecast = get_forecast(SMI[0], month)[0][0]
			if (year < supply_year):
				adj_forecast = 0
			elif (supply_year == year):
				if (supply_month == month):
					days_in_month = get_days_in_month(month, year)
					adj_forecast = adj_forecast * (1-(supply_day/days_in_month))
				elif (month < supply_month):
					adj_forecast = 0
				else:
					adj_forecast = adj_forecast
			
			cursor.execute("""INSERT OR IGNORE INTO adj_forecast(SMI, month, year, adj_val)
			VALUES (?,?,?,?)""", (SMI[0], month, year, adj_forecast))

	else:
		print (SMI[0], "does not have a supply date apparently")

connection.commit()
connection.close()


