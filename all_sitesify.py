#!/usr/bin/python3

import sys, csv, sqlite3, os, glob, re
from collections import defaultdict
from openpyxl import Workbook

##############################
#                            #
# PROGRAM EXECUTION          #
#                            #
##############################

if (len(sys.argv) != 2):
	print ("Usage: python3 all_sitesify.py <encompass_folder>")
	exit(1)

enc_folder = sys.argv[1]
DATABASE = "dataset.db"
monthList = [1,2,3,4,5,6,7,8,9,10,11,12]

##############################
#                            #
# SQL DATABASE CREATION      #
#                            #
##############################

connection = sqlite3.connect(DATABASE)
cursor = connection.cursor()

cursor.execute("""CREATE TABLE IF NOT EXISTS DAILY_GEN(
	SMI varchar(10),
	datatype varchar(30), 
	obs_day int,
	obs_month int,
	obs_year int,
	value float,
	unique(SMI, obs_day, obs_month, obs_year)
	)""")

cursor.execute("""CREATE TABLE IF NOT EXISTS MONTH_GEN(
	SMI varchar(10),
	month int,
	year int,
	val float,
	unique(SMI, month, year)
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

# prepare sql query execution without connection
def dbselect2(query, payload):
	if not payload:
		rows = cursor.execute(query)
	else:
		rows = cursor.execute(query,payload)
	results = []
	for row in rows:
		results.append(row)
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

# convert month from string to int
def month_to_num(month):
	if month == "Jan": 
		return 1
	elif month == "Feb":
		return 2
	elif month == "Mar":
		return 3
	elif month == "Apr":
		return 4
	elif month == "May":
		return 5
	elif month == "Jun":
		return 6
	elif month == "Jul":
		return 7
	elif month == "Aug":
		return 8
	elif month == "Sep":
		return 9
	elif month == "Oct":
		return 10
	elif month == "Nov":
		return 11
	elif month == "Dec":
		return 12
	else:
		return 0

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

# return SMI's generation for the given (month, year)
def get_month_gen(SMI, date):
	month = date[0]
	year = date[1]
	query = """SELECT sum(value), datatype from DAILY_GEN where SMI=? and 
			obs_month=? and obs_year=?"""
	payload = (SMI[0], month, year,)
	result = dbselect2(query, payload)
	return result

##############################
#                            #
# READ IN ENCOMPASS REPORT   #
#                            #
##############################

encompass_files = os.path.join(enc_folder,"*")

for file in glob.glob(encompass_files):

	row_count = 0
	columns = defaultdict(list)
	enc_SMIs = []

	with open(file,'r', encoding='utf-8') as enc_in:
		reader = csv.reader(enc_in)
		for row in reader:
			for (i,v) in enumerate(row):
				columns[i].append(v)
			row_count += 1

	for col in columns:
		SMIs = re.findall(r'[a-zA-Z0-9]{10}',columns[col][0])
		datatype = columns[col][0].split("- ")[-1]
		enc_dataset = columns[col][1:]
		dates = columns[0][1:]

		for SMI in SMIs:
			if bool(re.search(r'\d', SMI)) == False:
				pass
			else:
				if SMI not in enc_SMIs:
					enc_SMIs.append(SMI)
					print ("Storing daily data for SMI: " + SMI)
				if (SMI and (datatype=="kWh Generation" or datatype=="kWh Generation Generation"
					or datatype=="kWh Generation B1")):
					for i in range(row_count-1):
						if (len(dates[i]) < 11):
							day = re.sub(r'-.*', '', dates[i])
							month = re.sub(r'[^a-zA-Z]', '', dates[i])
							month = month_to_num(month)
							year = re.sub(r'.*-', '', dates[i])
						else:
							day = dates[i][4:6]
							month = month_to_num(dates[i][7:10])
							year = dates[i][13:15]
						if (month in monthList):
							cursor.execute("""INSERT OR IGNORE INTO DAILY_GEN(SMI, datatype, 
								obs_day, obs_month, obs_year, value) VALUES (?,?,?,?,?,?)""", 
								(SMI, datatype, day, month, year, enc_dataset[i]))

connection.commit()
connection.close()

##############################
#                            #
# Daily gen to monthly gen   #
#                            #
##############################

connection = sqlite3.connect(DATABASE)
cursor = connection.cursor()

SMIs = get_all_SMIs()
dates = get_all_months()
for SMI in SMIs:
	print ("Collating monthly data for SMI: " + SMI[0])
	for date in dates:
		month_gen = get_month_gen(SMI, date)
		cursor.execute("""INSERT OR IGNORE INTO MONTH_GEN(SMI, month, year, val)
			VALUES (?,?,?,?)""", (SMI[0], date[0], date[1], month_gen[0][0]))
connection.commit()
connection.close()

print ("Complete!")