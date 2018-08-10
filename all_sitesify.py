#!/usr/bin/python3

# Author: Connor McLeod
# Contact: con.mcleod92@gmail.com
# Source code: https://github.com/con-mcleod/MonthlyPerf_Report
# Latest Update: 10 August 2018

import sys, csv, sqlite3, os, glob, re
from collections import defaultdict

monthList = [1,2,3,4,5,6,7,8,9,10,11,12]

##############################
#                            #
# SQL DATABASE CREATION      #
#                            #
##############################

def create_tables(cxn):
	"""
	Function to create tables in sqlite3
	:param cxn: the connection to the sqlite3 database
	:return:
	"""

	cursor = cxn.cursor()

	cursor.execute("DROP TABLE IF EXISTS DAILY_GEN")

	cursor.execute("""CREATE TABLE IF NOT EXISTS DAILY_GEN(
		SMI varchar(10),
		datatype varchar(30), 
		obs_day int,
		obs_month int,
		obs_year int,
		value float,
		unique(SMI, obs_day, obs_month, obs_year)
		)""")

	cursor.close()

##############################
#                            #
# Database handling          #
#                            #
##############################

def dbselect(cxn, query, payload):
	"""
	Function to select data from an sqlite3 table
	:param cxn: connection to the sqlite3 database
	:param query: the query to be run
	:param payload: the payload for any query parameters
	:return results: the results of the search
	"""
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

def dbexecute(cxn, query, payload):
	"""
	Function to execute an sqlite3 table insertion
	:param cxn: connection to the sqlite3 database
	:param query: the query to be run
	:param payload: the payload for any query parameters
	:return:
	"""
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

def month_to_num(month):
	"""
	Function to convert the month string to an integer
	:param month: the month as a String
	:return: return an integer to represent the month
	"""
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


def daily_gen_insert(cxn, SMI, datatype, day, month, year, enc_dataset):
	"""
	Function to insert into the daily_gen table
	:param cxn: the connection to the sqlite3 table
	:param SMI: the SMI
	:param datatype: the type of data collected - in this case it is kWh Generation
	:param day: the day of interest
	:param month: the month of interest
	:param year: the year of interest
	:param enc_dataset: the generation data from Encompass
	:return:
	"""
	query = """INSERT OR IGNORE INTO DAILY_GEN(SMI, datatype, 
		obs_day, obs_month, obs_year, value) VALUES (?,?,?,?,?,?)"""
	payload = (SMI, datatype, day, month, year, enc_dataset)
	dbexecute(cxn, query, payload)

##############################
#                            #
# READ IN ENCOMPASS REPORT   #
#                            #
##############################

if __name__ == '__main__':

	# terminate program if not executed correctly
	if (len(sys.argv) != 2):
		print ("Usage: python3 all_sitesify.py <encompass_folder>")
		exit(1)

	# set up the locations for data retrieval and storage, connect to db and create tables
	DATABASE = "dataset.db"
	cxn = sqlite3.connect(DATABASE)
	create_tables(cxn)
	enc_folder = sys.argv[1]
	
	# variables for counting how much data is processed
	data_count = 0
	SMI_count = 0

	# for each encompass file in the encompass folder collect and store the data
	encompass_files = os.path.join(enc_folder,"*")
	for file in glob.glob(encompass_files):

		row_count = 0
		columns = defaultdict(list)
		enc_SMIs = []

		# read each csv file by transcribing rows to columns for simpler extraction
		with open(file,'r', encoding='utf-8') as enc_in:
			reader = csv.reader(enc_in)
			for row in reader:
				for (i,v) in enumerate(row):
					columns[i].append(v)
				row_count += 1

		# for each column extract the specific data and store into the database
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
						print ("Collating daily data for SMI: " + SMI)
					if (SMI and (datatype=="kWh Generation" or datatype=="kWh Generation Generation"
						or datatype=="kWh Generation B1")):
						SMI_count += 1
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
								daily_gen_insert(cxn, SMI, datatype, day, month, year, enc_dataset[i])
							data_count += 1

	cxn.commit()
	cxn.close()

	print ("Complete!")
	print ("Collated " + str(data_count) + " data points for " + str(SMI_count) + " unique SMIs")