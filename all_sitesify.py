#!/usr/bin/python3

import sys, csv, sqlite3, os, glob, re
from collections import defaultdict
from openpyxl import Workbook


##############################
#                            #
# SQL DATABASE CREATION      #
#                            #
##############################

def create_tables(cxn):

	cursor = cxn.cursor()

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

def daily_gen_insert(cxn, SMI, datatype, day, month, year, enc_dataset):
	query = """INSERT OR IGNORE INTO DAILY_GEN(SMI, datatype, 
		obs_day, obs_month, obs_year, value) VALUES (?,?,?,?,?,?)"""
	payload = (SMI, datatype, day, month, year, enc_dataset)
	dbexecute(cxn, query, payload)
	return

##############################
#                            #
# READ IN ENCOMPASS REPORT   #
#                            #
##############################

if __name__ == '__main__':

	if (len(sys.argv) != 2):
		print ("Usage: python3 all_sitesify.py <encompass_folder>")
		exit(1)

	enc_folder = sys.argv[1]
	DATABASE = "dataset.db"
	monthList = [1,2,3,4,5,6,7,8,9,10,11,12]

	cxn = sqlite3.connect(DATABASE)
	create_tables(cxn)

	encompass_files = os.path.join(enc_folder,"*")
	data_count = 0
	SMI_count = 0

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