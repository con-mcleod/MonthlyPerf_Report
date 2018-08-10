#!/usr/bin/python3

import sys, csv, sqlite3, os, glob, re
from xlrd import open_workbook
from datetime import datetime
import time

##############################
#                            #
# SQL DATABASE CREATION      #
#                            #
##############################

def create_tables(cxn):
	cursor = cxn.cursor()

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
		export_control int check(export_control in (0,1)),
		site_type varchar(8)
		)""")

	cursor.execute("""CREATE TABLE IF NOT EXISTS FORECAST(
		SMI varchar(10),
		month int,
		val float
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

def forecast_insert(cxn, SMI, key, value):
	query = """INSERT OR IGNORE INTO forecast(SMI, month, val)
				VALUES (?,?,?)"""
	payload = (SMI, key, value)
	dbexecute(cxn, query, payload)

def smi_details_insert(cxn, SMI, ref_no, ECS, installer, PVsize, panel_brand, 
					address, postcode, state, site_status, install_date, 
					supply_date, tariff, export_control, site_type):
	query = """INSERT OR IGNORE into SMI_DETAILS(SMI, ref_no, ECS, installer, 
				PVsize, panel_brand, address, postcode, state, site_status, 
				install_date, supply_date, tariff, export_control, site_type) 
				VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"""
	payload = (SMI, ref_no, ECS, installer, PVsize, panel_brand, 
					address, postcode, state, site_status, install_date, 
					supply_date, tariff, export_control, site_type)
	dbexecute(cxn, query, payload)
	

##############################
#                            #
# READ IN SALESFORCE REPORT  #
#                            #
##############################

if __name__ == '__main__':

	if (len(sys.argv) != 2):
		print ("Usage: python3 salesforcify.py <SF file>")
		exit(1)

	SF_file = sys.argv[1]
	DATABASE = "dataset.db"

	cxn = sqlite3.connect(DATABASE)
	create_tables(cxn)


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
	SMI_count = 0

	for row in range(1, num_rows):
		results = {}
		for col in range(0, num_cols):
			val = sheet.cell(row,col).value
			# val = re.sub(r'\"', '', val)

			for dkey, dval in headings_dict.items():
				if col == dkey:
					results[dval] = val

		SMI = results["SMI"]
		if isinstance(SMI,str) == False:
			SMI = str(SMI)[0:10]
		ref_no = results["ref_no"]
		ECS = results["ECS"]
		installer = results["installer"]
		PVsize = results["PVsize"]
		if PVsize:
			PVsize = float(PVsize)
			if (PVsize > 100):
				site_type = "C&I"
			elif (SMI[0] in ["A","B","C","D","E","F","G"]):
				site_type = "SME"
			elif (SMI[0] in ["W","X","Y","Z"]):
				site_type = "Resi"
		else:
			site_type = ''
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
		if tariff:
			tariff = re.sub(r'[^0-9\.]','',tariff[5:])
		if (SMI == "6203778594" or SMI == "6203779394"):
			tariff = 9
		elif (SMI == "B162191181" or SMI == "B165791182" or SMI == "D170092557"):
			tariff = 14
		elif SMI == "C172991611":
			tariff = 16.02
		elif (SMI == "G161391137" or SMI == "G161391138"):
			tariff = 19.63

		for key, value in results.items():
			if isinstance(key, int):
				# value = re.sub(r'[^0-9\.]','',value)
				if value == "":
					value = 0
				value = float(value)
				forecast_insert(cxn, SMI, key, value)
				
		print ("Collected SMI details for: " + SMI)
		SMI_count += 1

		smi_details_insert(cxn, SMI, ref_no, ECS, installer, PVsize, panel_brand, address, postcode, state, 
					site_status, install_date, supply_date, tariff, export_control, site_type)

	cxn.commit()
	cxn.close()

	print ("Complete!")
	print ("Collected details for " + str(SMI_count) + " unique SMIs")