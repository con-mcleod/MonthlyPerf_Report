#!/usr/bin/python3

# Author: Connor McLeod
# Contact: con.mcleod92@gmail.com
# Source code: https://github.com/con-mcleod/MonthlyPerf_Report
# Latest Update: 10 August 2018

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
	"""
	Function to create tables in sqlite3
	:param cxn: the connection to the sqlite3 database
	:return:
	"""

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

def cols_to_nums(headings):
	"""
	Function to find the relevant heading and assign it to a variable
	This function makes it so that the if the Salesforce report changes in structure
	the report will still continue to work
	:param headings: the Salesforce report headings
	:return headings_nums: a dictionary which has an ordered list of required headings
	"""
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


def forecast_insert(cxn, SMI, month, value):
	"""
	Function to insert data into the sqlite forecast table
	:param cxn: the connection to the sqlite3 database
	:param SMI: the given SMI
	:param month: the given month
	:param value: the forecast value for that SMI and month from Salesforce report 
	:return:
	"""
	query = """INSERT OR IGNORE INTO forecast(SMI, month, val)
				VALUES (?,?,?)"""
	payload = (SMI, month, value)
	dbexecute(cxn, query, payload)


def smi_details_insert(cxn, SMI, ref_no, ECS, installer, PVsize, panel_brand, 
					address, postcode, state, site_status, install_date, 
					supply_date, tariff, export_control, site_type):
	"""
	Function to insert data into the sqlite smi_details table
	:param cxn: the connection to the sqlite3 database
	:param SMI: SMI from SF report
	:param ref_no: reference number from SF report
	:param ECS: ECS order number from SF report
	:param installer: installer from SF report
	:param PVsize: PV system size from SF report
	:param panel_brand: panel make from SF report
	:param address:	address from SF report
	:param postcode: postcode from SF report
	:param state: state from SF report
	:param site_status: status of site from SF report
	:param install_date: install date from SF report
	:param supply_date: supply date from SF report
	:param tariff: tariff as an integer from SF report
	:param export_control: export control boolean from SF report
	:param site_type: type of site (Resi, SME, C&I) from SF report
	:return:
	"""
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

	# terminate program if not executed correctly
	if (len(sys.argv) != 2):
		print ("Usage: python3 salesforcify.py <SF file>")
		exit(1)

	# connect to the database and create the tables
	DATABASE = "dataset.db"
	cxn = sqlite3.connect(DATABASE)
	create_tables(cxn)

	# read from Salesforce excel file using xlrd package
	SF_file = sys.argv[1]
	wb = open_workbook(SF_file)
	sheet = wb.sheet_by_index(0)

	# reads in the excel headings and allocates them to variables regardless of excel file order
	num_rows = sheet.nrows - 6
	num_cols = sheet.ncols
	headings = []

	for row in range(0, 1):
		for col in range(0, num_cols):
			val = sheet.cell(row,col).value
			headings.append(val)

	# function to assign headers to specific variables
	headings_dict = cols_to_nums(headings)
	SMI_count = 0

	# for each SMI grab the relevant data and store it into the sqlite tables
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

		# hardcoded tariffs for specific sites that aren't populated in Salesforce
		if (SMI == "6203778594" or SMI == "6203779394"):
			tariff = 9
		elif (SMI == "B162191181" or SMI == "B165791182" or SMI == "D170092557"):
			tariff = 14
		elif SMI == "C172991611":
			tariff = 16.02
		elif (SMI == "G161391137" or SMI == "G161391138"):
			tariff = 19.63

		# insert into the forecast table the forecast values for the given SMI
		# key = month
		for key, value in results.items():
			if isinstance(key, int):
				# value = re.sub(r'[^0-9\.]','',value)
				if value == "":
					value = 0
				value = float(value)
				forecast_insert(cxn, SMI, key, value)
				
		print ("Collected SMI details for: " + SMI)
		SMI_count += 1

		# inesrt into the smi_details table the details for given smi
		smi_details_insert(cxn, SMI, ref_no, ECS, installer, PVsize, panel_brand, address, postcode, state, 
					site_status, install_date, supply_date, tariff, export_control, site_type)

	cxn.commit()
	cxn.close()

	print ("Complete!")
	print ("Collected details for " + str(SMI_count) + " unique SMIs")