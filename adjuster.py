#!/usr/bin/python3

# Author: Connor McLeod
# Contact: con.mcleod92@gmail.com
# Source code: https://github.com/con-mcleod/MonthlyPerf_Report
# Latest Update: 10 August 2018

import sys, csv, sqlite3, os, glob, re
from xlrd import open_workbook

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

	cursor.execute("DROP TABLE IF EXISTS ADJ_FORECAST")
	cursor.execute("DROP TABLE IF EXISTS MONTH_GEN")

	cursor.execute("""CREATE TABLE IF NOT EXISTS ADJ_FORECAST(
		SMI varchar(10),
		month int,
		year int,
		adj_val float
		)""")

	cursor.execute("""CREATE TABLE IF NOT EXISTS MONTH_GEN(
		SMI varchar(10),
		month int,
		year int,
		val float,
		unique(SMI, month, year)
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

def get_all_SMIs(cxn):
	"""
	Function to grab all SMIs from the Encompass reports
	:param cxn: connection to sqlite3 database
	:return all_SMIs: list of all SMIs
	"""
	query = "SELECT distinct(SMI) from DAILY_GEN"
	payload = None
	all_SMIs = dbselect(cxn, query, payload)
	return all_SMIs


def get_all_months(cxn):
	"""
	Function to return all months included in the Encompass reports
	:param cxn: connection to sqlite3 database
	:return all_dates: list of all months in report [mm, yy]
	"""
	query = """SELECT obs_month, obs_year from DAILY_GEN 
			group by obs_month, obs_year order by obs_year, obs_month"""
	payload = None
	all_dates = dbselect(cxn, query, payload)
	return all_dates


def create_month_gen(cxn, SMI, date):
	"""
	Function to create the monthly generation for an SMI in a given month
	:param cxn: connection to sqlite3 database
	:param SMI: the SMI of interest
	:param date: the month and year of interest
	:return result: the monthly generation
	"""
	month = date[0]
	year = date[1]
	query = """SELECT sum(value), datatype from DAILY_GEN where SMI=? and 
			obs_month=? and obs_year=?"""
	payload = (SMI[0], month, year,)
	result = dbselect(cxn, query, payload)
	return result


def get_adj_month_gen(cxn, SMI, date, supply_day):
	"""
	Function to return an SMI's monthly generation with it's supply date
	occuring in that month
	:param cxn: connection to sqlite3 database
	:param SMI: the SMI of interest
	:param date: the month and year of interest
	:param supply_day: the supply date for the SMI to adjust generation to
	:return result: the adjusted monthly generation having removed generation before supply date
	"""
	month = date[0]
	year = date[1]
	query = """SELECT sum(value), datatype from DAILY_GEN where SMI=? and
			obs_month=? and obs_year=? and obs_day>=?"""
	payload = (SMI[0], month, year, supply_day)
	result = dbselect(cxn, query, payload)
	return result


def get_supply_date(cxn, SMI):
	"""
	Function to get the supply date of an SMI from the Salesforce data
	:param cxn: connection to sqlite3 database
	:param SMI: the SMI of interest
	:return result: the supply date string
	"""
	query = "SELECT supply_date from SMI_DETAILS where SMI=?"
	payload = (SMI)
	result = dbselect(cxn, query, payload)
	return result


def get_forecast(cxn, SMI, month):
	"""
	Function to get the forecast for the given SMI in a specific month
	:param cxn: connection to sqlite3 database
	:param SMI: the SMI of interest:
	:param month: the month of interest
	:return result: the forecast value for that SMI for that month
	"""
	query = "SELECT val from FORECAST where SMI=? and month=?"
	payload = (SMI, month)
	result = dbselect(cxn, query, payload)
	return result


def get_days_in_month(month, year):
	"""
	Function to return the number of days in a month given a month and year
	:param month: the month of interest
	:param year: the year of interest
	:return: number of days in the given month
	"""
	if month in [1,3,5,7,8,10,12]:
		return 31
	elif month in [4,6,9,11]:
		return 30
	else:
		if year in [16, 20, 24, 28]:
			return 29
		else:
			return 28


def month_gen_insert(cxn, SMI, month, year, gen):
	"""
	Function to insert data in the sqlite table month_gen
	:param cxn: connection to sqlite3 database
	:param SMI: the SMI of interest
	:param month: the month in which the generation data occurred
	:param year: the year in which the generation data occurred
	:param gen: the generation data after supply date adjustment
	:return:
	"""
	query = """INSERT OR IGNORE INTO MONTH_GEN(SMI, month, year, val)
		VALUES (?,?,?,?)"""
	payload = (SMI, month, year, gen)
	dbexecute(cxn, query, payload)


def adj_forecast_insert(cxn, SMI, month, year, adj_forecast):
	"""
	Function to insert data in the sqlite table adj_forecast
	:param cxn: connection to sqlite3 database
	:param SMI: the SMI of interest
	:param month: the month in which the generation data occurred
	:param year: the year in which the generation data occurred
	:param adj_forecast: the forecast value after supply date adjustment
	:return:
	"""
	query = """INSERT OR IGNORE INTO adj_forecast(SMI, month, year, adj_val)
			VALUES (?,?,?,?)"""
	payload = (SMI, month, year, adj_forecast)
	dbexecute(cxn, query, payload)

##############################
#                            #
# Translating monthly data   #
#                            #
##############################


if __name__ == '__main__':

	# terminate program if not executed correctly
	if (len(sys.argv) != 1):
		print ("Usage: python3 adjuster.py")
		exit(1)

	# connect to the database and create the tables
	DATABASE = "dataset.db"
	cxn = sqlite3.connect(DATABASE)
	create_tables(cxn)

	# get all the sites of interest and dates of data
	SMIs = get_all_SMIs(cxn)
	dates = get_all_months(cxn)

	# for each site manipulate the given data for required adjustments
	for SMI in SMIs:

		print ("Collating monthly data and adjusting forecast for SMI: " + SMI[0])

		supply_date = get_supply_date(cxn, SMI)

		if bool(supply_date) and supply_date[0][0] != '':
			supply_year = int(supply_date[0][0][2:4])
			supply_month = int(supply_date[0][0][5:7])
			supply_day = int(supply_date[0][0][8:10])

			for date in dates:

				month = date[0]
				year = date[1]

				adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]

				# logic to adjust generation/forecast based on supply date
				if (year < supply_year):
					month_gen = [[0]]
					adj_forecast = 0
				elif (supply_year == year):
					if (month == supply_month):
						month_gen = get_adj_month_gen(cxn, SMI, date, supply_day)
						days_in_month = get_days_in_month(month, year)
						adj_forecast = adj_forecast * (1-(supply_day/days_in_month))
					elif (month < supply_month):
						month_gen = [[0]]
						adj_forecast = 0
					else:
						month_gen = create_month_gen(cxn, SMI, date)
						adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]
				else:
					month_gen = create_month_gen(cxn, SMI, date)
					adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]

				# hardcoded solution for the solar farm sites
				if (SMI[0]=="6203778594" or SMI[0]=="6203779394"):
					adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]
					if year == 16:
						adj_forecast = 0.933*adj_forecast
					elif year == 17:
						adj_forecast = 0.926*adj_forecast
					elif year == 18:
						adj_forecast = 0.919*adj_forecast

				month_gen_insert(cxn, SMI[0], month, year, month_gen[0][0])
				adj_forecast_insert(cxn, SMI[0], month, year, adj_forecast)
		else:
			# this handles sites that do not have a supply date populated in Salesforce
			print (SMI[0], "does not have a supply date apparently so forecast remains the same")
			for date in dates:
				month = date[0]
				year = date[1]
				adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]
				adj_forecast_insert(cxn, SMI[0], month, year, adj_forecast)


	cxn.commit()
	cxn.close()

	print ("Complete!")


