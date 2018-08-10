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

# return all SMIs from encompass report
def get_all_SMIs(cxn):
	query = "SELECT distinct(SMI) from DAILY_GEN"
	payload = None
	all_SMIs = dbselect(cxn, query, payload)
	return all_SMIs

# return the range of dates from encompass report
def get_all_months(cxn):
	query = """SELECT obs_month, obs_year from DAILY_GEN 
			group by obs_month, obs_year order by obs_year, obs_month"""
	payload = None
	all_dates = dbselect(cxn, query, payload)
	return all_dates

# return SMI's generation for the given (month, year)
def get_month_gen(cxn, SMI, date):
	month = date[0]
	year = date[1]
	query = """SELECT sum(value), datatype from DAILY_GEN where SMI=? and 
			obs_month=? and obs_year=?"""
	payload = (SMI[0], month, year,)
	result = dbselect(cxn, query, payload)
	return result

# return SMI's monthly generation starting from supply day
def get_adj_month_gen(cxn, SMI, date, supply_day):
	month = date[0]
	year = date[1]
	query = """SELECT sum(value), datatype from DAILY_GEN where SMI=? and
			obs_month=? and obs_year=? and obs_day>=?"""
	payload = (SMI[0], month, year, supply_day)
	result = dbselect(cxn, query, payload)
	return result

# get the supply start date of given SMI
def get_supply_date(cxn, SMI):
	query = "SELECT supply_date from SMI_DETAILS where SMI=?"
	payload = (SMI)
	result = dbselect(cxn, query, payload)
	return result

# get forecast value for given SMI
def get_forecast(cxn, SMI, month):
	query = "SELECT val from FORECAST where SMI=? and month=?"
	payload = (SMI, month)
	result = dbselect(cxn, query, payload)
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

def month_gen_insert(cxn, SMI, month, year, gen):
	query = """INSERT OR IGNORE INTO MONTH_GEN(SMI, month, year, val)
		VALUES (?,?,?,?)"""
	payload = (SMI, month, year, gen)
	dbexecute(cxn, query, payload)

def adj_forecast_insert(cxn, SMI, month, year, adj_forecast):
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

	if (len(sys.argv) != 1):
		print ("Usage: python3 adjuster.py")
		exit(1)

	DATABASE = "dataset.db"

	cxn = sqlite3.connect(DATABASE)
	create_tables(cxn)


	SMIs = get_all_SMIs(cxn)
	dates = get_all_months(cxn)
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
				if (SMI[0]=="6203778594" or SMI[0]=="6203779394"):
					if date[1] == 16:
						adj_forecast = 0.933*adj_forecast
					elif date[1] == 17:
						adj_forecast = 0.926*adj_forecast
					elif date[1] == 18:
						adj_forecast = 0.919*adj_forecast
				

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
						month_gen = get_month_gen(cxn, SMI, date)
						adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]
				else:
					month_gen = get_month_gen(cxn, SMI, date)
					adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]

				month_gen_insert(cxn, SMI[0], month, year, month_gen[0][0])
				adj_forecast_insert(cxn, SMI[0], month, year, adj_forecast)
		else:
			print (SMI[0], "does not have a supply date apparently so forecast remains the same")
			for date in dates:
				month = date[0]
				year = date[1]
				adj_forecast = get_forecast(cxn, SMI[0], month)[0][0]
				adj_forecast_insert(cxn, SMI[0], month, year, adj_forecast)


	cxn.commit()
	cxn.close()

	print ("Complete!")


