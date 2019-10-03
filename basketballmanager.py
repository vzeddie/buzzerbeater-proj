#!/usr/bin/python

import requests
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import xmltodict
import datetime
import time
import argparse

# ------------------- #
# 		GOOGLE		  #
# ------------------- #
def gsheets_get_row(sheet_obj, row_num):
	return sheet_obj.row_values(row_num)

def gsheets_set_row(sheet_obj, row_num, new_player_info):
	cell_range = sheet_obj.range("H{}:T{}".format(row_num + 1, row_num + 1))
	ctr = 0
	for cell in cell_range:
		cell.value = new_player_info[ctr]
		ctr = ctr + 1
	sheet_obj.update_cells(cell_range)

def gsheets_update_timestamp(sheet_obj, row_num, timestamp):
	cell_range = sheet_obj.range("Y{}:Y{}".format(row_num + 1, row_num +1 ))
	for cell in cell_range:
		cell.value = timestamp
	sheet_obj.update_cells(cell_range)

# Returns the row number for the playerid or -1 if not found
def gsheets_find_playerid(sheet_obj, id_column, playerid):
	try:
		return sheet_obj.col_values(id_column).index(str(playerid))
	except:
		return -1


# ------------------- #
# 		 BBAPI		  #
# ------------------- #

# Public player information
def get_player_info(session, playerid):
	r = session.get('http://bbapi.buzzerbeater.com/player.aspx?playerid={}'.format(playerid))
	xmldict = xmltodict.parse(r.text)
	return xmldict['bbapi']['player']
def map_player_info(player_info):
	ans = list()
	for item in player_info['skills']:
		try:
			ans.append(player_info['skills'][item]["#text"])
		except:
			ans.append(player_info['skills'][item])
	# Exclude "experience"
	return ans[1:-1]


def bbapi_main(userid, apikey, playerid):
	# Start a session which will retain auth cookie
	s = requests.Session()
	# Required to login
	s.get('http://bbapi.buzzerbeater.com/login.aspx?login={}&code={}'.format(userid, apikey))

	# ------------------- #
	# 	  MAIN LOGIC	  #
	# ------------------- #
	tmp = get_player_info(s, playerid)
	tmp = map_player_info(tmp)

	# Required to log out to invalidate session cookie
	s.get('http://bbapi.buzzerbeater.com/logout.aspx')

	return tmp

def main(username, apikey, playerid, worksheet_name=None):
	# Get player info from BBAPI
	new_player_info = bbapi_main(username, apikey, playerid)

	# Specify spreadsheet information
	SHEET_KEY = "REDACTED"
	PLAYERID_COLUMN = 2
	# use creds to create a client to interact with the Google Drive API
	scope = ['https://spreadsheets.google.com/feeds']
	creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
	client = gspread.authorize(creds)


	# Update spreadsheet with new player info

	# Automated version - slower
	if worksheet_name is None:
		sheet = client.open_by_key(SHEET_KEY)
		for ws in sheet.worksheets():
			print "Checking worksheet: {} for playerID #{}".format(ws.title, playerid)
			row_num = gsheets_find_playerid(sheet.worksheet(ws.title), PLAYERID_COLUMN, playerid)
			if row_num >= 0:
				found_ws = sheet.worksheet(ws.title)
				print "PlayerID #{} found in worksheet {}. Updating...".format(playerid, ws.title)
				gsheets_set_row(found_ws, row_num, new_player_info)
				gsheets_update_timestamp(found_ws, row_num, datetime.date.today().strftime("%b/%d/%Y"))
				return True
		print "Unable to find playerID #{}".format(playerid)
		return False
	# Explicit version - faster
	else:
		print "Attempting to access worksheet: {} for playerID #{}".format(worksheet_name, playerid)
		try:
			sheet = client.open_by_key(SHEET_KEY).worksheet(worksheet_name)
			row_num = gsheets_find_playerid(sheet, PLAYERID_COLUMN, playerid)
			if row_num >= 0:
				print "PlayerID #{} found in worksheet {}. Updating...".format(playerid, worksheet_name)
				gsheets_set_row(sheet, row_num, new_player_info)
				gsheets_update_timestamp(sheet, row_num, datetime.date.today().strftime("%b/%d/%Y"))
				return True
			else:
				print "Unable to find playerID #{} in worksheet {}".format(playerid, worksheet_name)
		except:
			print "ERROR: Unable to access worksheet: {}".foramt(worksheet_name)


if __name__ == "__main__":

	# Get user arguments
	parser = argparse.ArgumentParser(description="Queries BuzzerBeater API for skills of a player <PLAYERID> and updates the skills for that player in the Google Spreadsheet")
	parser.add_argument('-u', '--username', type=str, help="BuzzerBeater username", required=True, dest='username')
	parser.add_argument('-p', '--password', type=str, help="BuzzerBeater API key", required=True, dest='password')
	parser.add_argument('-pid', '--playerid', type=int, help="PlayerID", required=True, dest='playerid')
	parser.add_argument('-ws', '--worksheet', type=str, help="Worksheet name. If none if specified, this will go through all of the worksheets which will take some time", required=False, default=None, dest="ws_name")
	args = parser.parse_args()
	#start = time.time()
	#end = time.time()
	#print "Timing statistic: {}".format(end - start)

	main(args.username, args.password, args.playerid, args.ws_name)




