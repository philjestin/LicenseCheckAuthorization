#Philip Middleton
#Python script used for checking license numbers on a state website to see if they have expired or not.

#Import Modules
from bs4 import BeautifulSoup
import urllib
import datetime
from datetime import date
import gdata.docs
import gdata.docs
import gdata.docs.service
import gdata.spreadsheet.service
import re, os
import smtplib

config = {
	# Credentials for google drive acct with access to spreadsheet
	'drive_username' : '',
	'drive_password' : '',
	
	# Name of spreadsheet with license numbers
	'doc_name' : 'Current Michigan Tech EMS Roster',
	# Name of the column holding the number
	'doc_column' : 'License Number',
	
	# Credentials for mailserver
	'smtp_user' : '',
	'smtp_pass' : '',
	'smtp_server': 'smtp.gmail.com:587',
	
	# Email recipient
	'email_to' : '',

	# Email 'from' address
	'email_from' : '',

	# If person is expiring in less than this, they are expiring soon
	'days' : 690
}
 
"""
Connects to Google Spreadsheets and returns an iterable list of rows
Uses global username, password and doc_name by default
"""
def get_google_data():
	global config
	# Connect to Google
	gd_client = gdata.spreadsheet.service.SpreadsheetsService()
	gd_client.email = config['drive_username']
	gd_client.password = config['drive_password']
	gd_client.ProgrammaticLogin()

	q = gdata.spreadsheet.service.DocumentQuery()
	q['title'] = config['doc_name']
	q['title-exact'] = 'true'
	feed = gd_client.GetSpreadsheetsFeed(query=q)

	# If spreadsheet not found, exit
	if feed.entry == []:
		print("Spreadsheet not found. Exiting")
		exit(1)

	spreadsheet_id = feed.entry[0].id.text.rsplit('/',1)[1]
	feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
	worksheet_id = feed.entry[0].id.text.rsplit('/',1)[1]
	
	# Get rows from the spreadsheet
	rows = gd_client.GetListFeed(spreadsheet_id, worksheet_id).entry
	return rows

"""
Gets the current date
"""
def get_current_date():
	now = datetime.datetime.now()
	current_date = now.strftime( " %m/%d/%Y ")
	return current_date


"""
Go the page that has the link that needs to be followed for scraping the web data from.
returns invalid if url is wrong, or license is wrong
returns inactive if Status is not Active
returns active, returns the date
"""
def check_license(licenseNumber):
	url = 'http://www7.dleg.state.mi.us/free/piresults.asp?license_number=' + licenseNumber
	string_plus = 'http://www7.dleg.state.mi.us/free/'

	#Open the page to get the url that the data is being pulled from.
	soup = BeautifulSoup(urllib.urlopen(url).read())
	close = False
	count = 0

	#Gets the complete URL of the page that holds the data
	for tag in soup.find_all('a', href=True):
		count += 1
		if(count == 7):
			go_to_url = tag['href']
			completeURL = string_plus + go_to_url

	if(completeURL == ""):
		return "invalid"

	#Go to the complete URL that the data is pulled from
	soup = BeautifulSoup(urllib.urlopen(completeURL).read())

	#Reset the conditions
	close = False
	count = 0

	#Get the Name of the person in which the License Number belongs to
	for td in soup.find_all('td'):
		if(td.string != None and td.string.strip() == "Name and Address"):
			close = True
		if(close):
			count += 1
		if(count>2):
			print('Name: ' + td.string)
			break

	#Reset the conditions
	count = 0
	close = False

	#Gets the date that the license expires
	for td in soup.find_all('td'):
		if(td.string != None and td.string.strip() == "Active"):
			close = True
		if(close):
			count += 1
		if(count>2):
			dateOfExpiration = td.string
			return dateOfExpiration

    #If the user is not Active, then the count will not be greater than 0, therefore skip that person so it doesn't break the script when comparing the next date.       
	# If count is zero, they are inactive
	if(count==0):
		return "inactive"

"""
Compare the dates
"""
def expriring_Soon(date_of_expiration, current_date):
	#Strip the dates of any unneeded characters
	current_date = current_date.replace('/', '').replace(' ', '')
	date_of_expiration = date_of_expiration.replace('/', '').replace(' ', '')

	#Make the two dates datetime objects so they can be compared
	current_date = datetime.datetime.strptime(current_date, '%m%d%Y')
	date_of_expiration = datetime.datetime.strptime(date_of_expiration, '%m%d%Y')

	#Compare the two days
	days = (date_of_expiration - current_date).days
	return days

"""
Sends an email with the body that was passed to it
"""
def send_email(body):
	global config
	#Email information
	#Authentication information with Google Mail
	username = config['smtp_user']
	passwd = config['smtp_pass']
	fromaddr = config['email_from']
	toaddrs = config['email_to']

	subject = 'EMS License Experiation Test'
	message = 'Subject: %s\n\n%s' % (subject, body)
	mailProcess = smtplib.SMTP(config['smtp_server'])
	mailProcess.starttls()
	mailProcess.login(username,passwd)
	mailProcess.sendmail(fromaddr, toaddrs, message)
	mailProcess.quit()	

def main():
	global config

	#Lists for the expiring soon licenses and the expired already licenses
	expired_list = []
	expiring_soon_list = []

	#Connect to google Data and get the information from the spreadsheet
	license_rows = get_google_data()

	#Get the current date
	current_date = get_current_date()

	#Loop through licenses inside the Speadsheet
	column_name = config['doc_column'].lower().replace(' ','')
	for person in license_rows:
		
		# Temporarily removing sheet 'status' functionality - may not be needed
		#status = person.custom['status'].text
		status='Active'

		license = person.custom[column_name].text
		person_name = person.custom['firstname'].text + ' ' + person.custom['lastname'].text

		# If they are Active in the spreadsheet skip them
		# If they don't have a license listed skip them
		if status != 'Active':
			continue
		if license == None:
			continue
		
		# Check their license status online
		license_status = check_license(license)

		if license_status == 'inactive':
			print("%s\n" % license_status)
			info = [person_name, license]
			expired_list.append(info)

		#if you get to this point license_status is the date that their license expires
		else:
			days = expriring_Soon(license_status, current_date)
			if(days < config['days']):
				# add the person, licesnse, and date of expiration to expiring_soon_list
				print('Is expiring on %s\n' % license_status)
				info = [person_name, license, license_status]
				expiring_soon_list.append(info)
			else:
				print("Active\n")

	body = "List of expired members: \n"
	for data in expired_list:
		name = data[0]
		license = data[1]
		body += name + " with license " + license + " has expired.\n"

	body += "\nList of members expiring soon: \n"
	for data in expiring_soon_list:
		name = data[0]
		license = data[1]
		date = data[2]
		body += name + " with license " + license + " is expiring on: " + date + "\n"

	send_email(body)

if __name__ == '__main__':
	main()
