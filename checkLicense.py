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
	'drive_username' : 'example@example.com',
	'drive_password' : 'password',
	
	# Name of spreadsheet with license numbers
	'doc_name' : 'Example SpreadSheet',
	
	# Credentials for mailserver
	'smtp_user' : 'example@mail.com',
	'smtp_pass' : 'password',
	'smtp_server': 'smtp.gmail.com:587'
	
	# Email recipient
	'email_to' : '',

	# Email 'from' address
	'email_from' : '',
}
 
"""
Connects to Google Spreadsheets and returns an iterable list of rows
Uses global username, password and doc_name by default
"""
def get_google_data():
	# Connect to Google
	gd_client = gdata.spreadsheet.service.SpreadsheetsService()
	gd_client.email = config.drive_username 
	gd_client.password = config.drive_password
	gd_client.ProgrammaticLogin()

	q = gdata.spreadsheet.service.DocumentQuery()
	q['title'] = config.doc_name
	q['title-exact'] = 'true'
	feed = gd_client.GetSpreadsheetsFeed(query=q)

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
	currentDate = now.strftime( " %m/%d/%Y ")
	return currentDate


"""
Go the page that has the link that needs to be followed for scraping the web data from.
returns invalid if url is wrong, or license is wrong
returns inactive if Status is not Active
returns active, returns the date
"""
def check_license(licenseNumber):
	url = 'http://www7.dleg.state.mi.us/free/piresults.asp?license_number=' + licenseNumber
	stringPlus = 'http://www7.dleg.state.mi.us/free/'

	#Open the page to get the url that the data is being pulled from.
	soup = BeautifulSoup(urllib.urlopen(url).read())
	close = False
	count = 0

	#completeURL = ""

	#Gets the complete URL of the page that holds the data
	for tag in soup.find_all('a', href=True):
		count += 1
		if(count == 7):
			goToURL = tag['href']
			completeURL = stringPlus + goToURL
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
def expriring_Soon(date_Of_Expiration, current_Date):
	#Strip the dates of any unneeded characters
	current_Date = current_Date.replace('/', '')
	current_Date = current_Date.replace(' ', '')
	date_Of_Expiration = date_Of_Expiration.replace('/', '')
	date_Of_Expiration = date_Of_Expiration.replace(' ', '')

	#Make the two dates datetime objects so they can be compared
	current_Date = datetime.datetime.strptime(current_Date, '%m%d%Y')
	date_Of_Expiration = datetime.datetime.strptime(date_Of_Expiration, '%m%d%Y')

	#Compare the two days
	days = (date_Of_Expiration - current_Date).days
	return days

"""
Sends an email with the body passed to it
"""
def send_email(body):
	#Email information
	#Authentication information with Google Mail
	username = config.smtp_user
	passwd = config.smtp_pass
	fromaddr = config.email_from
	toaddrs = config.email_to

	SUBJECT = 'EMS License Experiation Test'
	Message = 'Subject: %s\n\n%s' % (SUBJECT, body)
	mailProcess = smtplib.SMTP(config.smtp_server)
	mailProcess.starttls()
	mailProcess.login(username,passwd)
	mailProcess.sendmail(fromaddr, toaddrs, Message)
	mailProcess.quit()	

def main():
	#licenseArray:  This list will hold the licenses pulled from the Google Spreadsheet.
	liceenseArray = []
	current_position = -1

	#Lists for the expiring soon licenses and the expired already licenses
	expiredList = []
	expiringSoonList = []

	#Connect to google Data and get the information from the spreadsheet
	licenesRows = get_google_data()

	#Get the current date
	currentDate = get_current_date()

	#Loop through licenses inside the Speadsheet, if they are Active in the spreadsheet skip them,  If they don't have a license listed skip them.
	for person in licenesRows:
		status = person.custom['status'].text
		license = person.custom['license'].text
		personName = person.custom['firstname'].text + ' ' + person.custom['lastname'].text
		if status != 'Active':
			continue
		if license == None:
			continue
		
		#Call check_license
		checkLicense = check_license(license)

		if checkLicense == 'inactive':
			print(checkLicense)
			info = [personName, license]
			expiredList.append(info)

		#if you get to this point it is the date that their license expires, add the person, licesnse, and date of expiration to expiringSoonList
		else:
			days = expriring_Soon(checkLicense, currentDate)
			if(days < 688):
				print('Is expiring on' + checkLicense)
				info = [personName, license, checkLicense]
				expiringSoonList.append(info)

	body = "List of expired users: \n"
	for data in expiredList:
		name = data[0]
		license = data[1]
		body += name + " with license " + license + " has expired.\n"

	body += "\nList of users expiring soon: \n"
	for data in expiringSoonList:
		name = data[0]
		license = data[1]
		date = data[2]
		body += name + " with license " + license + " is expiring on: " + date + "\n"

	send_email(body)

if __name__ == __main__:
	main()