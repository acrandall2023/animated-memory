'''
This script is intended to pull information from JIRA about IVC related tickets to export to google sheets
'''

# Keyring is used for JIRA authentication
import keyring
# json is used to export information to json file with pretty printed text
import json
# CSV is used to export to a csv file
import csv
# parser from dateutil to convert ISO8601 date to MM-DD-YYYY
from dateutil import parser


# using the Jira module from Atlassian to pull the information from JIRA
from atlassian import Jira

# Constants to replace customfields in the export
ADDRESS = "customfield_11520"
ITEMS = "customfield_11516"
SITE_CONTACT = "customfield_11519"
PHONE = "customfield_11518"
INSURANCE = "customfield_11506"
SITE_NAME = "customfield_11507"
TRK_NUM = "customfield_11509"
ARRIVAL = "customfield_11505"

class IVCJira:

	def get_ivc(ivc):

#		using keyring to log in to JIRA
		jira = Jira(
			url='https://servicedesk.cenic.org/',
			username = keyring.get_password("cas", "user"),
			password = keyring.get_password("cas", "password")
			)

#		establish ivc_info as a list
		ivc_info = []

#		contacting jira to get information about ivc ticket provided from ivc.txt
		results = {}
		try:
			results = jira.issue(ivc)
		except:
			ivc_info.append(ivc)

#		interpret results into individual variables, if the ticket doesn't exist, go to exception
		try:
			ticket = results["key"]
			issue = results["fields"]["issuetype"]["name"]
			created = parser.isoparse(results["fields"]["created"]).strftime('%m-%d-%Y')
			status = results["fields"]["status"]["name"]
			requestor = results["fields"]["reporter"]["displayName"]
			contact = results["fields"][SITE_CONTACT]
			phone = results["fields"][PHONE]
			requestItems = results["fields"][ITEMS]
			insurance = str(results["fields"][INSURANCE])
			siteName = results["fields"][SITE_NAME]
			tracking = results["fields"][TRK_NUM]
			arrivalDate = results["fields"][ARRIVAL]
		except:
			ticket = "No ticket"

#		append all information to ivc_info, if the results are empty, no result
		if results == {}:
			ivc_info.append(ticket)
			ivc_info.append("No result")
		else:
			ivc_info.append(ticket)
			ivc_info.append(issue)
			ivc_info.append(created)
			ivc_info.append(requestor)
			ivc_info.append(requestItems)
			ivc_info.append(siteName)
			ivc_info.append(contact)
			ivc_info.append(phone)
			ivc_info.append(insurance)
			ivc_info.append(arrivalDate)
			ivc_info.append(tracking)
			ivc_info.append(status)
		return ivc_info


def main():

#	ivc.txt contains the tickets being checked
	with open("ivc.txt", "r", encoding="UTF-8") as f:
		ivcs = f.read().replace(", ", "\n").split()

#	results are sent to ivc_result.csv
	with open("ivc_result.csv", "w", encoding="UTF-8") as f:
		writer = csv.writer(f)
		for ivc in ivcs:
			writer.writerow(IVCJira.get_ivc(ivc))

if __name__ == "__main__":
	main()
