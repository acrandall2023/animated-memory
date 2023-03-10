"""
This script is intended to pull information from JIRA about IVC related tickets to export to google sheets
"""

# Keyring is used for JIRA authentication
import keyring

# parser from dateutil to convert ISO8601 date to MM-DD-YYYY
from dateutil import parser

# pygsheets used to interface with Google Sheets
import pygsheets

# using the Jira module from Atlassian to pull the information from JIRA
from atlassian import Jira

# For timing decorator
from time import time

# Constants to replace customfields in the export
ADDRESS = "customfield_11520"
ITEMS = "customfield_11516"
SITE_CONTACT = "customfield_11519"
PHONE = "customfield_11518"
INSURANCE = "customfield_11506"
SITE_NAME = "customfield_11507"
TRK_NUM = "customfield_11509"
ARRIVAL = "customfield_11505"

# using keyring to log in to JIRA
jira = Jira(
    url="https://servicedesk.cenic.org/",
    username=keyring.get_password("cas", "user"),
    password=keyring.get_password("cas", "password"),
)


def timer_func(func):
    # This function shows the execution time of
    # the function object passed
    def wrap_func(*args, **kwargs):
        t1 = time()
        result = func(*args, **kwargs)
        t2 = time()
        print(f"Function {func.__name__!r} executed in {(t2-t1):.4f}s")
        return result

    return wrap_func


class IVCJira:
    @timer_func
    def get_ship():
        # The query to find open tickets
        _jql = 'project = "Inventory Control" AND issuetype = "Outbound Shipping" AND status not in (Deleted, New) AND updated > endOfDay(-4) ORDER BY createdDate ASC'

        # using the jql function with the _jql query
        queryJira = jira.jql(_jql)

        # establish ivc as a blank list
        ivc = []

        # add each ticket to ivc
        for i in range(len(queryJira["issues"])):
            ivc += queryJira["issues"][i]["key"].split()
        return ivc

    @timer_func
    def get_req():
        _jql = 'project = "Inventory Control" AND issuetype = "Equipment Request" AND status not in (Deleted, New) AND updated > endOfDay(-4) ORDER BY createdDate asc'
        queryJira = jira.jql(_jql)
        ivc = []
        for i in range(len(queryJira["issues"])):
            ivc += queryJira["issues"][i]["key"].split()
        return ivc

    @timer_func
    def get_ret():
        _jql = 'project = "Inventory Control" AND issuetype = "Equipment Returns" AND status not in (Deleted, New) AND updated > endOfDay(-4) ORDER BY createdDate asc'
        queryJira = jira.jql(_jql)
        ivc = []
        for i in range(len(queryJira["issues"])):
            ivc += queryJira["issues"][i]["key"].split()
        return ivc

    @timer_func
    def get_ivc(ivc):
        # establish ivc_info as a list
        ivc_info = []

        # contacting jira to get information about ivc ticket provided from ivc.txt
        results = {}
        try:
            results = jira.issue(ivc)
        except:
            ivc_info.append(ivc)

        # interpret results into individual variables, if the ticket doesn't exist, go to exception
        try:
            ticket = results["key"]
            issue = results["fields"]["issuetype"]["name"]
            created = parser.isoparse(results["fields"]["created"]).strftime("%m-%d-%Y")
            status = results["fields"]["status"]["name"]
            requestor = results["fields"]["reporter"]["displayName"]
            # contact = results["fields"][SITE_CONTACT]
            # phone = results["fields"][PHONE]
            requestItems = results["fields"][ITEMS]
            # insurance = str(results["fields"][INSURANCE])
            siteName = results["fields"][SITE_NAME]
            tracking = results["fields"][TRK_NUM]
            arrivalDate = results["fields"][ARRIVAL]
        except:
            ticket = "No ticket"

        # append all information to ivc_info, if the results are an empty dictionary, append no result
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
            # ivc_info.append(contact)
            # ivc_info.append(phone)
            # ivc_info.append(insurance)
            ivc_info.append(arrivalDate)
            ivc_info.append(tracking)
            ivc_info.append(status)
        return ivc_info


@timer_func
def main():
    ivcship = IVCJira.get_ship()
    ivcreq = IVCJira.get_req()
    ivcret = IVCJira.get_ret()

    # Authenticate with GSheets using client_secret.json in current directory
    gc = pygsheets.authorize(local=True)

    # Open the sheet to edit
    sh = gc.open("IVC Tracker")

    # Which sheet of the whole worksheet to edit
    wkship = sh.worksheet("title", "OutboundShipping")
    wkreq = sh.worksheet("title", "EquipmentRequest")
    wkret = sh.worksheet("title", "EquipmentReturns")

    # for each item in ivcship
    for i in ivcship:
        # Check if the value already exists, if it does update it
        if wkship.find(i, matchCase=True) != []:
            existRow = wkship.find(i, matchCase=True)[0].row
            wkship.update_row(existRow, IVCJira.get_ivc(i))
        # If it does not exist already, create new row and add information
        else:
            wkship.append_table(IVCJira.get_ivc(i), overwrite=False)

    for i in ivcreq:
        if wkreq.find(i, matchCase=True) != []:
            existRow = wkreq.find(i, matchCase=True)[0].row
            wkreq.update_row(existRow, IVCJira.get_ivc(i))
        else:
            wkreq.append_table(IVCJira.get_ivc(i), overwrite=False)

    for i in ivcret:
        if wkret.find(i, matchCase=True) != []:
            existRow = wkret.find(i, matchCase=True)[0].row
            wkret.update_row(existRow, IVCJira.get_ivc(i))
        else:
            wkret.append_table(IVCJira.get_ivc(i), overwrite=False)
    print("Done. IVC Tracker has been updated with latest information")


if __name__ == "__main__":
    main()
