"""
This script is intended to pull information from JIRA about
IVC related tickets to export to google sheets
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

# os is used to get the directory for the credentials for pygsheets
import os

# makes prints prettier
from rich import print


# Constants to replace customfields in the export
ADDRESS = "customfield_11520"
ITEMS = "customfield_11516"
SITE_CONTACT = "customfield_11519"
PHONE = "customfield_11518"
INSURANCE = "customfield_11506"
SITE_NAME = "customfield_11507"
TRK_NUM = "customfield_11509"
ARRIVAL = "customfield_11505"
DELSPEED = "customfield_11600"
JIRAURL = "https://servicedesk.cenic.org/"
# path for credentials for pygsheets
GCPATH = os.path.expanduser("~") + "/Documents/Python"

# using keyring to log in to JIRA
jira = Jira(
    JIRAURL,
    username=keyring.get_password("cas", "user"),
    password=keyring.get_password("cas", "password"),
)


class IVCJira:
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

    def get_ivc(ivc):
        # establish ivc_info as a list
        ivc_info = []
        shipinfo = []

        # contacting jira to get information about ivc ticket provided by for loop
        results = {}
        try:
            results = jira.issue(ivc)
        except Exception:
            ivc_info.append(ivc)

        # interpret results into individual variables, if the ticket doesn't exist, go to exception
        try:
            ticket = results["key"]
            created = parser.isoparse(results["fields"]["created"]).strftime("%m-%d-%Y")
            status = results["fields"]["status"]["name"]
            requestor = results["fields"]["reporter"]["displayName"]
            requestItems = results["fields"][ITEMS]
            siteName = results["fields"][SITE_NAME].replace(" ", "_").replace(".", "")
            tracking = results["fields"][TRK_NUM]
            if (results["fields"][DELSPEED]) is None:
                arrivalDate = results["fields"][ARRIVAL]
            else:
                arrivalDate = results["fields"][DELSPEED]["value"]

        except Exception:
            ticket = "No ticket"

        # append all information to ivc_info, if the results are an empty dictionary, append no result
        if results == {}:
            ivc_info.append(ticket)
            ivc_info.append("No result")
        else:
            ivc_info.append(ticket)
            ivc_info.append(siteName)
            ivc_info.append(requestItems)
            ivc_info.append(created)
            ivc_info.append(requestor)
            ivc_info.append(arrivalDate)
            ivc_info.append(tracking)
            ivc_info.append(status)
            shipinfo.append(ticket)
            shipinfo.append(tracking)
            shipinfo.append(siteName)
        return ivc_info, shipinfo


def main():
    totalT1 = time()
    ivcship = IVCJira.get_ship()

    """
    Authenticate with GSheets using client_secret.json in ~/Documents/Python
    After authenticating, pygsheets will create a file named "sheets.googleapis.com-python.json" in ~/Documents/Python
    If an Unauthorized error occurs, delete "sheets.googleapis.com-python.json" and run this script again to re-authenticate
    """

    gc = pygsheets.authorize(credentials_directory=GCPATH, local=True)

    # Open the sheet to edit
    sh = gc.open("IVC Tracker")

    # Which sheet of the whole worksheet to edit
    wkship = sh.worksheet("title", "OutboundShipping")
    shipdate = sh.worksheet("title", "ShipDate")

    # for each item in ivcship
    print(f"[green]Total number of tickets to check:[/green] {len(ivcship)}")
    for i in ivcship:
        # Check if the value already exists, if it does update it
        if wkship.find(i, matchCase=True) != []:
            t1 = time()
            existRow = wkship.find(i, matchCase=True)[0].row
            wkship.update_row(existRow, IVCJira.get_ivc(i)[0])
            t2 = time()
            print(
                f"[green]ivcship[/green] for [red]{i}[/red] executed in {(t2-t1):.4f}s"
            )

        # If it does not exist already, create new row and add information
        else:
            t1 = time()
            wkship.append_table(IVCJira.get_ivc(i)[0], overwrite=False)
            t2 = time()
            print(
                f"[green]ivcship[/green] for [red]{i}[/red] executed in {(t2-t1):.4f}s - New line added"
            )

        if shipdate.find(i, matchCase=True) != []:
            existShip = shipdate.find(i, matchCase=True)[0].row
            shipdate.update_row(existShip, IVCJira.get_ivc(i)[1])
        else:
            shipdate.append_table(IVCJira.get_ivc(i)[1], overwrite=False)

    totalT2 = time()
    print(f"[blue]Total Time taken is[/blue] [cyan]{(totalT2-totalT1):.4f}s[/cyan]")


if __name__ == "__main__":
    main()
