""" 
    Script to pull information from PDF files and apply them to New Orders Spreadsheet
"""

from pypdf import PdfReader
import keyring
import pygsheets
from dateutil import parser
from atlassian import Jira
import os
from rich import print


APPROVERS = "customfield_10100"
JIRAURL = "https://servicedesk.cenic.org/"
GCPATH = os.path.expanduser("~") + "/Documents/Python"
POPATH = (
    os.path.expanduser("~")
    + "/Google Drive/Shared drives/CENIC Purchasing/1. Unsigned PO/21000 FY23-24"
)


def _make_results(po):
    if os.path.isfile(POPATH + "/PO " + po + ".pdf"):

        pdf = PdfReader(POPATH + "/PO " + po + ".pdf")
        page = pdf.pages[0]

        results = list(page.extract_text().split("\n"))
    else:
        results = "No PO File exists for " + po
    return results


def _get_po(results):
    purchase_order = [x for x in results if "PURCHASE ORDER" in x]
    purchase_order = str(purchase_order).strip(" [']")[-5:]
    return purchase_order


def _get_date(results):
    podate = [
        x for x in results if "La Mirada, CA 90638 16700 Valley View, Suite 168" in x
    ]
    podate = str(podate)[51:].strip("]'")
    return podate


def _get_ticket(results):
    ticket = [x for x in results if "CENIC CENIC" in x]
    ticket = str(ticket)[-9:].strip(" CENIC PUR-']")
    if int(ticket) > 10000:
        ticket = "NOC-" + ticket
        approval1 = ""
        approval2 = "NA"
        approval3 = ""
    elif int(ticket) < 10000:
        ticket = "PUR-" + ticket
        jiraQuery = Jira(
            JIRAURL,
            username=keyring.get_password("cas", "user"),
            password=keyring.get_password("cas", "password"),
        )
        approval = jiraQuery.get_issue(ticket, fields=[APPROVERS])

        try:
            # approval1 is Core/Manager approval
            approval1 = parser.isoparse(
                (approval["fields"][APPROVERS][0]["completedDate"]["iso8601"])
            ).strftime("%m-%d-%Y")
        except Exception:
            approval1 = ""
        try:
            # approval2 is Director approval
            approval2 = parser.isoparse(
                (approval["fields"][APPROVERS][1]["completedDate"]["iso8601"])
            ).strftime("%m-%d-%Y")
        except Exception:
            approval2 = ""
        try:
            # approval3 is Finance approval
            approval3 = parser.isoparse(
                (approval["fields"][APPROVERS][2]["completedDate"]["iso8601"])
            ).strftime("%m-%d-%Y")
        except Exception:
            approval3 = ""
    return ticket, approval1, approval2, approval3


def _get_requestor(results):
    requestor = [x for x in results if "Deliver To:" in x]
    requestorStart = str(requestor).rfind("To:")
    requestorEnd = str(requestor).rfind("TERMS")
    requestor = str(requestor)[requestorStart + 3 : requestorEnd].strip(" '")
    return requestor


def _get_segment(results):
    segment = [x for x in results if "Goods/Services for:" in x]
    segmentStart = str(segment).rfind("for:")
    segment = str(segment)[segmentStart + 4 : -2].strip()
    return segment


def _get_vendor(results):
    vendor = [x for x in results if "Vendor:" in x]
    vendorStart = str(vendor).rfind("Vendor:")
    vendorEnd = str(vendor).rfind("Requested By")
    vendor = str(vendor)[vendorStart + 7 : vendorEnd].strip()
    return vendor


def _get_cost(results):
    cost = [x for x in results if "GRAND TOTAL:" in x]
    costStart = str(cost).rfind("$")
    costEnd = str(cost).rfind("Disc")
    if costEnd == -1:
        costEnd = str(cost).rfind("CENIC") - 6
    cost = str(cost)[costStart + 1 : costEnd].strip(" ) ']")
    return cost


def _append(results):
    final = []
    final.append(_get_po(results))
    final.append(_get_date(results))
    ticketApproval = []
    ticketApproval = _get_ticket(results)
    final.append(ticketApproval[0])
    final.append(_get_requestor(results))
    final.append(_get_segment(results))
    final.append(_get_vendor(results))
    final.append(_get_cost(results))
    final.append("")
    final.append("")
    final.append("")
    final.append("")
    final.append(ticketApproval[1])
    final.append(ticketApproval[2])
    final.append(ticketApproval[3])
    return final


def main():
    # Credentials are assumed to be stored in the same directory and named client_secret.json
    print("[red]POPATH constant must be changed for Fiscal Year!")
    print("[green]Authorizing with GSheets")
    gc = pygsheets.authorize(credentials_directory=GCPATH, local=True)
    print("[green]Selecting New Orders spreadsheet")
    sh = gc.open_by_url(
        "https://docs.google.com/spreadsheets/d/1RNbFKIuhh3kQLjkuf0niexCstz98UNApQ4M1oih2QDI"
    )
    wk1 = sh.worksheet("index", "0")

    findPos = wk1.get_col(1, include_tailing_empty=False)
    po = str(int(findPos[-1]) + 1)

    while po.isalpha() == False:
        poNum = str(int(po) - 1)
        print(f"Checking for PO - {po}")
        noRow = int(wk1.find(poNum, matchCase=True, matchEntireCell=True)[0].row) + 1
        results = _make_results(po)
        if type(results) is list:
            # To be figured out still: How to apply formatting to added rows
            # wk1.add_rows(1)
            # wk1.apply_format()
            wk1.update_row(noRow, _append(results))
            po = str(int(po) + 1)
        else:
            po = "alpha"
    else:
        print("[green]New Orders spreadsheet updated to current[/green]")


"""
def main():
    po = "24005"
    results = _make_results(po)
    print(_append(results))
"""

if __name__ == "__main__":
    main()
