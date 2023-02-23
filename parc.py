"""
information retrieval script testing
"""

import csv
import requests
import keyring
import pygsheets
from dateutil import parser
from bs4 import BeautifulSoup
from atlassian import Jira


APPROVERS = "customfield_10100"


class PARC:
    url = (
        "https://parc.cenic.org/cenic-parc/printpurchaseorder.cfm?purchaseordersnumber="
    )

    def __init__(self):
        self.soup = None
        self.po_elem = None

    def _get_jira(self) -> str:
        self.jira = self.soup.find(text="RT NUMBER:")
        jiraticket = self.jira.next_element.next_element.text.strip("\r\n         ")
        if jiraticket:
            if int(jiraticket) > 10000:
                return "NOC-" + jiraticket
            elif int(jiraticket) < 10000:
                return "PUR-" + jiraticket
        return "PO Does not exist"

    def _get_approval(self) -> str:
        jiraticket = self._get_jira()
        if "PUR" in jiraticket:
            jiraQuery = Jira(
                url="https://servicedesk.cenic.org/",
                username=keyring.get_password("cas", "user"),
                password=keyring.get_password("cas", "password"),
            )
            results = jiraQuery.get_issue(jiraticket, fields=[APPROVERS])

            try:
                # approval1 is Finance approval
                approval1 = parser.isoparse(
                    (results["fields"][APPROVERS][0]["completedDate"]["iso8601"])
                ).strftime("%m-%d-%Y")
            except:
                approval1 = ""

            try:
                # approval2 is Director approval
                approval2 = parser.isoparse(
                    (results["fields"][APPROVERS][1]["completedDate"]["iso8601"])
                ).strftime("%m-%d-%Y")
            except:
                approval2 = ""

            try:
                # approval3 is Core approval
                approval3 = parser.isoparse(
                    (results["fields"][APPROVERS][2]["completedDate"]["iso8601"])
                ).strftime("%m-%d-%Y")
            except:
                approval3 = ""

            return [approval1, approval2, approval3]
        else:
            return ["", "", ""]

    def _get_ordered(self) -> str:
        self.order = self.soup.find(text="Requested By")
        ordered = self.order.next_element
        if ordered:
            return ordered.text.strip("PUR-1234567890").strip(": ")
        return "Ordered By Not Found"

    def _get_segment(self) -> str:
        self.seg = self.soup.find(text="Goods/Services for:")
        segment = self.seg.next_element
        if segment:
            return segment.text.strip(" ")
        return "Segment Not Found"

    def _get_vendor(self) -> str:
        self.vend = self.soup.find(text="Vendor Information:")
        vendor = self.vend.next_element.next_element
        if vendor:
            return vendor.text.strip("\r\n         ")
        return "Vendor Not Found"

    def _get_grand_total(self) -> str:
        self.grandtot = self.soup.find(class_="grandtotalline")
        grand_total = self.grandtot.next_element
        if grand_total:
            return grand_total.text.strip(
                "\r\n               GRAND TOTAL\r\n               \r\n                  \r\n               \r\n               $     "
            )
        return "Grand Total Not Found"

    def _get_date(self) -> str:
        self.date = self.soup.find(text="DATE:")
        podate = self.date.next_element.next_element
        if podate:
            return podate.text.strip("\r\n         ")
        return "Date Not Found"

    def get_po_info(
        self,
        po: str,
        podate: bool = True,
        jiraticket: bool = True,
        ordered: bool = True,
        segment: bool = True,
        vendor: bool = True,
        grand_total: bool = True,
        status: bool = True,
        approval: bool = True,
    ) -> list:
        # Return list of data for given PO Number
        po_info = [po]

        resp = requests.get(self.url + po)
        self.soup = BeautifulSoup(resp.text, "html.parser")

        if podate:
            po_info.append(self._get_date())
        if jiraticket:
            po_info.append(self._get_jira())
        if ordered:
            po_info.append(self._get_ordered())
        if segment:
            po_info.append(self._get_segment())
        if vendor:
            po_info.append(self._get_vendor())
        if grand_total:
            po_info.append(self._get_grand_total())
        if approval:
            po_info.append("")
            po_info.append("")
            po_info.append("")
            po_info.append("")
            po_info.append("")
            po_info.append("")
            po_info.append(self._get_approval()[2])
            po_info.append(self._get_approval()[1])
            po_info.append(self._get_approval()[0])

        return po_info


def main():
    parc = PARC()

    gcpath = "/Users/acrandall/Documents/Python/OAuth/client_secret.json"
    gc = pygsheets.authorize(client_secret=gcpath, local=True)
    sh = gc.open("Inventory: New Orders Tracking")
    wk1 = sh.worksheet("index", "0")

    findPos = wk1.get_col(1, include_tailing_empty=False)
    po = str(int(findPos[-1]) + 1)

    while parc.get_po_info(po)[2] != "PO Does not exist":
        poNum = str(int(po) - 1)
        noRow = int(wk1.find(poNum, matchCase=True, matchEntireCell=True)[0].row) + 1
        wk1.update_row(noRow, parc.get_po_info(po))
        po = str(int(po) + 1)
    else:
        print("End")


if __name__ == "__main__":
    main()
