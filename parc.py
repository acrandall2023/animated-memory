"""
PARC information retrieval script testing
"""

import csv
import requests
import keyring
from dateutil import parser
from bs4 import BeautifulSoup
from atlassian import Jira


class PARC:
	url = "https://parc.cenic.org/cenic-parc/printpurchaseorder.cfm?purchaseordersnumber="

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
				url='https://servicedesk.cenic.org/',
				username = keyring.get_password("cas", "user"),
				password = keyring.get_password("cas", "password")
			)
			results = jiraQuery.get_issue(
				jiraticket, fields=["customfield_10100"]
			)

			try:
				# approval1 is Finance approval
				approval1 = (
					parser.isoparse((results["fields"]["customfield_10100"][0]["completedDate"]["iso8601"])).strftime('%m-%d-%Y')
				)
			except:
				approval1 = ("")

			try:
				#approval2 is Director approval
				approval2 = (
					parser.isoparse((results["fields"]["customfield_10100"][1]["completedDate"]["iso8601"])).strftime('%m-%d-%Y')
				)
			except:
				approval2 = ("")

			try:
				#approval3 is Core approval
				approval3 = (
					parser.isoparse((results["fields"]["customfield_10100"][2]["completedDate"]["iso8601"])).strftime('%m-%d-%Y')
				)
			except:
				approval3 = ("")

			return [approval1, approval2, approval3]
		else:
			return ["","",""]

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
			return grand_total.text.strip("\r\n               GRAND TOTAL\r\n               \r\n                  \r\n               \r\n               $     ")
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


def main(input_file: str = "pos.txt"):
	# This file should contain all PO's to be checked, outputs to csv

	parc = PARC()
	with open(input_file, "r", encoding="UTF-8") as f:
		pos = f.read().replace(", ", "\n").split()

	with open("po_result.csv", "w", encoding="UTF-8") as f:
		writer = csv.writer(f)
		for po in pos:
			writer.writerow(parc.get_po_info(po))

if __name__ == "__main__":
	main()
