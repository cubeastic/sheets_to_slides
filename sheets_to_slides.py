# -*- coding: utf-8 -*-
try:
    import requests    # Import the library that allows us to contact the web service
except ImportError(requests):
    print("Please install the requested module")
import xml.etree.ElementTree as Et  # Import the function for xml handling
from sys import exit    # Import the exit function
from os import path     # Import all the os functions


# A class to import text from Google Sheets and export it to Slides with pictures as backgrounds
class SheetsToSlides:

    def __init__(self):
        self.config_file = "config.xml"
        if path.exists(self.config_file):
            self.sheet_addr = self.get_config("google_sheet")
            self.slides_addr = self.get_config("google_slide")
            self.pics_dir = self.get_config("pics_dir")
        else:
            print("ERROR: xml config file was not found")
            exit()
        self.json_format = "https://spreadsheets.google.com/feeds/cells/{0}/1/public/values?alt=json"
        self.full_addr = ""
        self.jdata = ""

    # Loop through the config file and return the requested value from the requested tag
    def get_config(self, field):
        t = Et.parse(self.config_file)
        for i in t.getroot():
            if i.tag == field:
                return i.text

    # Split the address given to multiple parts using a / as separator and return the sheet ID
    def extract_id(self):
        return self.sheet_addr.split("/")[5]

    # The sheet must be published, otherwise it cannot be accessed
    def published_check(self):
        try:
            if requests.get(self.json_format.format(self.extract_id())).status_code == 200:
                self.full_addr = self.json_format.format(self.extract_id())
                self.jdata = requests.get(self.full_addr).json()
                return True
        except ValueError:
            return False

    # Count the rows that has data and use it as a counter to loop through and collect the quotes
    def get_quotes(self):
        rows = int(self.jdata["feed"]["openSearch$totalResults"]["$t"])
        quotes = []
        for quote in range(0, rows):
            quotes.append(str(self.jdata["feed"]["entry"][quote]["content"]["$t"]))
        return quotes

    # Pull the data from the first column, from all the rows
    def sheets_phase(self):
        if self.published_check():
            return self.get_quotes()
        else:
            print("ERROR: document is not published")


a = SheetsToSlides()
print(a.sheets_phase())