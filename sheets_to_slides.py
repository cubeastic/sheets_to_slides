# -*- coding: utf-8 -*-
try:
    import requests    # Import the library that allows us to contact the web service
except ImportError(requests):
    print("Please install the requested module")


class SheetsToSlides:

    def __init__(self, address):
        self.address = address
        self.json_format = "https://spreadsheets.google.com/feeds/cells/{0}/1/public/values?alt=json"
        self.full_addr = ""

    def extract_id(self):
        return self.address.split("/")[5]

    def published_check(self):
        try:
            if requests.get(self.json_format.format(self.extract_id())).status_code == 200:
                self.full_addr = self.json_format.format(self.extract_id())
                return True
        except ValueError:
            return False

    def get_quotes(self):
        rows = int(requests.get(self.full_addr).json()["feed"]["openSearch$totalResults"]["$t"])
        quotes = []
        for quote in range(0, rows):
            quotes.append(str(requests.get(self.full_addr).json()["feed"]["entry"][quote]["content"]["$t"]))
        return quotes

    def sheets_phase(self):
        if self.published_check():
            return self.get_quotes()
        else:
            print("ERROR: document is not published")


a = SheetsToSlides("https://docs.google.com/spreadsheets/d/11_Xr3gSMAkslChwx_in3h0EJVdPHtG5m9Kv7o-WbyTc/edit#gid=0")
print(a.sheets_phase())