# -*- coding: utf-8 -*-
import httplib2     # Google Related
from apiclient import discovery     # Google Related
from oauth2client import client     # Google Related
from oauth2client import tools      # Google Related
from oauth2client.file import Storage       # Google Related
from random import randint      # Import randit to generate random Object numbers
from time import strftime       # Import strftime for the presentations name
try:
    import argparse     # Google Related
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError(argparse):
    flags = None
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
            self.pics_url = self.get_config("pics_url")
        else:
            print("ERROR: xml config file was not found")
            exit()
        self.json_format = "https://spreadsheets.google.com/feeds/cells/{0}/1/public/values?alt=json"
        self.full_addr = "" # Full address to make it easier on the eye
        self.jdata = ""     # Local Google Sheet JSON data
        self.p_id = ""      # Google Presentation ID
        self.s_id = ""      # Google Slide ID
        self.tbox_id = ""   # Text box ID
        self.session = None # Google API session ID
        self.current_time = strftime("%d.%m.%Y")    # Use as the presentations name
        self.quotes = []    # Hold the list of quotes

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
        for quote in range(0, rows):
            self.quotes.append(str(self.jdata["feed"]["entry"][quote]["content"]["$t"]))
        return True if len(self.quotes) > 1 else False

    # Pull the data from the first column, from all the rows
    def sheets_phase(self):
        if self.published_check():
            return self.get_quotes()
        else:
            print("ERROR: document is not published")

    def get_credentials(self):
        SCOPES = 'https://www.googleapis.com/auth/presentations'
        CLIENT_SECRET_FILE = 'client_secret.json'
        APPLICATION_NAME = 'Google Sheets to Slides'
        credential_path = "credentials/cred.json"
        store = Storage(credential_path)
        creds = store.get()
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            creds = tools.run_flow(flow, store, flags)
        return creds

    def create_pst(self):
        body = {
            'title': self.current_time
        }
        presentation = self.session.presentations().create(body=body).execute()
        self.p_id = presentation.get('presentationId')
        return True if self.p_id else False

    def create_slide(self):
        requests = [
            {
                'createSlide': {
                    'insertionIndex': '0'
                }
            }
        ]
        body = {
            'requests': requests
        }
        response = self.session.presentations().batchUpdate(presentationId=self.p_id, body=body).execute()
        self.s_id = response.get('replies')[0].get('createSlide').get('objectId')
        return True if self.s_id else False

    def create_text_box(self, q):
        element_id = "textbox" + str(randint(1, 100))
        pt350 = {
            'magnitude': 350,
            'unit': 'PT'
        }
        requests = [
            {
                'createShape': {
                    'objectId': element_id,
                    'shapeType': 'TEXT_BOX',
                    'elementProperties': {
                        'pageObjectId': self.s_id,
                        'size': {
                            'height': pt350,
                            'width': pt350
                        },
                        'transform': {
                            'scaleX': 1,
                            'scaleY': 1,
                            'translateX': 350,
                            'translateY': 100,
                            'unit': 'PT'
                        }
                    }
                }
            },
            {
                'insertText': {
                    'objectId': element_id,
                    'insertionIndex': 0,
                    'text': q
                }
            }
        ]
        body = {
            'requests': requests
        }
        response = self.session.presentations().batchUpdate(presentationId=self.p_id, body=body).execute()
        self.tbox_id = response.get('replies')[0].get('createShape').get('objectId')
        return True if self.tbox_id else False

    def change_background(self):
        requests = [
            {
                'updatePageProperties':
                    {
                        "objectId": self.s_id,
                        "pageProperties": {
                            "pageBackgroundFill": {
                                "stretchedPictureFill": {
                                    "contentUrl": self.pics_url
                                }
                            }
                        },
                        "fields": "pageBackgroundFill"
                    }
            }
        ]
        body = {
            'requests': requests
        }
        self.session.presentations().batchUpdate(presentationId=self.p_id, body=body).execute()

    def style_text_box(self):
        requests = [
            {
                'updateTextStyle': {
                    'objectId': self.tbox_id,
                    'style': {
                        'fontFamily': 'Arial',
                        'fontSize': {
                            'magnitude': 48,
                            'unit': 'PT'
                        },
                        'foregroundColor': {
                            'opaqueColor': {
                                'rgbColor': {
                                    'blue': 1.0,
                                    'green': 0.0,
                                    'red': 0.0
                                }
                            }
                        }
                    },
                    'fields': 'foregroundColor,fontFamily,fontSize'
                }
            }
        ]
        body = {
            'requests': requests
        }
        self.session.presentations().batchUpdate(presentationId=self.p_id, body=body).execute()

    def slides_phase(self):
        # Start by pulling the data from the sheet file
        if self.sheets_phase():
            # Start talking with Google
            credentials = self.get_credentials()
            http = credentials.authorize(httplib2.Http())
            self.session = discovery.build('slides', 'v1', http=http)
            # Create the presentation
            if self.create_pst():
                for q in self.quotes:
                    # Create the slide and make the magic happen
                    if self.create_slide():
                        if self.create_text_box(q):
                            self.style_text_box()
                        self.change_background()


a = SheetsToSlides()
a.slides_phase()