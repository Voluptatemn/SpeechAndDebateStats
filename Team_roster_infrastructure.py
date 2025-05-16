from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd
import numpy as np
import regex as re

Directory = "C:/Users/rebec/OneDrive/Desktop/Speech and Debate/Team Code"
# The ID and range of a sample spreadsheet.

def get_sheet_values(SPREADSHEET_ID,RANGE_NAME, SCOPES, token_name, directory = Directory):
  
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    try:
        os.chdir(directory)
    except:
        directory = Directory

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token_name):
        creds = Credentials.from_authorized_user_file(token_name, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'team_creds.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_name, 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=RANGE_NAME).execute()

    return result

def convert_to_sheets(dataframe, spreadsheetID, rangeName, creds = None):
    '''Makes the column names of a dataframe the first row of the dataframe. Then uploads dataframe to Google Sheets'''
    first_row = pd.DataFrame(data = [dataframe.columns], columns = dataframe.columns)
    df = pd.concat([first_row, dataframe])
    
    spreadsheet_id = '1KTmWPY-huDOmhLX0pPkJp9TyhZR0U5tzn-dg9-3TbGk'
    range_name = 'WI23 Attendance'
    
    values = df.values.tolist()

    body = {
        'values': values
    }
    
    service = build('sheets', 'v4', credentials=creds)
    result = service.spreadsheets().values().update(
    spreadsheetId=spreadsheetID, range=rangeName,
    valueInputOption='RAW', body=body).execute()


def convert_to_pandas(values):
    
    df = pd.DataFrame(values, columns = values[0])
    df = df.drop(0)
    df = df.fillna(value = "")
    
    return df

def intersection(lst1, lst2):
    return list(set(lst1) & set(lst2))


def get_spreadsheet(spreadsheet_id, range_name, token_name, directory = Directory):
    
    try:
        os.chdir(directory)
    except:
        directory = Directory

    import os.path
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    import os
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    os.chdir(directory)
    SCOPES = ['https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/script.scriptapp', 'https://www.googleapis.com/auth/drive',"https://www.googleapis.com/auth/spreadsheets"]
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(token_name):
        creds = Credentials.from_authorized_user_file(token_name, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'team_creds.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_name, 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=spreadsheet_id,
                                range=range_name).execute()
    values = result.get('values', [])

    return result

def get_spreadsheet_df(result):

    values =  result.get('values', [])

    df = pd.DataFrame(values, columns = values[0])
    df = df.drop(0)
    df = df.fillna(value = "")
    
        
    return df

def get_id(url):
    """
    gets id from google workplace url
    """

    str1 = re.sub("https://docs.google.com/spreadsheets/d/","",url)
    str2 = re.sub("/edit.+","",str1)

    return str2


class Spreadsheet2:
    def __init__(self, url, range_name, token_name, directory = Directory):

        try:
            os.chdir(directory)
        except:
            directory = Directory

        self.spreadsheet_id = get_id(url)
        self.range_name = range_name
        self.spreadsheet = get_spreadsheet(self.spreadsheet_id, range_name, token_name)
        self.dataframe = get_spreadsheet_df(self.spreadsheet)
        self.token_name = token_name
        self.directory = directory



    def update_spreadsheet(self):

        """
        Updates the Spreadsheet range with self.dataframe

        """

        import os.path
        from googleapiclient.discovery import build
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials

        import os
        """Shows basic usage of the Sheets API.
        Prints values from a sample spreadsheet.
        """

        token_name = self.token_name
        SCOPES = ['https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/script.scriptapp', 'https://www.googleapis.com/auth/drive',"https://www.googleapis.com/auth/spreadsheets"]
        os.chdir(self.directory)
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(token_name):
            creds = Credentials.from_authorized_user_file(token_name, SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'team_creds.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(token_name, 'w') as token:
                token.write(creds.to_json())

        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        new_frame = self.dataframe

        first_row = pd.DataFrame(data = [new_frame.columns], columns = new_frame.columns)
        df = pd.concat([first_row, new_frame])

        df = df.fillna("")
    
        values = df.values.tolist()

        body1 = {
            'values': values
        }
        result = service.spreadsheets().values().update(spreadsheetId=self.spreadsheet_id, range = self.range_name, valueInputOption = 'RAW', body = body1).execute()


class Spreadsheet:
    def __init__(self, spreadsheet_id, range_name, token_name, directory = Directory):

            
        try:
            os.chdir(directory)
        except:
            directory = Directory

        self.spreadsheet_id = spreadsheet_id
        self.range_name = range_name
        self.spreadsheet = get_spreadsheet(spreadsheet_id, range_name, token_name)
        self.dataframe = get_spreadsheet_df(self.spreadsheet)
        self.token_name = token_name
        self.directory = directory



    def update_spreadsheet(self):

        """
        Updates the Spreadsheet range with self.dataframe

        """

        import os.path
        from googleapiclient.discovery import build
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials

        import os
        """Shows basic usage of the Sheets API.
        Prints values from a sample spreadsheet.
        """

        token_name = self.token_name
        SCOPES = ['https://www.googleapis.com/auth/script.projects', 'https://www.googleapis.com/auth/script.scriptapp', 'https://www.googleapis.com/auth/drive',"https://www.googleapis.com/auth/spreadsheets"]
        os.chdir(self.directory)
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(token_name):
            creds = Credentials.from_authorized_user_file(token_name, SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(token_name, 'w') as token:
                token.write(creds.to_json())

        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        new_frame = self.dataframe

        first_row = pd.DataFrame(data = [new_frame.columns], columns = new_frame.columns)
        df = pd.concat([first_row, new_frame])

        df = df.fillna("")
    
        values = df.values.tolist()

        body1 = {
            'values': values
        }
        result = service.spreadsheets().values().update(spreadsheetId=self.spreadsheet_id, range = self.range_name, valueInputOption = 'RAW', body = body1).execute()

      


def fix_findall(lst):
    while True:
        if ([] in lst) == True:
            lst.remove([])
        else:
            return [item[0] for item in lst]

#Below code is to test if you can access and edit the spreadsheet properly
#It should print out "Numbers: and 10-60" in the spreadsheet whose url is below

#Temp = Spreadsheet2('https://docs.google.com/spreadsheets/d/1OjByk7d7b6I-2Rin3f0XKZTR1NsQNZAdUJ9g4Q8sqbE/edit#gid=735963550','Complete Team Roster', 'token_team', Directory)
#print(Temp.dataframe)

# initialize list elements
#data = [10,20,30,40,50,60]
  
# Create the pandas DataFrame with column name is provided explicitly
#df = pd.DataFrame(data, columns=['Numbers'])

#Temp.dataframe = df
#Temp.update_spreadsheet()
