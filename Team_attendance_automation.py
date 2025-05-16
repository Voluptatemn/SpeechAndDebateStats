from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd
import os
import numpy as np
from team_roster_infrastructure import Spreadsheet
from team_roster_infrastructure import Spreadsheet2
import time
from tabulate import tabulate

os.remove("token_team.json")
attendance_backend = Spreadsheet('1w4AxeQtuxsbcSxiC7lstVDho5TTZzgJuTERZTFE-Fe4', "backend","token_team.json", 
                                 directory = "C:/Users/rebec/OneDrive/Desktop/Speech and Debate/Team Code")
team_roster = Spreadsheet2('https://docs.google.com/spreadsheets/d/1KTmWPY-huDOmhLX0pPkJp9TyhZR0U5tzn-dg9-3TbGk/edit?pli=1#gid=1380927623', 
                           "WI23 Attendance","token_team.json")

# team_roster.dataframe["Excused Absences"] = 0
# for i in team_roster.dataframe.index:
#     attendance = 0p
#     excused = 0
#     for col in ["9/27/2021","9/30/2021","10/4/2021","10/7/2021","10/11/2021","10/14/2021"]:
#         if (team_roster.dataframe.loc[i, col] == "1") | (team_roster.dataframe.loc[i, col] == "E"):
#             attendance += 1
#         if team_roster.dataframe.loc[i, col] == "E":
#             excused += 1
#     print(excused)

#     team_roster.dataframe.loc[i, "GBM Attendance"] = attendance 
#     team_roster.dataframe.loc[i, "Excused Absences"] = excused 



backend = attendance_backend.dataframe

need_check = backend[backend["Run Check"] == "0"]

attendance_backend.dataframe["Run Check"] = 1
attendance_backend.update_spreadsheet()

for i in need_check.index:

    id = need_check.loc[i, "Responses ID"]

    date = need_check.loc[i, "Date"]

    team_roster.dataframe[date] = 0

    print("part 1")

    # get responses from form
    responses = Spreadsheet(id,"Form Responses 1", token_name= "token_team.json").dataframe.iloc[:,[0,1,2]]
    responses.columns = ["Timestamp","Email","Are you a new member?"]


    print("part 2")

    # drop unncessary columns
    responses = responses.drop(labels = ["Timestamp", "Are you a new member?"], axis = 1)

    # add 1 to GBM attendance if the team roster email is in the responses email column
    print(team_roster.dataframe.columns)
    team_roster.dataframe.loc[team_roster.dataframe.Email.isin(responses.Email),date] = 1
    # get members from responses that are not in team roster and put them in object
    new_members = responses[~responses.Email.isin(team_roster.dataframe.Email)]

    print("part 3")

    # set new members GBM to 1
    new_members[date] = 1

    # concat team roster dataframe with new memebrs dataframe
    team_roster.dataframe = pd.concat([team_roster.dataframe,new_members])

    # update team roster

    print("part 4")
    team_roster.update_spreadsheet()

# practice attendance
