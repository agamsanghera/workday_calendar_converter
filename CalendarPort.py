'''
Author: Agam Sanghera
username: agamsanghera
'''
import pandas as pd
#read excel as pandas dataframe
a= pd.read_excel("Current_Schedule.xlsx")
#convert columns into the format uploadable on google calendar
a["Instructor"]=a["Instructional Format"]+","+a["Instructor"].fillna(" ")+','+a["Meeting Patterns"].str.split('|',expand=True)[1].str.split(" - ",expand=True)[0]
a["Start Date"]=a["Meeting Patterns"].str.split('|',expand=True)[0].str.split(" - ",expand=True)[0]
a["End Date"]=a["Meeting Patterns"].str.split('|',expand=True)[0].str.split(" - ",expand=True)[1]
a["Start Time"]=a["Meeting Patterns"].str.split('|',expand=True)[2].str.split(" - ",expand=True)[0]
a["End Time"]=a["Meeting Patterns"].str.split('|',expand=True)[2].str.split(" - ",expand=True)[1]
a["Location"]=a["Meeting Patterns"].str.split('|',expand=True)[3].str.split(" - ",expand=True)[0]
#a["Class Days"]=a["Meeting Patterns"].str.split('|',expand=True)[1].str.split(" - ",expand=True)[0]
a=a.drop(columns=["Instructional Format","Delivery Mode","Section","Meeting Patterns"])
#rename columns as expected input for google calendar
a=a.rename(columns={"Course Listing":"Subject","Instructor":"Description"})
#make the classes show as events in the specified timeframe
a["All Day Event"]="False"
#display all the columns in the head
pd.set_option('display.max_columns', None)
print(a.head())
#save as CSV which can be uploaded to google calendar
a.to_csv("GoogleCalendar.csv",index=False)