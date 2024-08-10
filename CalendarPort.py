import pandas as pd
a= pd.read_excel("Current_Schedule.xlsx")
a["Instructor"]=a["Instructional Format"]+","+a["Instructor"].fillna(" ")+','+a["Meeting Patterns"].str.split('|',expand=True)[1].str.split(" - ",expand=True)[0]
a["Start Date"]=a["Meeting Patterns"].str.split('|',expand=True)[0].str.split(" - ",expand=True)[0]
a["End Date"]=a["Meeting Patterns"].str.split('|',expand=True)[0].str.split(" - ",expand=True)[1]
a["Start Time"]=a["Meeting Patterns"].str.split('|',expand=True)[2].str.split(" - ",expand=True)[0]
a["End Time"]=a["Meeting Patterns"].str.split('|',expand=True)[2].str.split(" - ",expand=True)[1]
a["Location"]=a["Meeting Patterns"].str.split('|',expand=True)[3].str.split(" - ",expand=True)[0]
#a["Class Days"]=a["Meeting Patterns"].str.split('|',expand=True)[1].str.split(" - ",expand=True)[0]
a=a.drop(columns=["Instructional Format","Delivery Mode","Section","Meeting Patterns"])
a=a.rename(columns={"Course Listing":"Subject","Instructor":"Description"})
a["All Day Event"]="False"
pd.set_option('display.max_columns', None)
print(a.head())
a.to_csv("GoogleCalendar.csv",index=False)