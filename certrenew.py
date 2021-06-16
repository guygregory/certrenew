# Certification Renewal Report for Partners
# Guy Gregory - guy.gregory@microsoft.com

import pandas as pd
import plotly.express as px
import plotly
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Specify the location of the file which can be downloaded from Partner Center Insights:
# https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions


root = Tk()
root.withdraw()
root.filename = filedialog.askopenfilename(title="Please select the Trainings report from Partner Center Insights", filetypes=(("CSV files", "*.csv"),("TSV files", "*.tsv"), ("All files", "*.*")) )
inputfile = root.filename

# Load the dataframe from the .csv or .tsv file
name, ext = os.path.splitext(inputfile)

if ext.lower() == ".csv":
    dfl = pd.read_csv(inputfile, header=0)
    
elif ext.lower() == ".tsv":
    dfl = pd.read_csv(inputfile, sep='\t', header=0)
    
else:
    messagebox.showerror(message="Not a valid file type, please use Partner Skills Report - Trainings in either .csv or .tsv format", title="Certification Renewal Tool")
    sys.exit()


# Check to see whether this report includes the new fields, and exit gracefully if not

if 'ExpirationDate' not in dfl.columns:
    messagebox.showerror(message="Some of the newer fields are missing from this report. The likely cause is that the report is older than June 17th 2021. Please download a new version of this report from Partner Center, and retry.", title="Certification Renewal Tool")
    sys.exit()

# Check to see whether this report is from a Report Viewer (limited) or Executive Report Viewer (includes PII)

if 'IndividualFirstName' in dfl.columns:
    ExecutiveReportViewer = True
    
elif 'AADUserId' in dfl.columns:
    ExecutiveReportViewer = False
    messagebox.showinfo(message="This report has been obtained by a user with the Report Viewer role. To view Personal (PII) information, please download the report using a member of the Executive Report Viewer role. This program will now continue, but with limited data.", title="Certification Renewal Tool")
    
else:
    messagebox.showerror(message="Unexpected column headings found, please use Partner Skills Report with unedited column headings", title="Certification Renewal Tool")
    sys.exit()



# Before we remove a bunch of columns, capture the partner name and MPN ID

MPNId = dfl.loc[0].at['MPNId']
PartnerName = dfl.loc[0].at['PartnerName']
PartnerCityLocation = dfl.loc[0].at['PartnerCityLocation']
PartnerCountryLocation = dfl.loc[0].at['PartnerCountryLocation']

# Create a simplified version for the filename
PartnerNameAlphaNumericOnly = ''.join(e for e in PartnerName if e.isalnum())

# Remove unwanted columns
trainingstable = dfl.drop(columns=["Month", "IcMCP", "MCPID", "MPNId", "PartnerName","PartnerCityLocation","PartnerCountryLocation"])

# Keep only training activities that are Certifications
certtable = trainingstable[trainingstable.TrainingType.eq("Certification")]

# Remove the Training Type column
certtable = certtable.drop(columns=["TrainingType"])

# TrainingActivityIds for Specialty Exams
#3106  Azure for SAP Workloads
#3131  Azure IoT Developer

# Keep the Specialty Exams, and anything containing the words 'Associate' or 'Expert'
rbstable = certtable[certtable.TrainingActivityId.eq("3106") | certtable.TrainingActivityId.eq("3131") | certtable.TrainingTitle.str.contains("Expert|Associate")].copy()

# Drop anything with a TrainingActivityId of < 3050
rbstable = rbstable.drop(rbstable[rbstable.TrainingActivityId.astype(int) < 3050].index)

# Important - Pandas stores dates in datetime format by default, so to display only the date element, then use dt.date 
# for example: df['just_date'] = df['dates'].dt.date

rbstable["TrainingCompletionDateTime"] = pd.to_datetime(rbstable["TrainingCompletionDate"])
rbstable["ExpirationDateTime"] = pd.to_datetime(rbstable["ExpirationDate"])

# Convert value to date format
rbstable["TrainingCompletionDateOnly"] = pd.to_datetime(rbstable["TrainingCompletionDateTime"]).dt.date

rbstable["CertRenewalDeadline"] = pd.to_datetime(rbstable["ExpirationDateTime"]).dt.date

#Create a new column for the start of the cert renewal window, and set the Cert Renewal Window Opens values to 6 months before the Exam Expiry Date
rbstable["CertRenewalWindowOpens"] = rbstable["CertRenewalDeadline"] + relativedelta(months=-6)

#Clean-up unused columns
rbstable = rbstable.drop(columns=["TrainingCompletionDate","TrainingCompletionDateOnly"])

#Re-order the columns into the desired order
if ExecutiveReportViewer == True:
    rbstable = rbstable[["TrainingActivityId", "TrainingTitle", "IndividualFirstName", "IndividualLastName", "Email", "CorpEmail", "TrainingCompletionDateTime", "CertRenewalWindowOpens", "CertRenewalDeadline"]]

else:
    rbstable = rbstable[["TrainingActivityId", "TrainingTitle", "TrainingCompletionDateTime", "CertRenewalWindowOpens", "CertRenewalDeadline"]]

df = rbstable # Create a copy of the dataframe


# replacing null values in dataframe with a blank space to avoid %{customdata[1]}

if ExecutiveReportViewer == True:
    df["IndividualFirstName"].fillna(" ", inplace = True)
    df["IndividualLastName"].fillna(" ", inplace = True) 
    df["CorpEmail"].fillna(" ", inplace = True) 

# Sort by date, and then reindex
df = df.sort_values(by='CertRenewalWindowOpens')
df = df.reset_index()

# Assign Columns to variables
tasks = df.index
start = df['CertRenewalWindowOpens']
finish = df['CertRenewalDeadline']

# Create Gantt Chart

if ExecutiveReportViewer == True:
    fig = px.timeline(df, x_start=start, x_end=finish, y=tasks, title='Microsoft Certification Renewal Insights - ' + PartnerName + " (" + str(MPNId) + ")", 
    
    hover_name=df['TrainingTitle'], 
    hover_data=[df['TrainingCompletionDateTime'], df['IndividualFirstName'], df['IndividualLastName'], df['Email'], df['CorpEmail'], df['TrainingActivityId']])

else:
    fig = px.timeline(df, x_start=start, x_end=finish, y=tasks, title='Microsoft Certification Renewal Insights - ' + PartnerName + " (" + str(MPNId) + ")")
    
# Hide the y axis
fig.update_yaxes(title='y', visible=False, showticklabels=False)

# Upade/Change Layout
fig.update_yaxes(autorange='reversed')
fig.update_layout(
        title_font_size=36,
        font_size=18,
        title_font_family='Helvetica'
        )

# Save Graph and Export to HTML
plotly.offline.plot(fig, filename='Certification Renewals - ' + PartnerNameAlphaNumericOnly + '.html')