# Certification Renewal Report for Partners
# Guy Gregory - guy.gregory@microsoft.com

import pandas as pd
import plotly.express as px
import plotly
import plotly.figure_factory as ff 
import xlsxwriter
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Specify the location of the file which can be downloaded from Partner Center Insights:
# https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions

# If today's date is after June 1st, then we may have changed the certificate renewal dates
# from 2 years to 1 year, so this tool won't be reliable after this date

if datetime.today() >= datetime(2021, 6, 1):
    messagebox.showerror(message="This version of the tool has expired, please download the latest version from https://github.com/guygregory/certrenew", title="Certification Renewal Tool")
    quit()

root = Tk()
root.withdraw()
root.filename = filedialog.askopenfilename(title="Please select the Trainings report from Partner Center Insights", filetypes=(("CSV files", "*.csv"),("TSV files", "*.tsv"), ("All files", "*.*")) )
inputfile = root.filename

# Load the dataframe from the .csv or .tsv file
name, ext = os.path.splitext(inputfile)

if ext.lower() == ".csv":
    #print(".csv file")
    dfl = pd.read_csv(inputfile, header=0)
    
elif ext.lower() == ".tsv":
    #print(".tsv file")
    dfl = pd.read_csv(inputfile, sep='\t', header=0)
    
else:
    messagebox.showerror(message="Not a valid file, please use Partner Skills Report - Trainings in either .csv or .tsv format", title="Certification Renewal Tool")
    quit()

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

# Important - Pandas stores dates in datetime format by default, so to display only the date element, then use dt.date 
# for example: df['just_date'] = df['dates'].dt.date

rbstable["TrainingCompletionDateTime"] = pd.to_datetime(rbstable["TrainingCompletionDate"])

# Convert value to date format
rbstable["TrainingCompletionDateOnly"] = pd.to_datetime(rbstable["TrainingCompletionDateTime"]).dt.date


# Initial certification validity = 2 years, and are subject to two extensions:
#
#                           Start Date      End Date            Extension
# March 2020 extension      March 26 2020   December 31 2020    6 months
# November 2020 extension   January 1 2021	June 30 2021        6 months

# Cert Completion Date              Years                       Notes
# July 1 2018 – December 31 2018    3       Certs in this window get both 6 month extensions
# January 1 2019 – June 30 2019	    2.5     These certs would get the November 2020 extension
# August 1 2019 – now	            2       No extensions for these certs

ext1start = datetime(2020, 3, 25)
ext1end = datetime(2021, 1, 1)
ext2start = datetime(2020, 12, 31)
ext2end = datetime(2021, 7, 1)
earliest = datetime(2000, 1, 1)

rbstable["CertRenewalDeadlineDateTime"] = rbstable["TrainingCompletionDateTime"]

rbstable["CertRenewalDeadlineDateTime"] = rbstable["CertRenewalDeadlineDateTime"].apply(lambda x: x + relativedelta(months=24) if x > earliest else x + relativedelta(months=0))
rbstable["CertRenewalDeadlineDateTime"] = rbstable["CertRenewalDeadlineDateTime"].apply(lambda x: x + relativedelta(months=6) if ext1start < x < ext1end else x + relativedelta(months=0))
rbstable["CertRenewalDeadlineDateTime"] = rbstable["CertRenewalDeadlineDateTime"].apply(lambda x: x + relativedelta(months=6) if ext2start < x < ext2end else x + relativedelta(months=0))

rbstable["CertRenewalDeadline"] = pd.to_datetime(rbstable["CertRenewalDeadlineDateTime"]).dt.date

#Create a new column for the start of the cert renewal window, and set the Cert Renewal Window Opens values to 6 months before the Exam Expiry Date
rbstable["CertRenewalWindowOpens"] = rbstable["CertRenewalDeadline"] + relativedelta(months=-6)

#Clean-up unused columns
rbstable = rbstable.drop(columns=["TrainingCompletionDate","TrainingCompletionDateOnly","CertRenewalDeadlineDateTime"])

#Re-order the columns into the desired order
rbstable = rbstable[["TrainingActivityId", "TrainingTitle", "IndividualFirstName", "IndividualLastName", "Email", "CorpEmail", "TrainingCompletionDateTime", "CertRenewalWindowOpens", "CertRenewalDeadline"]]

#Export to .xlsx

# Specify the output filename for the Excel spreadsheet.:
outputfile = "Certification Renewals - " + PartnerNameAlphaNumericOnly + ".xlsx"

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter(outputfile, engine="xlsxwriter")

# Convert the dataframe to an XlsxWriter Excel object.
rbstable.to_excel(writer, sheet_name="Sheet1", index=False)

# Get the xlsxwriter workbook and worksheet objects.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# Set the column widths in Sheet1
worksheet.set_column('A:A', 16)
worksheet.set_column('B:B', 40)
worksheet.set_column('C:D', 20)
worksheet.set_column('G:I', 27)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

# Use this for debugging if you need to output the certs to the console instead of Excel
# print (rbstable.head(50))


### Create and Save the HTML Report ###

df = rbstable # Create a copy of the dataframe


# replacing null values in dataframe with a blank space to avoid %{customdata[1]}
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
fig = px.timeline(df, x_start=start, x_end=finish, y=tasks, title='Microsoft Certification Renewal Insights - ' + PartnerName + " (" + str(MPNId) + ")", 

hover_name=df['TrainingTitle'], 
hover_data=[df['TrainingCompletionDateTime'], df['IndividualFirstName'], df['IndividualLastName'], df['Email'], df['CorpEmail'], df['TrainingActivityId']])

# Hide the y axis
fig.update_yaxes(title='y', visible=False, showticklabels=False)

# Upade/Change Layout
fig.update_yaxes(autorange='reversed')
fig.update_layout(
        title_font_size=36,
        font_size=18,
        title_font_family='Helvetica'
        )

# Interactive Gantt
#fig = ff.create_gantt(df)

# Save Graph and Export to HTML
plotly.offline.plot(fig, filename='Certification Renewals - ' + PartnerNameAlphaNumericOnly + '.html')
#plotly.io.write_html(fig, file='Certification Renewals - ' + PartnerNameAlphaNumericOnly + '.html', auto_open=False)