# Certification Renewal Report for Partners
# Guy Gregory - guy.gregory@microsoft.com

import pandas as pd
import plotly.express as px
import plotly
import xlsxwriter
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Specify the location of the .tsv file which can be downloaded from Partner Center Insights:
# https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions

inputfile = r"C:\CertRenew\Export_trainings_Lifetime_EXAMPLE_PLEASE_EDIT_THIS_FILENAME.tsv"

# Load the dataframe from the .csv or .tsv file
name, ext = os.path.splitext(inputfile)

if ext.lower() == ".csv":
    #print(".csv file")
    dfl = pd.read_csv(inputfile, header=0)
    
elif ext.lower() == ".tsv":
    #print(".tsv file")
    dfl = pd.read_csv(inputfile, sep='\t', header=0)
    
else:
    print("Not a valid file, please use Partner Skills Report - Trainings in either .csv or .tsv format")
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

# NB - the original cutoff was certs that expired 'before 31st December 2020', but that was later revised to 'before 30th June 2021' when we announced certifcation renewals
# https://techcommunity.microsoft.com/t5/microsoft-learn-blog/an-important-update-on-microsoft-training-and-certification/ba-p/1489671
extensiondatecutoff = datetime(2019, 7, 1)

# Use this for debugging, it will add a column which will show if a cert is eligible for extension (True/False)
# rbstable["EligibleForExtension"] = rbstable["TrainingCompletionDateTime"].apply(lambda x: True if x < extensiondatecutoff else False)

rbstable["CertRenewalDeadlineDateTime"] = rbstable["TrainingCompletionDateTime"].apply(lambda x: x + relativedelta(months=30) if x < extensiondatecutoff else x + relativedelta(months=24))

rbstable["CertRenewalDeadline"] = pd.to_datetime(rbstable["CertRenewalDeadlineDateTime"]).dt.date

#Create a new column for the start of the cert renewal window, and set the Cert Renewal Window Opens values to 6 months before the Exam Expiry Date
rbstable["CertRenewalWindowOpens"] = rbstable["CertRenewalDeadline"] + relativedelta(months=-6)

#Clean-up unused columns
rbstable = rbstable.drop(columns=["TrainingCompletionDate","TrainingCompletionDateOnly","CertRenewalDeadlineDateTime"])

#Re-order the columns into the desired order
rbstable = rbstable[["TrainingActivityId", "TrainingTitle", "IndividualFirstName", "IndividualLastName", "Email", "CorpEmail", "TrainingCompletionDateTime", "CertRenewalWindowOpens", "CertRenewalDeadline"]]

#Export to .xlsx

# Specify the output location that you would like the final report to be saved.:
outputfile = r"C:\CertRenew\Certification Renewals - " + PartnerNameAlphaNumericOnly + ".xlsx"
#outputfile = r"C:\CertRenew\Certification_Renewals_EXAMPLE_PLEASE_EDIT_THIS_FILENAME.xlsx"

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

# Sort by date, and then reindex
df = df.sort_values(by='CertRenewalWindowOpens')
df = df.reset_index()

# Assign Columns to variables
tasks = df.index
start = df['CertRenewalWindowOpens']
finish = df['CertRenewalDeadline']

# Create Gantt Chart
fig = px.timeline(df, x_start=start, x_end=finish, y=tasks, title='Certification Renewal Windows', 

hover_name=df['TrainingTitle'], 
hover_data=[df['TrainingCompletionDateTime'], df['IndividualFirstName'], df['IndividualLastName'], df['Email'], df['CorpEmail'], df['TrainingActivityId']])

# Hide the y axis
fig.update_yaxes(title='y', visible=False, showticklabels=False)

# Upade/Change Layout
fig.update_yaxes(autorange='reversed')
fig.update_layout(
        title_font_size=42,
        font_size=18,
        title_font_family='Arial'
        )

# Interactive Gantt
#fig = ff.create_gantt(df)

# Save Graph and Export to HTML
plotly.offline.plot(fig, filename='Certification Renewal - ' + PartnerNameAlphaNumericOnly + '.html')
