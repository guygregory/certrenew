# Certification Renewal Report for Partners

[Certification Renewal Report](https://aka.ms/certrenew) allows you to take the [Trainings report from Partner Center Insights](https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions), and convert it into an overview of all the role-based and specialty certifications within your partner organisation. In addition, it will calculate the 6 month window during which a learner can renew their certifcation via a [free Microsoft Learn assessment](https://aka.ms/CertRenewalOverview). This new approach to help learners stay current will be introduced in February 2021. For more information, please see the [public announcment](https://aka.ms/CertRenewalBlog)

## How to use - Downloading the raw data from Partner Center Insights

* Browse to the [Partner Center Insights](https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions)
* Select report type "Trainings", select timeframe "Lifetime", and select either File extension "TSV" or "CSV" (both formats are supported)
* Click the Generate button to create the report
* Download the .CSV/.TSV file via the link in the "Generated Reports" section
* The script requires that you have Python installed, which can be downloaded [here](https://www.python.org/downloads/)

![](media/pci-training.png)

## How to use - Running the script

* The script requires that you have Python installed, which can be downloaded [here](https://www.python.org/downloads/)
* Required modules: [Pandas](https://pandas.pydata.org/) and [XlsxWriter](https://xlsxwriter.readthedocs.io/)
* Download/clone the script, and amend the inputfile and outputfile variables to the appropriate folders on your computer
* When running the .py script, there should be no console output
* If the program is successful, a file will be generated at the location specified by outputfile

## Please note

* This is not an official Microsoft tool, and is maintained in my spare time
* The 'CertRenewalWindowOpens' and 'CertRenewalDeadline' columns are calculated based on rules, and are not derived via API, so please do not treat these as authoratative
* If you need to check the actual dates, please refer to the learner's MCP Transcript, or their public YourAcclaim page

## Feedback

*  Reach me on Twitter [@guygregory](https://twitter.com/guygregory)
