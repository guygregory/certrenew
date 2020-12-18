# Certification Renewal Report for Partners

[Certification Renewal Report](https://aka.ms/certrenew) allows you to take the [Trainings report from Partner Center Insights](https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions), and convert it into an overview of all the role-based, and specialty certifications within your partner organisation. In addition, it will calculate the 6 month window during which a learner can renew their certifcation via a free Microsoft Learn assessment. This new approach to help learners stay current will be introduced in February 2021. For more information, please see the [public announcment](https://aka.ms/CertRenewalBlog)

## How to use - Downloading the raw data from Partner Center Insights

* Browse to the [Partner Center Insights](https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions)
* Select report type "Trainings", select timeframe "Lifetime", and select File extension "TSV"
* Click the Generate button to create the report
* Download the .TSV file via the link in the "Generated Reports" section.
* The script requires that you have Python installed, which can be downloaded [here](https://www.python.org/downloads/).

![](media/pci-training.png)

## How to use - Running the script

* The script requires that you have Python installed, which can be downloaded [here](https://www.python.org/downloads/).
* Required modules: [Pandas](https://pandas.pydata.org/) and [XlsxWriter](https://xlsxwriter.readthedocs.io/)
* Download/clone the script, and amend the inputfile and outputfile variables to the appropriate folders on your computer.
* Execute the .py script, there should be no console output, but if the program is successful, a file will be generated at the location specified by outputfile.

## Feedback

*  Reach me on Twitter [@guygregory](https://twitter.com/guygregory)
