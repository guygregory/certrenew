# Download the latest version!

[CertRenew.zip | Version 0.3 beta | .ZIP archive | Windows x64](https://github.com/guygregory/certrenew/releases/latest/download/CertRenew.zip)

**Certification Renewal Report now supports the updated Partner Center Insights training report format!** (June 17th 2021 or later)
![image](https://user-images.githubusercontent.com/16044916/122179947-e39fc800-ce7f-11eb-893d-387cc0aa9d47.png)

# Certification Renewal Report for Microsoft Partners

In March 2021, Microsoft introduced a new approach to help learners stay current with their certifications, please see the public announcement [here](https://aka.ms/CertRenewalBlog).

This [Certification Renewal Report](https://aka.ms/certrenew) allows partner admins to take the [Trainings report from Partner Center Insights](https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions), and convert it into an overview of all the role-based and specialty certifications within their partner organisation. In addition, it will calculate the 6 month window during which a learner can renew their certifcation via a [free Microsoft Learn assessment](https://aka.ms/CertRenewalOverview).
![](media/ganttsummary2.png)
## How to use - Downloading the source data from Partner Center Insights

* If you're new to Partner Center Insights, please refer to [this documentation first.](https://docs.microsoft.com/en-us/partner-center/pci-download-reports)
* For best results, please ensure you have the [Executive Report Viewer role](https://docs.microsoft.com/en-us/partner-center/pci-roles)
* Browse to [Partner Center Insights](https://partner.microsoft.com/en-us/dashboard/partnerinsights/analytics/downloads?report=TrainingCompletions)
* Select report type "Trainings", select timeframe "Lifetime", and select either File extension "TSV" or "CSV" (both formats are supported)
* Click the Generate button to create the report
* Download the .CSV/.TSV file via the link in the "Generated Reports" section

![](media/pci-training.png)



## How to use - Downloading and running the application
<!--
* The script requires that you have Python installed, which can be downloaded [here](https://www.python.org/downloads/)
* Required modules: [Pandas](https://pandas.pydata.org/), [XlsxWriter](https://xlsxwriter.readthedocs.io/), [Tkinter](https://docs.python.org/3/library/tkinter.html), [Plotly, and Plotly Express](https://plotly.com/python/gantt/)
* Download/clone the script onto your local computer
* When running the .py script, a File Open dialog box should appear, allowing you to select the CSV/TSV file
-->

* Download the latest version of the application from [here](https://github.com/guygregory/certrenew/releases/latest/download/CertRenew.zip)
* Open the .ZIP archive, and extract the CertRenew folder
* In the extracted CertRenew folder, run the certrenew.exe executable
![](media/folder.png)
* When prompted to select a file, choose the .CSV or .TSV file that you downloaded in the previous section
![](media/opendialog.png)
* If you want to try the program with test data, there is a sample .CSV file in the /CertRenew/Example folder

If the program is successful, a Gantt chart will open, showing a visual representation of the certification renewal windows for the organisation:
![](media/ganttsummary2.png)

Mousing-over the individual certification will expose additional detail:
![](media/detail.png)

## Please note

* This is not an official Microsoft tool, and is maintained in my spare time

## Feedback

*  Reach me on Twitter [@guygregory](https://twitter.com/guygregory)
