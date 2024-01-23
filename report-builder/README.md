# report-builder

Macro created in VBA to automate a process that was performed every month, reducing the time to complete the task from 40 minutes to 2 minutes.

This process involved loading and modifying specific information from an xlsx file into an xlsm file, which then is used to enhance monthly reporting.

This macro performs the following actions as per the given specifications:

- Copies data from the "New" and "Closed" sheets of the source file to the destination file.
- Updates the "WorkedTickets" sheet with information on tickets received and closed in the last month, while maintaining a record of the last 6 months and removing the earliest month.
- Updates the information regarding the quantity of each type of case in the "MonthlyNewCases" sheet.
- Updates the information on the "SLA" sheet concerning "Incident Response Within Deadline by Priority" and "Incident Resolution Within Deadline by Priority," categorizing them into critical, high, medium, and low priorities.

# usage-instructions

Download the source file ("report-data-extract.xlsx") and the destination file ("report-builder.xlsm") and locate both the source file and destination file in the same folder.

Open both the source and destination files.

In the "report-builder.xlsm" file, on the horizontal toolbar, click on "Developer" (You should enable this tool if you don't have it enabled), then click on "Macros" and select the "LoadData" macro.

The information from the source file will be copied and modified automatically according to the provided specifications in the destination file.

Notes: The information provided in the example files is for illustrative purposes and has been modified to mantain company confidentiality.
The "update-data.bas" file is provided for code review purposes only, downloading is not necessary.
