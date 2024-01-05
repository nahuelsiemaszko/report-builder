# report-builder

Macro created in VBA to automate a process that was performed every month, reducing the time to complete the task from 40 minutes to 2 minutes.

This process involved loading and modifying specific information from an xlsx file into an xlsm file, which then is used to enhance monthly reporting.

This macro performs the following actions as per the given specifications:

- Copy data from the "New" and "Closed" sheets of the source file to the destination file.
- Update the "WorkedTickets" sheet with information on tickets received and closed in the last month, while maintaining a record of the last 6 months and removing the earliest month.
- Update the information regarding the quantity of each type of case in the "MonthlyNewCases" sheet.
- Update the information on the "SLA" sheet concerning "Incident Response Within Deadline by Priority" and "Incident Resolution Within Deadline by Priority," categorizing them into critical, high, medium, and low priorities.

If you would like to try it out, please refer to the 'usage-instructions.md' file for detailed usage instructions.
