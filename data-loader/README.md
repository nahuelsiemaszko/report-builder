# data-loader

Macro created in VBA to automate a process that was performed daily and repetitively, reducing the time to complete the task from 30 minutes to 2 minutes.

This process involved loading and modifying specific information from an xlsx file into an xlsm file, which then makes changes to an internal company database.

This macro performs the following actions as per the given specifications:

- Detects the format of the source file to ensure the correct transfer of information.
- Copies the required columns from the source file to the destination file.
- Changes the rows in the 4 columns that contain "Y" and "N" to "1" and "0".
- Replaces empty fields with "~NULL~".
- Fills the rows in the "refdate" column with today's date.
- Generates a new spreadsheet in the same destination file with the name "Record" "dd-mm-yy" and the loaded information.

# usage-instructions

Download the source file (modify-data-error.xlsx" or "modify-data-error-type-2.xlsx") and the destination file ("data-loader.xlsm") and locate both the source file and destination file in the same folder.

The source file must be named "modify-data-error.xlsx". If you want to try with "modify-data-error-type-2.xlsx", you should rename it to "modify-data-error.xlsx".

Open both the source and destination files.

In the "data-loader.xlsm" file, on the horizontal toolbar, click on "Developer" (You should enable this tool if you don't have it enabled), then click on "Macros" and select the "LoadData" macro.

The information from the source file will be copied and modified automatically according to the provided specifications in the destination file.

Notes: The information provided in the example files is for illustrative purposes and has been modified for company confidentiality.
