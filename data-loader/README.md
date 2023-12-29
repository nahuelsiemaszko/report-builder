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

If you would like to try it out, please refer to the 'usage-instructions.md' file for detailed usage instructions.
