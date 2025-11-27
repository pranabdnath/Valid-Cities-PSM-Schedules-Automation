# Valid-Cities-PSM-Schedules-Automation
This script automates the extraction, filtering, and processing of scheduling data received through email.


# PSMS Data Processing Automation

This repository contains Google Apps Script functions designed to automate the processing of PSMS schedule data. The goal of these scripts is to streamline daily data extraction, filtering, formatting, and flag marking without manually handling CSV or ZIP attachments.

## Purpose
The script imports a ZIP file from the latest email matching a specific subject, extracts the CSV, filters the data based on city groups and matching conditions, formats dates, writes the final results into separate sheets, and generates 30-day Y/N activity flags.  
It also creates detailed logs for debugging and tracking.

## Features
- Reads the most recent email containing the specified subject line.
- Extracts CSV from ZIP attachments using Drive.
- Filters rows into North, South, and No-Date categories.
- Normalizes and formats specific date columns.
- Removes older data while retaining sheet headers.
- Deduplicates records based on Buyer and Project.
- Adds "LastActivity" column with formatted dates.
- Generates 30-day repeating Y/N flags.
- Logs each step into a dedicated ScriptLogs sheet.

## File Overview
- `validCitiesPSMSSchedules.js`: Main script responsible for extracting data, filtering results, and updating all output sheets.
- `helperFunctions.js`: Contains supporting functions for date parsing, flag marking, and logging.
  
## Sheets Used
- Schedules-Noida  
- Schedules-South  
- NoDate-BothCenters  
- ScriptLogs

## How It Works
1. Gmail is scanned for the most recent email with the defined subject line.  
2. The ZIP attachment is saved to Drive and extracted.  
3. The CSV file inside the ZIP is read and converted to a 2D array.  
4. Each row is checked for matching city groups and matching conditions.  
5. Cleaned and filtered data is inserted into the three target sheets.  
6. Date columns are standardized into a consistent format.  
7. A 30-day logic assigns “Y” or “N” for each Buyer–Project combination.  
8. Detailed logs are recorded in the ScriptLogs sheet.

## Notes
All references to internal tools, sheets, or mail subjects have been generalized for privacy. Replace the IDs, sheet names, and subject lines with your own environment-specific values.

