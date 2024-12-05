# Automated Email Attachment Retrieval from Outlook

This R script automates the retrieval of email attachments from Microsoft Outlook using the `RDCOMClient` package, which establishes a COM (Component Object Model) connection with Outlook. The script is designed to efficiently extract Excel file attachments from specific emails in your inbox for automated data processing.

## Key Features

- **Connects to Microsoft Outlook**: Uses the `RDCOMClient` package to establish a connection with Outlook.
- **Navigates to Specific Folders**: Accesses specific folders within the inbox to retrieve targeted emails.
- **Filters Emails by Date Range**: Searches for emails within a specified date range (e.g., September 7, 2022).
- **Identifies Excel Attachments**: Filters emails for attachments with the `.xls` file extension.
- **Automates Data Extraction**: Extracts the second Excel attachment, saves it temporarily, and imports it into R for analysis.

## Workflow

1. **Connecting to Outlook**:  
   The script uses `RDCOMClient` to establish a COM connection with Microsoft Outlook.

2. **Searching for Emails**:  
   It navigates to a specific folder within your inbox and filters emails by date (e.g., September 7, 2022).

3. **Identifying Attachments**:  
   For each email, the script looks for attachments, focusing on Excel files (those with the `.xls` extension).

4. **Extracting and Reading Excel Files**:  
   When an Excel file is found, the script:
   - Stores the filenames of all attachments in a variable.
   - Selects the second attachment from the list.
   - Saves the file temporarily.
   - Reads the file into R using the `readxl` package, skipping the first two rows of data.

This automated pipeline simplifies the process of extracting data from Excel attachments in Outlook emails, making it a valuable tool for automated data processing workflows.

## Technologies & Libraries

- **R Libraries**:  
   - `RDCOMClient`: Used to connect to Outlook and retrieve emails and attachments.
   - `readxl`: Used to read Excel files into R.

## Instructions

1. Install the required R packages if you haven't already:
   ```R
   install.packages("RDCOMClient")
   install.packages("readxl")
