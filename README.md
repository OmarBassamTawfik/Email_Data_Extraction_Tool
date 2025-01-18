# Email Data Extraction Tool

This repository contains a Python notebook (`code.ipynb`) developed during my internship as a **Data Engineer** at **Telecom Egypt**. The notebook is designed to extract structured data from email files (`.msg` format) and save the extracted data into an Excel file.

---

## Overview

The notebook automates the process of extracting specific information from emails related to network outages and site issues. It uses Python libraries such as `pandas`, `extract_msg`, and `re` (for regular expressions) to parse email content and extract relevant fields like **Ticket ID**, **Site ID**, **Outage Start/End**, **Status**, and more.

---

## Features

- **Email Parsing**: Extracts data from `.msg` email files.
- **Regex-Based Extraction**: Uses regular expressions to identify and extract specific fields from the email body.
- **Data Structuring**: Organizes the extracted data into a structured format using a Pandas DataFrame.
- **Excel Export**: Saves the extracted data into an Excel file for further analysis.

---

## How It Works

1. **Input Folder**: The notebook reads `.msg` files from a specified input folder.
2. **Data Extraction**:
   - Extracts fields such as **Ticket ID**, **Site ID**, **Outage Start/End**, **Status**, **Alarm Name**, and more using regex.
   - Handles missing or optional fields gracefully.
3. **Data Storage**:
   - Stores the extracted data in a Pandas DataFrame.
   - Exports the DataFrame to an Excel file (`Down_Sites.xlsx`) in a specified output folder.

---

## Key Libraries Used

- **`pandas`**: For data manipulation and structuring.
- **`extract_msg`**: For reading and parsing `.msg` email files.
- **`re`**: For regular expression-based text extraction.
- **`os`**: For file and directory operations.

---

## Example Output

The extracted data is saved in an Excel file (`Down_Sites.xlsx`) with the following columns:

| Column Name        | Description                                      |
|--------------------|--------------------------------------------------|
| Ticket ID          | Unique identifier for the ticket.               |
| Number of Sites    | Number of sites affected.                       |
| Site ID            | Identifier for the affected site.               |
| Related BSC/RNC    | Related Base Station Controller/RNC.            |
| Status             | Current status of the issue (e.g., Resolved).   |
| Alarm Name         | Name of the alarm triggered.                    |
| Outage Start       | Start time of the outage.                       |
| Outage End         | End time of the outage.                         |
| Operation Area     | Geographical area of the operation.             |
| Operation Zone     | Zone within the operation area.                 |
| Assignment         | Team or group assigned to resolve the issue.    |
| Subcontractor      | Subcontractor responsible for the site.         |
| Vendor             | Vendor associated with the site.                |
| Technology         | Technology used (e.g., RAN-3G).                |
| System             | System affected (e.g., Investigation).         |
| Subsystem          | Subsystem affected (e.g., Under Investigation).|
| Solution           | Proposed or implemented solution.              |
