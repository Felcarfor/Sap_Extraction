# SAP Data Extraction and Integration Tool

## Overview

This Python application facilitates data extraction from SAP using predefined queries (`mb51`) and integrates the extracted data into a Google Sheets document.

## Requirements

- Python 3.x
- Libraries:
  - `pandas`: Data manipulation and analysis.
  - `win32com`: Windows automation for interacting with SAP GUI.
  - `gspread_pandas`: Integration between pandas and Google Sheets for data updating.
  - `tkinter`: GUI library for user interface.
  - `oauth2client`: OAuth2 authentication for accessing Google services, essential for Google Sheets.

## Setup

### Installation

Ensure Python and required libraries are installed:
```bash
pip install pandas gspread-pandas pywin32 oauth2client
