# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Monthly KPI Report automation system for an accounting/consulting firm. The main script processes staff performance data from CSV ledgers, calculates billable hours and revenue targets, generates Excel reports, and sends automated performance emails to staff.

## Key Architecture

The project consists of a single Python script (`Automatedisplay_ledger.py`) that performs these sequential operations:

1. **Data Processing**: Reads and cleans ledger data from CSV
2. **KPI Calculations**: Computes billable hours, revenue, and target achievement
3. **Report Generation**: Creates a master Excel file with multiple sheets
4. **Individual Reports**: Generates personalized Excel reports for each staff member
5. **Email Automation**: Sends performance emails with attached reports

## Dependencies

The script requires these Python packages (no requirements.txt exists):
- pandas
- openpyxl
- yagmail

## Common Commands

```bash
# Install dependencies
pip install pandas openpyxl yagmail

# Run the main script
python Automatedisplay_ledger.py
```

## Important Configuration

1. **Input Files Required**:
   - CSV ledger file (currently hardcoded as "Mar ledger for Jun25 report.csv")
   - staff_emails.csv (contains Staff and Email columns)

2. **Hardcoded Values** in the script:
   - Staff billable rates and percentages (lines 37-43)
   - Excluded staff list (line 45)
   - Output Excel filename (line 7)
   - Email credentials (lines 229-230) - currently blank

3. **Key Data Fields** expected in the CSV:
   - [Ledger] Staff
   - [Ledger] Time
   - [Ledger] Invoiced Amount
   - [Ledger] Billable Amount
   - [Job] Job No.

## Development Notes

- The script has no error handling for missing files or malformed data
- Email credentials need to be configured before running email functionality
- All configuration is hardcoded within the script
- Job number "J003167" is treated specially for non-billable time calculations