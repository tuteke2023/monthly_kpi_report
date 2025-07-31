# Monthly KPI Report Automation System

An automated system for processing staff performance data, calculating KPIs, and generating personalized performance reports for an accounting/consulting firm.

## Overview

This system automates the monthly KPI reporting process by:
- Processing staff timesheet data from CSV ledgers
- Calculating billable hours and revenue targets
- Generating comprehensive Excel reports
- Sending automated performance emails with personalized reports

## Features

- **Automated Data Processing**: Reads and cleans ledger data from CSV files
- **KPI Calculations**: Computes billable hours, revenue, target achievement rates
- **Multi-sheet Excel Reports**: Creates master reports with summary and detailed views
- **Individual Staff Reports**: Generates personalized Excel files for each team member
- **Email Automation**: Sends performance emails with attached reports (requires configuration)

## Requirements

- Python 3.x
- pandas
- openpyxl
- yagmail

## Installation

1. Clone the repository:
```bash
git clone git@github.com:tuteke2023/monthly_kpi_report.git
cd monthly_kpi_report
```

2. Install dependencies:
```bash
pip install pandas openpyxl yagmail
```

## Configuration

Before running the script, you need to configure:

### 1. Input Files
- **CSV Ledger File**: Place your monthly ledger CSV in the project directory
- **staff_emails.csv**: Create a CSV file with columns:
  - `Staff`: Staff member names (must match ledger data)
  - `Email`: Corresponding email addresses

### 2. Script Configuration
Edit `Automatedisplay_ledger.py` to update:
- **Ledger filename** (line 7): Update to match your CSV file
- **Staff billable rates** (lines 37-43): Adjust rates and percentages as needed
- **Excluded staff list** (line 45): Modify staff exclusions
- **Email credentials** (lines 229-230): Add your email credentials

### 3. CSV Data Format
Your ledger CSV must include these columns:
- `[Ledger] Staff`: Staff member name
- `[Ledger] Time`: Hours worked
- `[Ledger] Invoiced Amount`: Invoiced revenue
- `[Ledger] Billable Amount`: Billable revenue
- `[Job] Job No.`: Job number (J003167 is treated as non-billable)

## Usage

Run the main script:
```bash
python Automatedisplay_ledger.py
```

The script will:
1. Process the CSV ledger data
2. Generate a master Excel report with multiple sheets
3. Create individual Excel reports for each staff member
4. Send emails with attached reports (if configured)

## Output

- **Master Report**: `Monthly_KPI_Report_[Date].xlsx` with sheets:
  - Summary statistics
  - Detailed performance data
  - Individual staff breakdowns
  
- **Individual Reports**: `[StaffName]_KPI_Report_[Date].xlsx` for each team member

## Notes

- Job number "J003167" is treated as non-billable time
- The script currently has no error handling for missing files
- Email functionality requires valid credentials to be configured
- All configuration is hardcoded in the script (consider using config files for production)

## Future Improvements

- Add error handling for missing/malformed data
- Move configuration to external config file
- Add command-line arguments for input/output files
- Implement logging for better debugging
- Add data validation and sanity checks

## License

This project is proprietary software for internal use.

## Support

For issues or questions, please contact the development team.