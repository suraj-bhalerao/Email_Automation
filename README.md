# Email Automation for Job Applications

This Python script automates sending job application emails to multiple recipients listed in an Excel file. It attaches your resume and sends personalized emails with random subject and body variations to avoid spam filters.

## Features

- Reads recipient emails from an Excel spreadsheet
- Sends emails with resume attachment
- Randomizes subject lines and email bodies
- Includes delays between sends to prevent spam detection
- Moves processed Excel file to 'isSent' folder

## Prerequisites

- Python 3.x
- Gmail account with App Password (for SMTP)
- Excel file with recipient emails in column A (starting from row 2)

## Installation

1. Clone or download this repository.
2. Install required packages:

```bash
pip install -r requirements.txt
```

## Setup

1. Create a `.env` file in the root directory with the following variables:

```
SENDER_EMAIL=your_email@gmail.com
APP_PASSWORD=your_app_password
EXCEL_PATH=data/your_excel_file.xlsx
RESUME_PATH=Suraj_Resume_02.pdf
```

- `SENDER_EMAIL`: Your Gmail address
- `APP_PASSWORD`: Generate an App Password from your Google Account settings
- `EXCEL_PATH`: Path to your Excel file containing emails
- `RESUME_PATH`: Path to your resume PDF

2. Place your Excel file in the `data/` folder.

## Usage

Run the script:

```bash
python email_auto.py
```

The script will:
- Send emails to all recipients in the Excel file
- Wait random delays (15-120 seconds) between sends
- Move the Excel file to `isSent/` after completion

## Folder Structure

```
EmailAutomation/
├── email_auto.py          # Main script
├── .env                   # Environment variables (create this)
├── data/                  # Folder for Excel files
├── isSent/                # Processed Excel files
├── Suraj_Resume_02.pdf    # Your resume
└── README.md              # This file
```

## Notes

- Ensure your Gmail has "Less secure app access" enabled or use App Passwords.
- The script uses Gmail's SMTP server.
- Be mindful of email sending limits to avoid account suspension.

## License

[Add license if any]