# Certification Expiration Management System

## Overview
This workbook is designed to manage the certification expiration details of healthcare professionals. It automates the process of sending email reminders as certification expiration dates approach.

## Features
- **Expiration Tracking:** Automatically tracks certification statuses, including good standing, expires soon, and expired.
- **Automated Email Reminders:** Sends email notifications at 90, 60, and 30 days before certification expiration.
- **Email Customization:** Allows configuration of email subjects and recipients through settings within the workbook.
- **Daily Updates:** Updates email flags and statuses once at midnight, ensuring reminders are sent timely.

## Functionality
1. **Changing Expiration Dates:** When an expiration date is updated, the system automatically clears the 'Email Sent' flag for that entry, ensuring no duplicate reminders are sent.
2. **Automated Processes:** The system uses VBA scripts to manage data entry, update statuses, and send email notifications efficiently.

## Setup
To use this workbook effectively:
1. Ensure macro settings are enabled in Excel to allow VBA scripts to run.
2. To send emails, ensure your system is configured to use Outlook with Office 365. In Excel, navigate to the 'Developer' tab, click on 'View Code', and adjust the macro settings by commenting out unnecessary details and ensuring the 'Send' commands are active.

## Customization and Expansion
- **Adding New Sheets:** Feel free to create new sheets for additional types of certifications. Ensure that any new sheets maintain the same format as existing ones for consistent functionality.
- **Data and Dashboard Customization:** Edit the data and dashboard settings as needed to fit your specific operational requirements.

