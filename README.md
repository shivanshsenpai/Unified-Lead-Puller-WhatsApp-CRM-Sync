# Automated-WhatsApp-CRM-Logistics-Tracker
A comprehensive Google Apps Script automation that integrates third-party lead generation (like IndiaMART) directly into Google Sheets. It automatically pulls new leads via API, categorizes them based on user-defined custom rules, and manages grouped counters. The script also features a background engine to sync these processed contacts directly to Google Contacts and includes a UI-driven tool to export categorized lead lists as CSV files for WhatsApp bulk messaging.
## Overview
This repository contains the generalized code for an automated CRM system. It connects a Google Sheets database with the WhatsApp API to manage logistics, track orders, and automate customer notifications securely and efficiently.

## Features
- **Google Sheets Integration:** Reads and updates CRM data dynamically without hardcoding specific sheet names or IDs.
- **WhatsApp API Automation:** Sends automated status updates and messages to users.
- **Logistics Management:** Tracks order workflows and triggers alerts based on real-time status changes.
- **Secure Configuration:** Designed to use environment variables (`.env`) for all sensitive data, keeping your keys and IDs out of the source code.

## Prerequisites
- Python 3.x
- Google Cloud Console Project (with Google Sheets API and Google Drive API enabled)
- WhatsApp Business API Account
- A Service Account JSON credential file

## Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/yourusername/whatsapp-crm-logistics.git](https://github.com/yourusername/whatsapp-crm-logistics.git)
   cd whatsapp-crm-logistics
