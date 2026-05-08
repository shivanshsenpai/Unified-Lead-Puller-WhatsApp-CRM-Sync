# Automated-WhatsApp-CRM-Logistics-Tracker
A comprehensive Google Apps Script automation that integrates third-party lead generation (like IndiaMART) directly into Google Sheets. It automatically pulls new leads via API, categorizes them based on user-defined custom rules, and manages grouped counters. The script also features a background engine to sync these processed contacts directly to Google Contacts and includes a UI-driven tool to export categorized lead lists as CSV files for WhatsApp bulk messaging.


# Unified Lead Puller & WhatsApp CRM Sync

## Overview
This Google Apps Script project automates the flow of incoming leads from third-party APIs into a Google Sheets CRM. It categorizes leads, assigns broadcast group numbers, syncs them to Google Contacts, and generates CSV exports tailored for WhatsApp broadcast tools. 

## Features
- **Automated API Lead Pulling:** Fetches leads based on a defined date range, handling pagination and rate limits gracefully.
- **Dynamic Categorization:** Uses regex-based rules from a `Config` sheet to sort leads and assign broadcast group numbers (auto-rolling over at 200 contacts per group).
- **Background Contact Sync:** Automatically syncs categorized leads to your Google Contacts using the People API, creating necessary contact groups on the fly.
- **CSV Export Engine:** UI-enabled modal to select specific categories and export them as WhatsApp-ready CSV files directly to your Google Drive.
- **Robust Error Handling & Logging:** Logs API rate limits, sync failures, and import statuses to a dedicated `Logs` sheet.

## Prerequisites
1. A Google Account with access to Google Sheets, Google Drive, and Google Contacts.
2. **Advanced Services:** You must enable the **People API** in your Google Apps Script project (via Services -> Add a service -> People API).
3. A valid API Key for your lead provider (e.g., IndiaMART).

## Installation & Setup
1. Create a new Google Sheet.
2. Go to `Extensions` > `Apps Script`.
3. Clear the default code and paste the generalized script below.
4. Go to `Services` (left sidebar), find **People API**, and click Add.
5. In the script, locate the `GLOBAL SETTINGS` section at the top and replace `'YOUR_API_KEY_HERE'` with your actual API key.
6. Refresh your Google Sheet. You will see a new custom menu: `IndiaMART + WhatsApp CRM`.
7. Click `Initialize CRM Automation` from the custom menu to set up the necessary background triggers.

## Required Sheets
The script will automatically generate the required sheets if they do not exist, but you should ensure the `Config` sheet is set up with your categorization rules.
- **customer:** Raw incoming leads.
- **Processed_data:** Categorized and formatted leads ready for syncing.
- **Config:** Your regex rules and group counters (e.g., Category Name | Initial | Keywords).
- **Logs:** System outputs and error logs.
