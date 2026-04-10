# AHA Website Automation Demo

A browser automation demo built as part of a **CSC 131 real-client course project**.  
This workflow was designed to help reduce repetitive manual steps for the client by automating part of the AHA training request process.

## Demo Video

[Watch the demo video](https://youtu.be/oiFJd4WCFI4)

## Project Context

This automation demo is one component of a larger CSC 131 client project.  
The broader project was built for a real client in class, and one of its goals was to reduce repetitive, time-consuming manual tasks in the client's workflow.

This specific automation focuses on part of the AHA training request process, including class lookup, request handling, and transferring student data into Google Sheets.

## What This Demo Shows

This demo shows an end-to-end workflow that:

- opens the AHA site with a saved login session
- goes to the training class listing page
- selects the organization
- types the instructor name
- selects the target date
- opens matching classes
- accepts pending requests
- extracts student information
- appends student records into Google Sheets

The demo uses fixed test values:

- Organization: `Sac State`
- Instructor: `Sac State`
- Date: `04/07/2026`

## My Contribution

My primary work in this demo focused on:

- building the browser automation workflow in Python with Playwright
- saving and reusing login state for repeated runs
- navigating the AHA class listing workflow
- opening class records and accepting pending requests
- extracting student information from the roster
- appending student records into Google Sheets

## Privacy and Demo Data Note

All login information, Google Sheet content, email content, and any student-related data shown in the demo video are **mock and demo-only data** used for testing and presentation purposes.

This public demo does **not** include:

- real client credentials
- real student private information
- sensitive production data

## Tech Stack

- Python
- Playwright
- Google Sheets API
- python-dotenv

## Files

- `setup_login.py` — first-time login setup and saved session creation
- `run_automation.py` — main automation workflow using fixed test values

## Installation

Run:

`pip install -r requirements.txt`

Then run:

`python -m playwright install`

## Run

First-time login setup:

`python setup_login.py`

Run the automation:

`python run_automation.py`

## Public Repo Notes

Sensitive local files are intentionally excluded from this public repository, including:

- `.env`
- `aha_auth.json`
- `google_sheet_api_key.json`

## Purpose

The main purpose of this automation was to reduce repetitive manual work in the client workflow and demonstrate how browser automation can support a more efficient training request handling process.