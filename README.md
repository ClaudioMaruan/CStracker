# CS2 Match Calendar

An automated solution to track Counter-Strike 2 matches and synchronize them with a personal Outlook Calendar.

## About

This project was created to solve a personal problem: keeping track of upcoming CS2 matches without manually checking schedules and adding events to my calendar.

The automation uses Python to retrieve match data from CitoAPI and synchronize relevant events with Outlook Calendar. GitHub Actions runs the process automatically on a scheduled basis.

## Tech Stack

- **Python:** Data processing and automation.
- **CitoAPI:** Match data retrieval.
- **Microsoft Graph API:** Outlook Calendar integration.
- **GitHub Actions:** Scheduled workflow execution.
- **GitHub Secrets:** Credential management.

## How It Works

1. GitHub Actions triggers the workflow on a schedule.
2. Python retrieves, filter and processes match data from CitoAPI.
3. Relevant matches are synchronized with Outlook Calendar.
4. Workflow results can be monitored through GitHub Actions.

## Key Learnings

- API integration and data processing.
- Task automation with Python.
- Scheduled workflow configuration.
- Credential management and troubleshooting.

## Purpose

This project is focused in solving a real-world problem through automation.
