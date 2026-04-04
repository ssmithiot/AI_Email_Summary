# Outlook Inbox Summarizer

A lightweight Windows app that reads your Outlook inbox, scores emails with your custom rules, and generates an AI summary using OpenAI.

## What It Does
- Connects to the Outlook desktop app on Windows
- Pulls inbox emails from Outlook
- Scores emails using built-in and custom rules
- Summarizes the inbox with OpenAI
- Lets you open emails directly in Outlook
- Tracks watched threads in the Watching panel
- Checks Sent Items to find replies you already sent

## Requirements
- Windows
- Microsoft Outlook desktop app installed and signed in
- Python 3.11+ installed
- An OpenAI API key

## Easiest Install
1. Download or clone this folder.
2. Double-click `setup.bat`.
3. When prompted, get your OpenAI API key from:
   `https://platform.openai.com/api-keys`
4. Paste the key into `.env` so it looks like:
   `OPENAI_API_KEY=sk-...`
5. Double-click `run_summary.bat`.

The browser app will open automatically.

## How To Get An OpenAI API Key
1. Sign in or create an account at:
   `https://platform.openai.com/`
2. Open the API keys page:
   `https://platform.openai.com/api-keys`
3. Create a new secret key
4. Copy it immediately
5. Put it in the `.env` file in this folder:
   `OPENAI_API_KEY=sk-...`

Important:
- Keep your API key private
- Do not share your `.env` file
- API usage may require a funded OpenAI account depending on your usage limits

## Running The App
- Web app: double-click `run_summary.bat`
- Console summary: double-click `ckemail.bat`

## First-Time Setup Notes
- `setup.bat` creates a local virtual environment in `venv`
- `rules.json` stores your scoring rules
- `watching.db` stores watched threads locally on your machine

## Files Your Friends Should Edit
- `.env`: add their OpenAI API key
- `rules.json`: optional, customize email scoring rules

## Privacy Notes
- Outlook emails are read from the local Outlook desktop app
- Watched threads are stored locally in `watching.db`
- Email content sent for summarization goes to OpenAI through the API

## Troubleshooting
- If setup fails, make sure Python is installed and available in PATH
- If Outlook cannot be found, open Outlook and make sure you are signed in
- If summaries fail, confirm your `OPENAI_API_KEY` is valid
- If a package is missing, rerun `setup.bat`

## Beta MVP
This is currently a beta MVP focused on easy local setup for friends and small-team testing.
