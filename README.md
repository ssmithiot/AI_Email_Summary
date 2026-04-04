# Outlook Inbox Summarizer

A lightweight Windows app that reads your Outlook inbox, scores emails with your custom rules, and generates an AI summary using OpenAI.

## Beta MVP
This is a beta MVP focused on the easiest local Windows install possible for friends and small-team testing.

## What It Does
- Connects to the Outlook desktop app on Windows
- Pulls inbox emails from Outlook
- Scores emails using built-in and custom rules
- Summarizes the inbox with OpenAI
- Lets you open emails directly in Outlook
- Tracks watched threads in the Watching panel
- Checks Sent Items to find replies you already sent

## Screenshots
Typical app areas your friends will see:
- Inbox Summary: AI-written summary with citations back to the source emails
- Watching: collapsible watched conversation threads
- Replies Found: sent replies you already handled
- Email Cards: scored emails with open, flag, and read actions

Tip: add screenshots to this README later to make friend installs even easier.

## Requirements
- Windows
- Microsoft Outlook desktop app installed and signed in
- Python 3.11+ installed
- An OpenAI API key

## Easiest Install
1. Download or clone this folder.
2. Double-click `START_HERE.bat`.
3. If setup appears, let it finish.
4. When prompted, get your OpenAI API key from:
   `https://platform.openai.com/api-keys`
5. Paste the key into `.env` so it looks like:
   `OPENAI_API_KEY=sk-...`
6. The app will launch in your browser.

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
- Easiest: double-click `START_HERE.bat`
- Web app directly: double-click `run_summary.bat`
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
- If larger inbox summaries struggle, the app now falls back to a local summary so you still get results
- If a package is missing, rerun `setup.bat`
