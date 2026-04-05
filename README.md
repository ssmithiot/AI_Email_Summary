# Outlook Inbox Summarizer

A lightweight Windows app that reads your Outlook inbox, scores emails with your custom rules, and generates an AI summary using OpenAI or Anthropic. Anthropic Sonnet is the default provider.

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
- An OpenAI or Anthropic API key

## Easiest Install
1. Download or clone this folder.
2. Make sure Outlook desktop is already open and signed in.
3. Double-click `START_HERE.bat`.
4. If setup appears, let it finish.
5. When Notepad opens `.env`, add your provider key:
   `DEFAULT_AI_PROVIDER=anthropic`
   `OPENAI_API_KEY=sk-...`
   or
   `ANTHROPIC_API_KEY=...`
6. Save the file, close Notepad, and run `START_HERE.bat` again.
7. Your browser should open the app at `http://localhost:5001`.

If your friend only reads one section of this README, make it this one.

## How To Get API Keys
### OpenAI
1. Sign in or create an account at:
   `https://platform.openai.com/`
2. Open the API keys page:
   `https://platform.openai.com/api-keys`
3. Create a new secret key
4. Copy it immediately
5. Put it in the `.env` file in this folder:
   `OPENAI_API_KEY=sk-...`

### Anthropic
1. Sign in or create an account at:
   `https://console.anthropic.com/`
2. Create an API key
3. Put it in the `.env` file in this folder:
   `ANTHROPIC_API_KEY=...`

Important:
- Keep your API key private
- Do not share your `.env` file
- API usage may require a funded OpenAI account depending on your usage limits

## Running The App
- Easiest: double-click `START_HERE.bat`
- Web app directly: double-click `run_summary.bat`
- Console summary: double-click `ckemail.bat`

## First 3 Minutes
1. Click `Summarise Inbox`
2. Review the AI summary and the scored email cards
3. Drag any important email into `Watching`
4. Click `Check Replies` to see which emails you already answered

## First-Time Setup Notes
- `setup.bat` creates a local virtual environment in `venv`
- `rules.json` stores your scoring rules
- `watching.db` stores watched threads locally on your machine

## Files Your Friends Should Edit
- `.env`: add their OpenAI key, Anthropic key, or both
- `rules.json`: optional, customize email scoring rules

## Windows Installer
- Use [build_installer.ps1](C:/projects/AI_Email_Summary/build_installer.ps1) to build the packaged app and installer.
- The Inno Setup script lives at [installer/AI_Email_Summary.iss](C:/projects/AI_Email_Summary/installer/AI_Email_Summary.iss).
- The installer prompts for OpenAI and Anthropic keys and writes `.env` during setup.

## Privacy Notes
- Outlook emails are read from the local Outlook desktop app
- Watched threads are stored locally in `watching.db`
- Email content sent for summarization goes to the selected AI provider through the API

## Troubleshooting
- If setup fails, make sure Python is installed and available in PATH
- If Outlook cannot be found, open Outlook and make sure you are signed in
- If summaries fail, confirm the selected provider key is valid and not still a placeholder
- If larger inbox summaries struggle, the app now falls back to a local summary so you still get results
- If a package is missing, rerun `setup.bat`
- If `START_HERE.bat` stops after opening `.env`, save your key and run it again
