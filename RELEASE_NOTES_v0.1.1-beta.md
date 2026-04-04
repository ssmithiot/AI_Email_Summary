# v0.1.1-beta

This beta smooths out the first-run experience and hardens the Outlook side of the app.

## Highlights
- Friendlier Outlook retry and error handling when Outlook is busy
- Cleaner first-run UI with better empty states and scan context
- Watching now calls out fresh thread activity more clearly
- Reply area stays visible and explains what it is for
- `START_HERE.bat` now catches the placeholder OpenAI key before launch
- README now gives a simpler install path and a quick "first 3 minutes" guide

## Best For
- Friends testing the app for the first time
- Lightweight local installs on Windows with Outlook desktop
- Users who want AI summaries plus watched-thread tracking

## Known Beta Notes
- Outlook must be open and signed in
- Large inboxes can still feel slower than small ones
- If OpenAI returns nothing useful, the app falls back to a local summary
