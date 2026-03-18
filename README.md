# MS Forms Auto

OpenClaw skill for automating Microsoft Forms submission with dual-calendar integration and M365 number-matching MFA support.

## Overview

Submits the Xaltius "Daily Productivity Logs" form automatically by pulling data from two calendars:

| Calendar | Feeds Into |
|----------|-----------|
| **Training Calendar (TMS)** | Training Hours (source of truth) |
| **Outlook Calendar** | Content Dev Hours, Other Items |

Learning Hours are auto-calculated as `8 - Training - Content Dev` to ensure an 8-hour workday. Learning Topic is randomized from a curated pool + any KT Sessions found on the calendar.

## Architecture

```
5:45 PM SGT → Pre-Fill cron → calendar-fetch.js → entry JSON
6:00 PM SGT → Submit cron → submit-daily.js → form submitted ✅
```

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/calendar-fetch.js` | Fetch both calendars, output all form fields as JSON |
| `scripts/submit-daily.js` | Submit form using saved auth + entry JSON |
| `scripts/mfa-login.js` | Interactive MFA login (headed mode, saves auth state) |
| `scripts/fill-form.js` | CLI form filler with argument flags |
| `scripts/setup-credentials.js` | One-time credential setup |

## Quick Start

### 1. Install

```bash
npm install
npx playwright install chromium
```

### 2. Configure Credentials

```bash
node scripts/setup-credentials.js
```

Creates `config/credentials.json` (gitignored) with your M365 email and password.

### 3. First Login (MFA)

```bash
xvfb-run --auto-servernum node scripts/mfa-login.js
```

When prompted for MFA, provide the 6-digit code from your Authenticator app. Auth state is saved for ~60-90 days.

### 4. Test Calendar Fetch

```bash
node scripts/calendar-fetch.js                    # Today
node scripts/calendar-fetch.js --date 2026-03-18  # Specific date
```

### 5. Submit Form

```bash
# Using an entry file
echo '{"date":"3/18/2026","trainingHours":"3",...}' > daily-entries/2026-03-18.json
node scripts/submit-daily.js

# Or directly with CLI flags
node scripts/fill-form.js --date "3/18/2026" --training 3 --content-dev 0 --learning 5
```

## Field Logic

| # | Field | Source |
|---|-------|--------|
| 1 | Date | Today (M/d/yyyy) |
| 2 | Training Hours | Training calendar events |
| 3 | Content Dev Hours | Outlook calendar (content dev blockouts) |
| 4 | Content Dev Topic | Event description from Outlook |
| 5 | Learning Hours | `max(0, 8 - training - contentDev)` |
| 6 | Learning Topic | 2-3 random topics + any KT Sessions |
| 7 | Other Items | "Preparation, testing, and rehearsal for my next class" + Outlook events (excludes KT Sessions, training calendar duplicates) |
| 8 | Team Hours | Blank |
| 9 | Team Description | Blank |

## Fallback Defaults

If calendar fetch fails, the script uses sensible defaults:

| Field | Default |
|-------|---------|
| Training Hours | 0 |
| Content Dev Hours / Topic | 0 / NA |
| Learning Hours / Topic | 8 / 3 random topics |
| Other Items | "Preparation, testing, and rehearsal for my next class" |

Each calendar is fetched independently — if one fails, the other still works.

## iCal Format Support

- UTC times (`DTSTART:20260318T110000Z`)
- Timezone-aware times (`DTSTART;TZID=Singapore Standard Time:...`)
- All-day events (`DTSTART;VALUE=DATE:20251129`)
- iCal line folding (RFC 5545)
- Supported timezones: SGT, IST, SE Asia Standard Time

## MFA Handling

This skill supports Microsoft **number-matching MFA**:
1. Script enters email + password
2. User opens Authenticator, sees a 6-digit code
3. User provides the code → script enters it on the login page
4. Auth state saved to `config/storageState.json`

When auth expires, the submission cron will ask for a fresh code.

## Configuration

| File | Purpose | Gitignored? |
|------|---------|-------------|
| `config/credentials.json` | M365 email + password | ✅ |
| `config/storageState.json` | Playwright browser session | ✅ |
| `config/form-values.json` | Legacy form value defaults | ❌ |
| `daily-entries/` | Daily submission audit trail | ✅ |

## OpenClaw Cron Jobs

Two cron jobs run the pipeline (Mon-Fri only):

| Job | Schedule | Purpose |
|-----|----------|---------|
| Daily Log Pre-Fill | 5:45 PM SGT | Fetches calendars, creates entry JSON |
| Daily Productivity Log | 6:00 PM SGT | Reads entry, submits form |

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Auth expired | Run `xvfb-run --auto-servernum node scripts/mfa-login.js` with fresh Authenticator code |
| Calendar fails | Script falls back to defaults (0/NA/8h), form still submits |
| Form URL changed | Update URL in `scripts/submit-daily.js` |
| Wrong field values | Run `node scripts/calendar-fetch.js --date YYYY-MM-DD` to debug |

## License

Private — Xaltius internal use.
