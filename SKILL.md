---
name: ms-forms-auto
description: Automate Microsoft Forms submission with dual-calendar integration and M365 MFA support. Use when: automating daily productivity log submissions, submitting Microsoft Forms with calendar-based auto-fill, setting up scheduled form submissions with training/content-dev/learning hour calculations, or working with org MS Forms that require M365 number-matching MFA. Triggers on phrases like "fill out MS form", "automate form submission", "daily form report", "submit Microsoft form", "daily productivity log", "training hours form".
---

# MS Forms Auto

Automate Microsoft Forms submission with Playwright, dual-calendar auto-fill, and M365 number-matching MFA support.

## Quick Start

### 1. Install Dependencies

```bash
cd <skill-directory>
npm install
npx playwright install chromium
```

### 2. Set Up Credentials (one-time)

```bash
node scripts/setup-credentials.js
```

Creates `config/credentials.json` (gitignored) with M365 email/password.

### 3. First Login with MFA

```bash
xvfb-run --auto-servernum node scripts/mfa-login.js
```

When prompted, enter the 6-digit code from your Authenticator app. Saves auth state to `config/storageState.json` for headless reuse.

### 4. Fetch Calendar Data

```bash
node scripts/calendar-fetch.js              # Today (SGT)
node scripts/calendar-fetch.js --date 2026-03-18  # Specific date
```

Returns JSON with all form fields auto-populated from both calendars.

### 5. Submit the Form

```bash
node scripts/submit-daily.js
```

Reads today's entry from `daily-entries/YYYY-MM-DD.json` and submits.

## Architecture

### Dual Calendar System

| Calendar | URL Source | Purpose |
|----------|-----------|---------|
| **Training (TMS)** | Built into script | Training Hours (source of truth) |
| **Outlook** | Built into script | Content Dev Hours, Other Items |

### Field Logic

| # | Field | Source |
|---|-------|--------|
| 1 | Date | Today (M/d/yyyy) |
| 2 | Training Hours | Training calendar events |
| 3 | Content Dev Hours | Outlook calendar (content dev blockouts) |
| 4 | Content Dev Topic | Event description from Outlook |
| 5 | Learning Hours | `max(0, 8 - training - contentDev)` |
| 6 | Learning Topic | 2-3 random topics + any KT Sessions |
| 7 | Other Items | "Preparation, testing, and rehearsal for my next class" + Outlook events (excludes KT Sessions, training duplicates) |
| 8 | Team Hours | Blank |
| 9 | Team Description | Blank |

### Fallback Defaults (if calendars fail)

| Field | Fallback |
|-------|----------|
| Training Hours | 0 |
| Content Dev Hours | 0 |
| Content Dev Topic | NA |
| Learning Hours | 8 |
| Learning Topic | 3 random topics from pool |
| Other Items | "Preparation, testing, and rehearsal for my next class" |

## Scripts

| Script | Purpose |
|--------|---------|
| `scripts/calendar-fetch.js` | Fetch both calendars, output all form fields |
| `scripts/submit-daily.js` | Submit form using entry JSON file |
| `scripts/mfa-login.js` | Interactive MFA login (headed mode) |
| `scripts/fill-form.js` | CLI form filler with `--date`, `--training`, etc. args |
| `scripts/setup-credentials.js` | One-time credential setup |

## Automated Pipeline (via OpenClaw Cron)

```
5:45 PM SGT → Pre-Fill cron → calendar-fetch.js → entry JSON
6:00 PM SGT → Submit cron → submit-daily.js → form submitted
```

Both cron jobs run Mon-Fri only.

## Configuration Files

| File | Purpose | Gitignored? |
|------|---------|-------------|
| `config/credentials.json` | M365 email + password | ✅ |
| `config/storageState.json` | Saved browser session (cookies) | ✅ |
| `config/form-values.json` | Legacy form value defaults | ❌ |
| `daily-entries/` | Daily submission audit trail | ✅ |

## MFA Handling

This skill supports Microsoft number-matching MFA:
1. Enter email + password
2. Authenticator app shows a 6-digit code
3. User provides the code → entered into login page
4. Auth state saved for ~60-90 days

When auth expires (exit code 2), run `mfa-login.js` to refresh.

## Troubleshooting

- **Auth expired**: Run `xvfb-run --auto-servernum node scripts/mfa-login.js` with a fresh Authenticator code
- **Calendar fetch fails**: Script uses graceful fallbacks (0/NA defaults), form still submits
- **Form URL changed**: Update the URL in `scripts/submit-daily.js`
- **Wrong field values**: Run `node scripts/calendar-fetch.js --date YYYY-MM-DD` to verify calendar data
