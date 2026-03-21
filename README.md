# MS Forms Auto

OpenClaw skill for automating Microsoft Forms submission with dual-calendar integration and M365 number-matching MFA support.

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

## Quick Start

### 1. Install Dependencies

```bash
cd <skill-directory>
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
# Normal cron mode (uses stored auth state)
node scripts/submit-daily.js

# If auth is expired and MFA is required, use --code flag:
node scripts/submit-daily.js --code 123456

# Or specify a date (for backfilling)
node scripts/submit-daily.js --date 2026-03-18 --code 123456
```

The `submit-daily.js` script will:
- Try existing auth state first
- If that fails and credentials-only login works, submit automatically
- If MFA is required, you must provide `--code` (MFA codes expire quickly!)
- After a successful MFA login, auth state is saved for future runs

Alternatively, use `submit-with-mfa.js` for a combined login+submit in one session:

```bash
xvfb-run --auto-servernum node scripts/submit-with-mfa.js --code 123456
```

## Field Mapping

| # | Field | Source |
|---|-------|--------|
| 1 | Date | Today (M/d/yyyy) |
| 2 | Training Hours | Training calendar events |
| 3 | Content Dev Hours | Outlook calendar (content dev blockouts) |
| 4 | Content Dev Topic | Event description from Outlook |
| 5 | Learning Hours | `max(0, 8 - training - contentDev)` |
| 6 | Learning Topic | 2-3 random topics + any KT Sessions |
| 7 | Other Items | "Preparation, testing, and rehearsal for my next class" + Outlook events (excludes KT Sessions, training duplicates) |
| 8 | Team Hours | Blank (not a team lead) |
| 9 | Team Description | Blank (not a team lead) |

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

## MFA Handling (Smart Authentication)

The `submit-with-mfa.js` script intelligently handles both MFA and non-MFA scenarios:

### Smart Behavior

1. **Auto-detection**: The script attempts credential login first, then checks if MFA is required
2. **Flexible invocation**:
   - `node scripts/submit-with-mfa.js` — Tries credentials only; if MFA appears, waits 30s for manual code entry or exits with code 2
   - `node scripts/submit-with-mfa.js --code XXXXXX` — Provides MFA code upfront; uses it if MFA appears, otherwise proceeds
3. **Graceful fallback**: If credentials alone are sufficient (no MFA prompt), it works without requiring a code
4. **Exit codes**:
   - `0` — Success
   - `1` — General error / invalid credentials
   - `2` — MFA required but no code provided (use --code flag)
   - `3` — MFA code wrong or expired
   - `4` — Form submission failed

### Recommended Usage

**For daily cron (6 PM SGT):**
```
node scripts/submit-with-mfa.js --code [MFA_CODE_HERE]
```
This ensures speed (code ready) and handles both cases:
- If MFA appears → code is used immediately
- If no MFA → code is ignored, submission proceeds

**For manual testing (unknown MFA requirement):**
```
node scripts/submit-with-mfa.js
```
The script will wait 30 seconds on the MFA screen if needed, allowing you to manually type the code from your Authenticator app.

## OpenClaw Cron Jobs

Two cron jobs run the pipeline (Mon-Fri only):

| Job | Schedule | Purpose |
|-----|----------|---------|
| Daily Log Pre-Fill | 5:45 PM SGT | Fetches calendars, creates entry JSON |
| Daily Productivity Log | 6:00 PM SGT | Reads entry, submits form |

The submit cron runs in the main session so it can request the MFA code from you. The improved smart auth ensures the submission succeeds whether MFA is required or not.

## Configuration Files

| File | Purpose | Gitignored? |
|------|---------|-------------|
| `config/credentials.json` | M365 email + password | ✅ |
| `config/storageState.json` | Playwright browser session | ✅ |
| `config/calendars.json` | Calendar URLs (contains auth tokens) | ✅ |
| `config/calendars.json.example` | Template for calendar URLs | ❌ |
| `config/credentials.json.example` | Template for credentials | ❌ |
| `config/form-values.json` | Legacy form value defaults | ❌ |
| `daily-entries/` | Daily submission audit trail | ✅ |

## Troubleshooting

- **Auth expired**: Run `xvfb-run --auto-servernum node scripts/mfa-login.js` with a fresh Authenticator code
- **Calendar fetch fails**: Script uses graceful fallbacks (0/NA defaults), form still submits
- **Form URL changed**: Update the URL in `scripts/submit-daily.js`
- **Wrong field values**: Run `node scripts/calendar-fetch.js --date YYYY-MM-DD` to debug
- **Detailed test plan**: See `references/TESTING_AUTH_WORKFLOW.md` for comprehensive authentication workflow testing scenarios

## Scripts Reference

| Script | Purpose |
|--------|---------|
| `scripts/calendar-fetch.js` | Fetch both calendars, output all form fields as JSON |
| `scripts/submit-with-mfa.js` | **Primary**: Combined MFA login + form submit in one session |
| `scripts/submit-daily.js` | Submit form using saved storageState (no MFA) |
| `scripts/mfa-login.js` | Standalone MFA login (headed mode, saves state) |
| `scripts/fill-form.js` | CLI form filler with `--date`, `--training`, etc. args |
| `scripts/setup-credentials.js` | One-time credential setup |

## License

Private — Xaltius internal use.
