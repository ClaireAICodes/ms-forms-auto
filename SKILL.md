---
name: ms-forms-auto
description: Automate Microsoft Forms submission using Playwright with M365 auto-login. Use when: automating daily/regular MS Forms reports, submitting responses to Microsoft Forms programmatically, setting up scheduled form submissions, or working with org MS Forms that require authentication. Triggers on phrases like "fill out MS form", "automate form submission", "daily form report", "submit Microsoft form", "org daily log".
---

# MS Forms Auto

Automate Microsoft Forms submission with headless Playwright and M365 auto-login.

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

This interactively prompts for M365 email/password and saves to `config/credentials.json` (gitignored).

### 3. Configure Form Values

Edit `config/form-values.json` with your daily default answers. Structure:

```json
{
  "defaults": {
    "<question-id-or-ordinal>": "<value>"
  },
  "perDate": {
    "2026-03-18": {
      "<question-id>": "<override>"
    }
  }
}
```

- **defaults**: Used when no per-date override exists
- **perDate**: Specific overrides for particular dates (checked first)

### 4. Test Submission

```bash
node scripts/submit.js --dry-run    # Preview values without submitting
node scripts/submit.js              # Submit for real
```

### 5. Schedule (cron)

```bash
# Daily at 5:30 PM SGT (09:30 UTC)
node scripts/submit.js --cron "30 9 * * 1-5"
```

Or use OpenClaw cron:

```json
{
  "name": "MS Forms Daily Report",
  "schedule": { "kind": "cron", "expr": "30 9 * * 1-5", "tz": "Asia/Singapore" },
  "payload": { "kind": "agentTurn", "message": "Submit the daily MS Form report using ms-forms-auto skill. Check config/form-values.json for today's values." },
  "sessionTarget": "isolated"
}
```

## How It Works

1. **Auth**: Launches headless Chromium, loads saved session from `config/auth-state.json`
2. **Login**: If session expired, auto-login with credentials from `config/credentials.json`
3. **Navigate**: Opens the form URL from `config/form-structure.json`
4. **Fill**: Matches question IDs to values from `config/form-values.json`
5. **Submit**: Clicks submit button, waits for confirmation
6. **Save**: Persists updated auth state for next run

## Auth State Management

- `config/auth-state.json` — saved Playwright browser session (cookies, localStorage)
- Sessions typically last days-weeks depending on org M365 policy
- When expired, auto-login refreshes it transparently
- If MFA is required, script will prompt for interactive approval

## Configuration Files

| File | Purpose | Gitignored? |
|------|---------|-------------|
| `config/credentials.json` | M365 email + password | ✅ Yes |
| `config/auth-state.json` | Saved browser session | ✅ Yes |
| `config/form-values.json` | Daily answer values | ❌ No (no secrets) |
| `config/form-structure.json` | Form URL + question IDs | ❌ No |

## Adding New Forms

1. Run: `node scripts/extract-form.js <form-url>` to discover question IDs
2. Add form structure to `config/form-structure.json`
3. Add default values to `config/form-values.json`
4. Run with `--form <form-name>` to target specific form

## Troubleshooting

- **Login fails**: Delete `config/auth-state.json` and re-run (forces fresh login)
- **MFA required**: Run with `--headed` flag to approve MFA interactively, then session saves
- **Form structure changed**: Re-run `extract-form.js` to update question IDs
- **Values not matching**: Check question IDs in `form-structure.json` match `form-values.json`
