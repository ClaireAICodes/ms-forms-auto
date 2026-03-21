#!/bin/bash
set -euo pipefail

# Change to skill directory
cd /home/ubuntu/.openclaw/workspace/skills/ms-forms-auto

# Run submission
echo "🚀 Submitting Daily Productivity Log..."
xvfb-run --auto-servernum node scripts/submit-with-mfa.js 2>&1 | tee /tmp/submit-output.log
exit_code=${PIPESTATUS[0]}

if [ $exit_code -ne 0 ]; then
  echo "❌ Submission failed with exit code $exit_code"
  exit $exit_code
fi

echo "✅ Submission succeeded"

# Get today's date in Asia/Singapore
date=$(TZ=Asia/Singapore date +%Y-%m-%d)
echo "📅 Date: $date"

# Find latest audit screenshot
screenshot=$(ls -t audit-screenshots/submitted-${date}-*.png 2>/dev/null | head -1 || true)
if [ -z "$screenshot" ]; then
  echo "⚠️ No audit screenshot found for $date"
  # Not a failure, continue
else
  echo "📸 Found screenshot: $screenshot"
  
  # Get email from credentials
  if [ -f config/credentials.json ]; then
    email=$(node -p "require('./config/credentials.json').email" 2>/dev/null || echo "")
  else
    email=""
  fi
  
  if [ -z "$email" ]; then
    echo "⚠️ Could not read email from credentials.json"
  else
    echo "📧 Sending email to $email..."
    # Check if gog is available
    if command -v gog &>/dev/null; then
      # Try to send email (non-interactive)
      if GOG_NONINTERACTIVE=1 gog gmail send --account "$email" --to "$email" \
          --subject "Daily Productivity Log - $date" \
          --body "The Daily Productivity Log for $date was submitted successfully. See attached audit screenshot." \
          --attachment "$screenshot" --no-input 2>/dev/null; then
        echo "✅ Email sent"
      else
        echo "⚠️ gog email send failed (non-zero exit). Continuing."
      fi
    else
      echo "⚠️ gog command not found. Email not sent."
    fi
  fi
fi

# Final success exit
exit 0
