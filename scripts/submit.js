#!/usr/bin/env node
/**
 * MS Forms Auto-Submitter
 * 
 * Submits Microsoft Forms responses using Playwright with M365 auto-login.
 * 
 * Usage:
 *   node submit.js                     Submit for today (real)
 *   node submit.js --dry-run           Preview values without submitting
 *   node submit.js --date 2026-03-17   Submit for specific date
 *   node submit.js --headed            Run with visible browser (for MFA)
 *   node submit.js --form daily-report Target specific form by name
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const CONFIG_DIR = path.join(__dirname, '..', 'config');
const AUTH_STATE_PATH = path.join(CONFIG_DIR, 'auth-state.json');
const CREDENTIALS_PATH = path.join(CONFIG_DIR, 'credentials.json');
const FORM_VALUES_PATH = path.join(CONFIG_DIR, 'form-values.json');
const FORM_STRUCTURE_PATH = path.join(CONFIG_DIR, 'form-structure.json');

// --- CLI Args ---
function parseArgs() {
  const args = process.argv.slice(2);
  return {
    dryRun: args.includes('--dry-run'),
    headed: args.includes('--headed'),
    date: args.find(a => a.startsWith('--date='))?.split('=')[1] || todayStr(),
    form: args.find(a => a.startsWith('--form='))?.split('=')[1] || null,
    help: args.includes('--help') || args.includes('-h'),
  };
}

function todayStr() {
  // SGT (UTC+8)
  const now = new Date();
  const sgt = new Date(now.getTime() + 8 * 60 * 60 * 1000);
  return sgt.toISOString().split('T')[0];
}

// --- Config Loaders ---
function loadJSON(filePath, required = true) {
  if (!fs.existsSync(filePath)) {
    if (required) {
      console.error(`❌ Missing required file: ${filePath}`);
      console.error(`   Run setup-credentials.js first or create the config file.`);
      process.exit(1);
    }
    return null;
  }
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function loadFormStructure(formName = null) {
  const structures = loadJSON(FORM_STRUCTURE_PATH);
  if (formName) {
    const form = structures.forms?.find(f => f.name === formName);
    if (!form) {
      console.error(`❌ Form "${formName}" not found in ${FORM_STRUCTURE_PATH}`);
      console.error(`   Available: ${structures.forms?.map(f => f.name).join(', ')}`);
      process.exit(1);
    }
    return form;
  }
  // Return first form (or default)
  return structures.default ? structures.forms?.find(f => f.name === structures.default) : structures.forms?.[0];
}

function loadValues(formName, date) {
  const values = loadJSON(FORM_VALUES_PATH);
  const formValues = values[formName] || values;
  
  // Start with defaults, overlay per-date overrides
  const defaults = formValues.defaults || {};
  const dateOverrides = formValues.perDate?.[date] || {};
  
  return { ...defaults, ...dateOverrides };
}

// --- M365 Login ---
async function ensureLoggedIn(page, credentials, dryRun) {
  // Check if we're already logged in by navigating to M365
  console.log('🔐 Checking authentication status...');
  
  await page.goto('https://forms.cloud.microsoft', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);
  
  // If we see a sign-in button, we need to log in
  const signInButton = await page.$('a[href*="signin"], button:has-text("Sign in"), [data-automation-id="signInButton"]');
  
  if (signInButton) {
    console.log('🔑 Session expired, logging in with credentials...');
    
    if (dryRun) {
      console.log('   [DRY RUN] Would log in with:', credentials.email);
      return;
    }
    
    // Click sign in
    await signInButton.click();
    await page.waitForTimeout(2000);
    
    // Enter email
    const emailInput = await page.waitForSelector('input[type="email"], input[name="loginfmt"]', { timeout: 10000 });
    await emailInput.fill(credentials.email);
    
    // Click Next
    const nextBtn = await page.$('input[type="submit"][value="Next"], button:has-text("Next")');
    if (nextBtn) await nextBtn.click();
    await page.waitForTimeout(3000);
    
    // Enter password
    const passwordInput = await page.waitForSelector('input[type="password"], input[name="passwd"]', { timeout: 10000 });
    await passwordInput.fill(credentials.password);
    
    // Click Sign In
    const signInBtn = await page.$('input[type="submit"][value="Sign in"], button:has-text("Sign in")');
    if (signInBtn) await signInBtn.click();
    await page.waitForTimeout(5000);
    
    // Check for MFA prompt
    const mfaPrompt = await page.$('text=Approve sign-in request, text=Enter code, text=More information required');
    if (mfaPrompt) {
      console.log('⚠️  MFA required! Waiting 60 seconds for approval...');
      console.log('   Check your phone and approve the sign-in request.');
      await page.waitForTimeout(60000);
    }
    
    // Check for "Stay signed in?" prompt
    const staySignedIn = await page.$('input[value="Yes"], button:has-text("Yes")');
    if (staySignedIn) {
      await staySignedIn.click();
      await page.waitForTimeout(2000);
    }
    
    console.log('✅ Login successful!');
    
    // Save auth state
    await page.context().storageState({ path: AUTH_STATE_PATH });
    console.log('💾 Auth state saved.');
  } else {
    console.log('✅ Already logged in (session valid).');
  }
}

// --- Form Filling ---
async function fillAndSubmit(page, form, values, dryRun) {
  console.log(`\n📝 Navigating to form: ${form.name}`);
  console.log(`   URL: ${form.url}`);
  
  await page.goto(form.url, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);
  
  // Wait for form to load
  await page.waitForSelector('[data-automation-id="questionItem"]', { timeout: 15000 });
  
  const questions = form.questions || [];
  let filledCount = 0;
  
  for (const q of questions) {
    const value = values[q.id] ?? values[String(q.ordinal)] ?? null;
    
    if (value === null || value === undefined) {
      if (q.required) {
        console.log(`   ⚠️  Q${q.ordinal} (${q.label}): REQUIRED but no value provided!`);
      } else {
        console.log(`   ⏭️  Q${q.ordinal} (${q.label}): skipped (optional, no value)`);
      }
      continue;
    }
    
    console.log(`   📌 Q${q.ordinal} (${q.label}): "${value}"`);
    
    if (dryRun) {
      filledCount++;
      continue;
    }
    
    try {
      // Find input by question ID
      const inputSelector = `#${q.questionId} input[data-automation-id="textInput"]`;
      const input = await page.waitForSelector(inputSelector, { timeout: 5000 });
      
      // Handle date picker vs text input
      if (q.type === 'date') {
        // Clear and type date in M/d/yyyy format
        await input.click();
        await input.fill('');
        // Convert YYYY-MM-DD to M/d/yyyy
        const [yyyy, mm, dd] = value.split('-');
        const dateStr = `${parseInt(mm)}/${parseInt(dd)}/${yyyy}`;
        await input.type(dateStr, { delay: 50 });
        // Press Tab to confirm
        await page.keyboard.press('Tab');
      } else {
        await input.click();
        await input.fill('');
        await input.type(String(value), { delay: 30 });
      }
      
      filledCount++;
      await page.waitForTimeout(300);
    } catch (err) {
      console.log(`   ❌ Failed to fill Q${q.ordinal}: ${err.message}`);
    }
  }
  
  console.log(`\n   Filled ${filledCount}/${questions.filter(q => values[q.id] || values[String(q.ordinal)]).length} questions with values`);
  
  if (dryRun) {
    console.log('\n🏁 DRY RUN — Skipping submission.');
    return { submitted: false, filled: filledCount };
  }
  
  // Click Submit button
  console.log('\n🚀 Submitting form...');
  try {
    const submitBtn = await page.waitForSelector('button[data-automation-id="submitButton"]', { timeout: 5000 });
    await submitBtn.click();
    await page.waitForTimeout(5000);
    
    // Check for success (usually shows "Your response was submitted" or similar)
    const successMsg = await page.$('text=Your response was submitted, text=has been recorded, text=submitted successfully');
    if (successMsg) {
      console.log('✅ Form submitted successfully!');
    } else {
      // Check for errors
      const errorMsg = await page.$('[data-automation-id="error"], .error-message');
      if (errorMsg) {
        const errorText = await errorMsg.textContent();
        console.log(`⚠️  Submission may have issues: ${errorText}`);
      } else {
        console.log('✅ Submission completed (no error detected).');
      }
    }
    
    return { submitted: true, filled: filledCount };
  } catch (err) {
    console.error(`❌ Submit failed: ${err.message}`);
    return { submitted: false, filled: filledCount, error: err.message };
  }
}

// --- Main ---
async function main() {
  const args = parseArgs();
  
  if (args.help) {
    console.log(`
MS Forms Auto-Submitter

Usage:
  node submit.js                     Submit for today
  node submit.js --dry-run           Preview without submitting
  node submit.js --date=2026-03-17   Specific date
  node submit.js --headed            Visible browser (for MFA)
  node submit.js --form=daily-report Specific form
`);
    process.exit(0);
  }
  
  console.log('🤖 MS Forms Auto-Submitter');
  console.log(`   Date: ${args.date}`);
  console.log(`   Mode: ${args.dryRun ? 'DRY RUN' : 'LIVE'}`);
  console.log(`   Browser: ${args.headed ? 'headed' : 'headless'}`);
  
  // Load configs
  const credentials = loadJSON(CREDENTIALS_PATH);
  const form = loadFormStructure(args.form);
  
  if (!form) {
    console.error('❌ No form configured. Add forms to config/form-structure.json');
    process.exit(1);
  }
  
  const values = loadValues(form.name, args.date);
  console.log(`   Form: ${form.name}`);
  console.log(`   Values: ${JSON.stringify(values, null, 2).split('\n').length - 1} fields\n`);
  
  // Launch browser
  const launchOptions = {
    headless: !args.headed,
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  };
  
  let context;
  
  // Try to use saved auth state
  if (fs.existsSync(AUTH_STATE_PATH)) {
    console.log('📂 Loading saved auth state...');
    context = await chromium.launchPersistentContext(
      path.join(CONFIG_DIR, '.browser-profile'),
      { ...launchOptions, storageState: AUTH_STATE_PATH }
    );
  } else {
    const browser = await chromium.launch(launchOptions);
    context = await browser.newContext();
  }
  
  const page = context.pages()[0] || await context.newPage();
  
  try {
    // Ensure logged in
    await ensureLoggedIn(page, credentials, args.dryRun);
    
    // Fill and submit
    const result = await fillAndSubmit(page, form, values, args.dryRun);
    
    // Save updated auth state
    if (!args.dryRun && result.submitted) {
      await context.storageState({ path: AUTH_STATE_PATH });
      console.log('💾 Auth state updated.');
    }
    
    console.log('\n✨ Done!');
    process.exit(result.submitted || args.dryRun ? 0 : 1);
  } catch (err) {
    console.error(`\n💥 Fatal error: ${err.message}`);
    process.exit(1);
  } finally {
    await context.close();
  }
}

main();
