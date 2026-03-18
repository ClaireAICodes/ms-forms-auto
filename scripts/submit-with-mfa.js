#!/usr/bin/env node
/**
 * Combined MFA Login + Form Submit in one browser session.
 * 
 * USAGE: node submit-with-mfa.js --code XXXXXX [--date 2026-03-18]
 * 
 * This script:
 * 1. Launches headed browser (xvfb-run)
 * 2. Navigates to the form → triggers MFA login flow
 * 3. Enters the 6-digit MFA code immediately (before it expires)
 * 4. Once authenticated, fills and submits the form
 * 5. Saves updated storageState for potential reuse
 * 
 * All in ONE session — the OIDC cookie stays fresh throughout.
 * 
 * Exit codes:
 *   0 = success
 *   1 = general error
 *   2 = MFA code wrong or expired
 *   3 = form submission failed
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(path.join(__dirname, '..'));
const CONFIG_DIR = path.join(ROOT_DIR, 'config');
const CREDS_FILE = path.join(CONFIG_DIR, 'credentials.json');
const AUTH_STATE = path.join(CONFIG_DIR, 'storageState.json');
const ENTRIES_DIR = path.join(ROOT_DIR, 'daily-entries');
const FORM_URL = 'https://forms.cloud.microsoft/r/LsxLaEv13i';

// Parse args
const args = process.argv.slice(2);
let mfaCode = null;
let targetDate = new Date().toISOString().split('T')[0];

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--code') mfaCode = args[++i];
  if (args[i] === '--date') targetDate = args[++i];
}

if (!mfaCode) {
  console.error('Usage: node submit-with-mfa.js --code XXXXXX [--date YYYY-MM-DD]');
  process.exit(1);
}

async function loginWithMFA(page, code) {
  const creds = JSON.parse(fs.readFileSync(CREDS_FILE, 'utf8'));

  console.log('🔐 Navigating to form (triggers MFA login)...');
  await page.goto(FORM_URL, { waitUntil: 'networkidle', timeout: 60000 });
  await page.waitForTimeout(5000);

  // If already on login page
  if (page.url().includes('login.microsoftonline.com')) {
    console.log('📧 Entering email...');
    await page.fill('input[type="email"]', creds.email);
    const nextBtn = await page.$('input[type="submit"][value="Next"]');
    if (nextBtn) await nextBtn.click();
    await page.waitForTimeout(4000);

    console.log('🔑 Entering password...');
    const pwdField = await page.$('input[type="password"]');
    if (pwdField) {
      await pwdField.fill(creds.password);
      const signInBtn = await page.$('input[type="submit"][value="Sign in"]');
      if (signInBtn) await signInBtn.click();
    }
    await page.waitForTimeout(3000);
  }

  // MFA challenge
  const codeInput = await page.$('input[type="tel"]');
  if (codeInput) {
    console.log('🔢 Entering MFA code...');
    await codeInput.fill(code);
    const submitBtn = await page.$('input[type="submit"]');
    if (submitBtn) await submitBtn.click();
    await page.waitForTimeout(4000);

    // Stay signed in?
    const yesBtn = await page.$('input[type="submit"][value="Yes"]');
    if (yesBtn) {
      await yesBtn.click();
      await page.waitForTimeout(3000);
    }
    
    // Wait for redirect back to form
    await page.waitForTimeout(3000);
  } else {
    console.log('ℹ️  No MFA prompt — may already be authenticated');
  }

  // Check if we landed on the form
  if (page.url().includes('login.microsoftonline.com')) {
    console.error('❌ Still on login page — MFA code may be wrong or expired');
    return false;
  }

  console.log('✅ Authenticated! Now on form page.');
  return true;
}

async function fillAndSubmit(page, dateStr) {
  // Load entry file
  const entryFile = path.join(ENTRIES_DIR, `${dateStr}.json`);
  if (!fs.existsSync(entryFile)) {
    console.error(`❌ No entry file found: ${entryFile}`);
    return false;
  }

  const entries = JSON.parse(fs.readFileSync(entryFile, 'utf8'));
  console.log(`\n📋 Filling form for ${dateStr}...`);

  // Wait for form to fully load
  await page.waitForTimeout(3000);

  // Scroll to load all lazy-rendered fields
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await page.waitForTimeout(2000);

  // Collect all visible text inputs
  const allInputs = await page.$$('input');
  const textInputs = [];
  for (const inp of allInputs) {
    const type = await inp.getAttribute('type');
    const visible = await inp.isVisible();
    if (visible && (!type || type === 'text')) {
      textInputs.push(inp);
    }
  }

  console.log(`Found ${textInputs.length} input fields`);

  const fields = [
    { index: 0, value: entries.date,            label: 'Date' },
    { index: 1, value: entries.trainingHours,    label: 'Training Hours' },
    { index: 2, value: entries.contentDevHours,  label: 'Content Dev Hours' },
    { index: 3, value: entries.contentDevTopic,  label: 'Content Dev Topic' },
    { index: 4, value: entries.learningHours,    label: 'Learning Hours' },
    { index: 5, value: entries.learningTopic,    label: 'Learning Topic' },
    { index: 6, value: entries.otherItems,       label: 'Other Items' },
    { index: 7, value: entries.teamHours || '',   label: 'Team Hours' },
    { index: 8, value: entries.teamDesc || '',    label: 'Team Description' },
  ];

  for (const field of fields) {
    if (!field.value) {
      console.log(`  ${field.label} → SKIPPED (empty)`);
      continue;
    }
    if (field.index < textInputs.length) {
      await textInputs[field.index].click();
      await textInputs[field.index].fill(String(field.value));
      const placeholder = await textInputs[field.index].getAttribute('placeholder');
      if (placeholder?.includes('number')) await page.keyboard.press('Tab');
      console.log(`  ${field.label} → "${field.value}"`);
      await page.waitForTimeout(200);
    }
  }

  await page.waitForTimeout(1000);

  // Submit
  const submitBtn = await page.$('button[data-automation-id="submitButton"]');
  if (submitBtn) {
    console.log('\n🟢 Submitting form...');
    await submitBtn.click();
    await page.waitForTimeout(8000);

    const pageText = await page.evaluate(() => document.body.innerText);
    if (pageText.includes('Your response was submitted') || pageText.includes('submitted')) {
      console.log('✅ FORM SUBMITTED SUCCESSFULLY!\n');
      
      // Save updated auth state
      const context = page.context();
      await context.storageState({ path: AUTH_STATE });
      console.log('💾 Auth state saved for potential reuse.');
      return true;
    }

    console.error('⚠️  Unclear result. Page text:', pageText.substring(0, 300));
    return false;
  }

  console.error('❌ Submit button not found');
  return false;
}

async function main() {
  console.log(`🚀 MS Forms Auto-Submit with MFA`);
  console.log(`   Date: ${targetDate}`);
  console.log(`   Time: ${new Date().toISOString()}\n`);

  const browser = await chromium.launch({
    headless: false,  // Required for MFA
    args: ['--no-sandbox']
  });
  const page = await browser.newPage();

  try {
    // Step 1: Login with MFA
    const loggedIn = await loginWithMFA(page, mfaCode);
    if (!loggedIn) {
      await browser.close();
      process.exit(2);
    }

    // Step 2: Fill and submit form (already on form page)
    const submitted = await fillAndSubmit(page, targetDate);
    await browser.close();

    if (submitted) {
      console.log('🎉 All done!');
      process.exit(0);
    } else {
      console.log('❌ Form submission failed');
      process.exit(3);
    }

  } catch (err) {
    console.error('❌ Error:', err.message);
    await browser.close();
    process.exit(1);
  }
}

main();
