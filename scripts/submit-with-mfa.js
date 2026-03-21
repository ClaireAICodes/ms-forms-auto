#!/usr/bin/env node
/**
 * Smart Login + Form Submit - Handles both MFA and non-MFA scenarios.
 * 
 * USAGE:
 *   node submit-with-mfa.js --code XXXXXX [--date 2026-03-18]   # When MFA may be required
 *   node submit-with-mfa.js [--date 2026-03-18]                # Skip MFA attempt (pure credential login)
 * 
 * This script intelligently handles two scenarios:
 * 1. MFA required: Enters credentials, detects MFA prompt, enters code
 * 2. No MFA: Enters credentials and proceeds directly to form
 * 
 * All in ONE browser session — OIDC cookie stays fresh throughout.
 * 
 * Exit codes:
 *   0 = success
 *   1 = general error / credentials invalid
 *   2 = MFA required but no code provided (use --code)
 *   3 = MFA code wrong or expired
 *   4 = form submission failed
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
let requireMFAAutoDetect = true; // By default, try to auto-detect if MFA is needed

for (let i = 0; i < args.length; i++) {
  if (args[i] === '--code') {
    mfaCode = args[++i];
    requireMFAAutoDetect = false; // If code provided, we'll use it if MFA appears
  }
  if (args[i] === '--date') targetDate = args[++i];
  if (args[i] === '--force-mfa') requireMFAAutoDetect = false;
  if (args[i] === '--no-mfa') requireMFAAutoDetect = true;
}

// Validate: If we're in strict mode (no auto-detect) and no code provided, error
if (!requireMFAAutoDetect && !mfaCode) {
  console.error('❌ MFA required but no code provided. Use: node submit-with-mfa.js --code XXXXXX');
  process.exit(2);
}

async function smartLogin(page, creds) {
  console.log('🔐 Navigating to form...');
  await page.goto(FORM_URL, { waitUntil: 'networkidle', timeout: 60000 });
  await page.waitForTimeout(5000);

  // Check if already on form (already authenticated via storageState)
  const currentUrl = page.url();
  if (!currentUrl.includes('login.microsoftonline.com')) {
    console.log('✅ Already authenticated (storageState valid)');
    return true;
  }

  console.log('📧 Entering email...');
  const emailInput = await page.$('input[type="email"]');
  if (!emailInput) {
    throw new Error('Email input not found on login page');
  }
  await emailInput.fill(creds.email);
  const nextBtn = await page.$('input[type="submit"][value="Next"]');
  if (nextBtn) await nextBtn.click();
  await page.waitForTimeout(4000);

  console.log('🔑 Entering password...');
  const pwdField = await page.$('input[type="password"]');
  if (!pwdField) {
    throw new Error('Password input not found');
  }
  await pwdField.fill(creds.password);
  const signInBtn = await page.$('input[type="submit"][value="Sign in"]');
  if (signInBtn) await signInBtn.click();
  await page.waitForTimeout(3000);

  // Check for immediate error after credentials
  const pageContent = await page.content();
  if (pageContent.includes('Incorrect') || pageContent.includes('Invalid') || pageContent.includes('error') || pageContent.includes('locked')) {
    throw new Error('Login failed: Invalid email/password or account locked');
  }

  // Detect if MFA is required
  let mfaDetected = false;
  let codeInput = await page.$('input[type="tel"]');
  if (!codeInput) codeInput = await page.$('input[placeholder*="code" i]');
  if (!codeInput) codeInput = await page.$('input[name*="verification" i]');
  
  if (!codeInput) {
    // Broad search for any input that looks like verification
    const allInputs = await page.$$('input');
    for (const inp of allInputs) {
      const type = await inp.getAttribute('type') || '';
      const placeholder = await inp.getAttribute('placeholder') || '';
      const name = await inp.getAttribute('name') || '';
      if ((type === 'tel' || type === 'text' || type === 'number') && 
          (placeholder.includes('code') || placeholder.includes('verification') || placeholder.includes('digit') ||
           name.includes('verification') || name.includes('code'))) {
        codeInput = inp;
        mfaDetected = true;
        break;
      }
    }
  } else {
    mfaDetected = true;
  }

  // If MFA detected
  if (mfaDetected) {
    console.log('🔐 MFA challenge detected');
    
    if (mfaCode) {
      // Code provided - enter it
      console.log('🔢 Entering MFA code...');
      await codeInput.fill(mfaCode);
      const submitBtn = await page.$('input[type="submit"]');
      if (submitBtn) await submitBtn.click();
      await page.waitForTimeout(4000);

      // "Stay signed in?" prompt
      const yesBtn = await page.$('input[type="submit"][value="Yes"]');
      if (yesBtn) {
        await yesBtn.click();
        await page.waitForTimeout(3000);
      }

      // Check for MFA errors
      const afterMFAContent = await page.content();
      if (afterMFAContent.includes('Incorrect') || afterMFAContent.includes('Invalid') || afterMFAContent.includes('code is incorrect')) {
        throw new Error('MFA code wrong or expired');
      }
      
      await page.waitForTimeout(3000);
    } else {
      // No code available and MFA detected - we can't proceed
      console.error('❌ MFA required but no code provided. Use --code flag.');
      console.log('   Waiting 30s for manual code entry (optional)...');
      // Keep browser open for 30s in case user wants to manually enter
      await page.waitForTimeout(30000);
      
      // Check if we're still on login page
      if (page.url().includes('login.microsoftonline.com')) {
        throw new Error('MFA required but no code provided');
      }
      // If somehow we got past MFA (user manually entered), continue
    }
  } else {
    console.log('✅ No MFA prompt — login successful with credentials only');
  }

  // Final check: Are we on the form?
  if (page.url().includes('login.microsoftonline.com')) {
    throw new Error('Authentication failed: Still on login page');
  }

  console.log('✅ Authenticated! Now on form page.');
  return true;
}

async function fillAndSubmit(page, dateStr) {
  const entryFile = path.join(ENTRIES_DIR, `${dateStr}.json`);
  if (!fs.existsSync(entryFile)) {
    console.error(`❌ No entry file found: ${entryFile}`);
    return false;
  }

  const entries = JSON.parse(fs.readFileSync(entryFile, 'utf8'));
  console.log(`\n📋 Filling form for ${dateStr}...`);

  await page.waitForTimeout(3000);

  // Scroll to load all fields
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await page.waitForTimeout(2000);

  // Collect visible text inputs
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
    } else {
      console.warn(`  ${field.label} → WARNING: Input field ${field.index} not found`);
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
      
      // Save updated auth state for future reuse
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
  console.log(`🚀 MS Forms Smart Submit`);
  console.log(`   Date: ${targetDate}`);
  console.log(`   Time: ${new Date().toISOString()}\n`);

  // Check credentials exist
  if (!fs.existsSync(CREDS_FILE)) {
    console.error(`❌ Credentials file not found: ${CREDS_FILE}`);
    console.error('   Run: node scripts/setup-credentials.js');
    process.exit(1);
  }

  const creds = JSON.parse(fs.readFileSync(CREDS_FILE, 'utf8'));

  const browser = await chromium.launch({
    headless: false,
    args: ['--no-sandbox']
  });
  const page = await browser.newPage();

  try {
    // Step 1: Smart login (handles MFA or no-MFA)
    const loggedIn = await smartLogin(page, creds);
    if (!loggedIn) {
      await browser.close();
      process.exit(2);
    }

    // Step 2: Fill and submit form
    const submitted = await fillAndSubmit(page, targetDate);
    await browser.close();

    if (submitted) {
      console.log('🎉 All done!');
      process.exit(0);
    } else {
      console.log('❌ Form submission failed');
      process.exit(4);
    }

  } catch (err) {
    console.error('❌ Error:', err.message);
    await browser.close();
    process.exit(1);
  }
}

main();
