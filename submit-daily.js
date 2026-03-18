#!/usr/bin/env node
/**
 * Auto-fill Daily Productivity Log
 * 
 * USAGE: node submit-daily.js
 * 
 * Requires:
 * - config/storageState.json (valid MS auth - run MFA login first if expired)
 * 
 * This script fills the form with values from daily-entries/ directory.
 * If no entry file exists for today, it prompts via the OpenClaw cron system.
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const CONFIG_DIR = path.join(__dirname, 'config');
const AUTH_STATE = path.join(CONFIG_DIR, 'storageState.json');
const ENTRIES_DIR = path.join(__dirname, 'daily-entries');
const FORM_URL = 'https://forms.cloud.microsoft/r/LsxLaEv13i';

async function fillForm(entries) {
  console.log(`🚀 Filling Daily Productivity Log — ${entries.date}\n`);
  
  if (!fs.existsSync(AUTH_STATE)) {
    console.error('❌ No auth state found. Need MFA login first.');
    process.exit(1);
  }
  
  const browser = await chromium.launch({ headless: true, args: ['--no-sandbox'] });
  const context = await browser.newContext({ storageState: AUTH_STATE });
  const page = await context.newPage();
  
  try {
    await page.goto(FORM_URL, { waitUntil: 'networkidle', timeout: 60000 });
    await page.waitForTimeout(3000);
    
    // Scroll to load all fields
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await page.waitForTimeout(2000);
    
    if (page.url().includes('login.microsoftonline.com')) {
      console.error('❌ Auth expired! Need MFA re-authentication.');
      console.log('TIP: Run the MFA login script with xvfb-run to refresh auth.');
      await browser.close();
      return { success: false, error: 'auth_expired' };
    }
    
    console.log('✅ Form loaded\n');
    
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
    
    const fields = [
      { index: 0, value: entries.date,          label: 'Date' },
      { index: 1, value: entries.trainingHours,  label: 'Training Hours' },
      { index: 2, value: entries.contentDevHours, label: 'Content Dev Hours' },
      { index: 3, value: entries.contentDevTopic, label: 'Content Dev Topic' },
      { index: 4, value: entries.learningHours,   label: 'Learning Hours' },
      { index: 5, value: entries.learningTopic,   label: 'Learning Topic' },
      { index: 6, value: entries.otherItems,      label: 'Other Items' },
      { index: 7, value: entries.teamHours || '',  label: 'Team Hours' },
      { index: 8, value: entries.teamDesc || '',   label: 'Team Description' },
    ];
    
    for (const field of fields) {
      if (!field.value) {
        console.log(`  ${field.label} → SKIPPED`);
        continue;
      }
      if (field.index < textInputs.length) {
        await textInputs[field.index].click();
        await textInputs[field.index].fill(field.value);
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
      console.log('\n🟢 Submitting...');
      await submitBtn.click();
      await page.waitForTimeout(8000);
      
      const pageText = await page.evaluate(() => document.body.innerText);
      if (pageText.includes('Your response was submitted') || pageText.includes('submitted')) {
        console.log('✅ SUBMITTED SUCCESSFULLY!\n');
        await context.storageState({ path: AUTH_STATE });
        await browser.close();
        return { success: true };
      }
      
      // Check for validation errors
      if (pageText.includes('required') || pageText.includes('error')) {
        console.log('⚠️ Validation error:', pageText.substring(0, 500));
        await browser.close();
        return { success: false, error: 'validation' };
      }
      
      console.log('⚠️ Unclear status. Page:', pageText.substring(0, 200));
      await browser.close();
      return { success: false, error: 'unclear' };
    }
    
    await browser.close();
    return { success: false, error: 'no_submit_button' };
    
  } catch (err) {
    console.error('❌ Error:', err.message);
    await browser.close();
    return { success: false, error: err.message };
  }
}

// If run directly, read today's entry file
if (require.main === module) {
  const today = new Date().toISOString().split('T')[0];
  const entryFile = path.join(ENTRIES_DIR, `${today}.json`);
  
  if (!fs.existsSync(entryFile)) {
    console.log(`No entry file for today: ${entryFile}`);
    console.log('Usage: Create a JSON file in daily-entries/ with your work data');
    process.exit(0);
  }
  
  const entries = JSON.parse(fs.readFileSync(entryFile, 'utf8'));
  fillForm(entries).then(result => {
    if (!result.success) {
      process.exit(1);
    }
  });
}

module.exports = { fillForm };
