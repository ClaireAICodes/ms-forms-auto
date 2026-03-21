const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const CRED_PATH = path.join(__dirname, '../config/credentials.json');
const STATE_PATH = path.join(__dirname, '../config/storageState.json');

async function simpleLogin() {
  console.log('🔐 Starting simple M365 login...');
  
  // Load credentials
  const creds = JSON.parse(fs.readFileSync(CRED_PATH, 'utf8'));
  console.log(`📧 Using account: ${creds.email}`);
  
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
  });
  const page = await context.newPage();
  
  try {
    // Navigate to the form (will redirect to login)
    console.log('🌐 Navigating to form URL...');
    await page.goto('https://forms.cloud.microsoft/r/LsxLaEv13i', { waitUntil: 'networkidle', timeout: 60000 });
    await page.waitForLoadState({ timeout: 30000 });
    
    // Debug: capture current URL and screenshot
    const currentUrl = page.url();
    console.log('📍 Current URL after navigation:', currentUrl);
    await page.screenshot({ path: path.join(__dirname, '../config/after-nav.png'), fullPage: true });
    console.log('📸 Screenshot saved to config/after-nav.png');
    
    // Check if already on form page (submit button exists)
    const submitBtn = await page.$('button[data-automation-id="submitButton"]');
    if (submitBtn) {
      console.log('✅ Already on form page - no login needed');
    } else {
      console.log('🔑 Need to login - filling credentials...');
      
      // Find email input
      const emailInput = await page.$('input[type="email"], input[name="loginfmt"], input[autocomplete="username"]');
      if (!emailInput) {
        console.log('❌ Email input not found. Available inputs:');
        const allInputs = await page.$$('input');
        for (let i = 0; i < allInputs.length; i++) {
          const el = allInputs[i];
          const attrs = await el.getAttributes();
          console.log(`  [${i}] type="${attrs.type || ''}" name="${attrs.name || ''}" placeholder="${attrs.placeholder || ''}"`);
        }
        throw new Error('Email input not found');
      }
      await emailInput.fill(creds.email);
      await emailInput.dispatchEvent('input');
      await page.waitForTimeout(1000);
      
      // Click Next
      const nextBtn = await page.$('input[type="submit"][value="Next"], input#idSIButton9');
      if (nextBtn) {
        await nextBtn.click();
        console.log('   clicked Next');
      } else {
        await page.keyboard.press('Enter');
      }
      await page.waitForLoadState({ timeout: 30000 });
      
      // Check if password field appears
      const pwdInput = await page.$('input[type="password"], input[name="passwd"]');
      if (pwdInput) {
        console.log('🔒 Password field found - entering password...');
        await pwdInput.fill(creds.password);
        await pwdInput.dispatchEvent('input');
        await page.waitForTimeout(1000);
        
        // Click Sign in
        const signInBtn = await page.$('input[type="submit"][value="Sign in"], input#idSIButton9');
        if (signInBtn) {
          await signInBtn.click();
          console.log('   clicked Sign in');
        } else {
          await page.keyboard.press('Enter');
        }
        await page.waitForLoadState({ timeout: 30000 });
      }
      
      // Check for MFA
      const mfaInput = await page.$('input[type="tel"], input[name="otc"], input[placeholder*="code" i]');
      if (mfaInput) {
        console.log('📱 MFA detected! Please enter the 6-digit code from your Authenticator app:');
        const mfaCode = await new Promise(resolve => {
          const readline = require('readline').createInterface({
            input: process.stdin,
            output: process.stdout
          });
          readline.question('', answer => {
            readline.close();
            resolve(answer.trim());
          });
        });
        await mfaInput.fill(mfaCode);
        await mfaInput.dispatchEvent('input');
        await page.waitForTimeout(1000);
        
        const submitMfa = await page.$('input[type="submit"], button[type="submit"]');
        if (submitMfa) await submitMfa.click();
        await page.waitForLoadState({ timeout: 30000 });
        console.log('✅ MFA submitted');
      }
      
      // Check for "Stay signed in?" prompt
      const staySignedBtn = await page.$('button:has-text("Yes")');
      if (staySignedBtn) {
        console.log('   "Stay signed in?" prompt detected - clicking Yes');
        await staySignedBtn.click();
        await page.waitForLoadState({ timeout: 30000 });
      }
      
      // Navigate to form again to ensure we reach it
      console.log('🎯 Navigating to form to verify login...');
      await page.goto('https://forms.cloud.microsoft/r/LsxLaEv13i', { waitUntil: 'networkidle', timeout: 60000 });
      await page.waitForLoadState({ timeout: 30000 });
    }
    
    // Verify we're on the form page
    const finalSubmitBtn = await page.$('button[data-automation-id="submitButton"]');
    if (!finalSubmitBtn) {
      throw new Error('Submit button not found - not on form page after login');
    }
    console.log('✅ Form page verified - login successful!');
    
    // Save auth state
    const state = await context.storageState();
    fs.writeFileSync(STATE_PATH, JSON.stringify(state, null, 2));
    console.log(`💾 Auth state saved to ${STATE_PATH}`);
    
  } catch (err) {
    console.error('❌ Error:', err.message);
    // Save screenshot for debugging
    await page.screenshot({ path: path.join(__dirname, '../config/login-error.png'), fullPage: true });
    console.log('📸 Screenshot saved to config/login-error.png');
  } finally {
    await browser.close();
    console.log('🔒 Browser closed');
  }
}

simpleLogin().catch(console.error);
