#!/usr/bin/env node

/**
 * Build script to inject environment variables into index.html
 * Reads from form-launcher/.env and creates form-launcher/index.html from index.template.html
 */

const fs = require('fs');
const path = require('path');

// Read .env file
const envPath = path.join(__dirname, 'form-launcher', '.env');
const templatePath = path.join(__dirname, 'form-launcher', 'index.template.html');
const outputPath = path.join(__dirname, 'form-launcher', 'index.html');

if (!fs.existsSync(envPath)) {
  console.error('❌ Error: .env file not found at form-launcher/.env');
  process.exit(1);
}

if (!fs.existsSync(templatePath)) {
  console.error('❌ Error: index.template.html not found');
  console.error('   Run: npm run create-template (if this is your first build)');
  process.exit(1);
}

// Parse .env file
const envContent = fs.readFileSync(envPath, 'utf8');
const envVars = {};

envContent.split('\n').forEach(line => {
  const trimmed = line.trim();
  if (trimmed && !trimmed.startsWith('#')) {
    const [key, ...valueParts] = trimmed.split('=');
    envVars[key.trim()] = valueParts.join('=').trim();
  }
});

// Verify required variables
const required = ['DELETE_API_KEY', 'DELETE_WEBAPP_URL', 'SHEETS_API_KEY', 'SPREADSHEET_ID'];
const missing = required.filter(key => !envVars[key] || envVars[key] === 'YOUR_WEB_APP_URL_HERE');

if (missing.length > 0) {
  console.warn('⚠️  Warning: The following variables need to be set in .env:');
  missing.forEach(key => console.warn(`   - ${key}`));
}

// Read template
let html = fs.readFileSync(templatePath, 'utf8');

// Replace placeholders
html = html.replace(/{{SPREADSHEET_ID}}/g, envVars.SPREADSHEET_ID || '');
html = html.replace(/{{SHEETS_API_KEY}}/g, envVars.SHEETS_API_KEY || '');
html = html.replace(/{{DELETE_WEBAPP_URL}}/g, envVars.DELETE_WEBAPP_URL || '');
html = html.replace(/{{DELETE_API_KEY}}/g, envVars.DELETE_API_KEY || '');

// Write output
fs.writeFileSync(outputPath, html, 'utf8');

console.log('✅ Build complete!');
console.log(`   Template: ${path.basename(templatePath)}`);
console.log(`   Output:   ${path.basename(outputPath)}`);
console.log(`   Variables injected: ${Object.keys(envVars).length}`);
