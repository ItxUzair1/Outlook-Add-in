const path = require('path');
const readline = require('readline');
const { scanDirectory } = require('./scanner');
const { parseEmailFile } = require('./parser');

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

rl.question('Please enter a folder path to scan for test emails (e.g., C:\\Users\\...\\Desktop\\test-emails): ', async (targetDir) => {
  rl.close();

  if (!targetDir || targetDir.trim() === '') {
    console.error('No path provided. Exiting.');
    return;
  }

  console.log(`\nScanning folder: ${targetDir} ...`);
  
  // 1. Scan for files
  const files = scanDirectory(targetDir.trim());
  
  console.log(`\nFound ${files.length} email files (.msg or .eml).`);
  
  if (files.length === 0) {
    console.log('No files found to test. Try another folder.');
    return;
  }

  // 2. Take the first 3 files to test parsing
  const testFiles = files.slice(0, 3);
  console.log(`\nTesting parsing on ${testFiles.length} files:\n`);

  for (let i = 0; i < testFiles.length; i++) {
    const filePath = testFiles[i];
    console.log(`--- File ${i + 1}: ${path.basename(filePath)} ---`);
    console.log(`Path: ${filePath}`);
    try {
      const emailData = await parseEmailFile(filePath);
      
      console.log(`Subject: ${emailData.subject}`);
      console.log(`Sender: ${emailData.sender}`);
      console.log(`Recipients: ${emailData.recipients}`);
      console.log(`Sent At: ${new Date(emailData.sentAt).toLocaleString()}`);
      console.log(`Has Attachments: ${emailData.hasAttachments ? '✅ Yes' : '❌ No'}`);
      console.log(`Body Snippet: ${emailData.body.substring(0, 100).replace(/\n/g, ' ')}...`);
      
    } catch (err) {
      console.error(`❌ Failed to parse: ${err.message}`);
    }
    console.log('');
  }

  console.log('✅ Test scan completed!');
});
