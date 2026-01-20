# Testing Guide

## Setup

1. **Copy the example environment file:**
   ```bash
   cp .env.example .env
   ```

2. **Edit `.env` with your AWS WorkMail credentials:**
   ```bash
   # Open .env in your editor and fill in:
   EWS_USERNAME=your-email@example.com
   EWS_PASSWORD=your-password
   EWS_URL=https://ews.mail.{region}.awsapps.com/EWS/Exchange.asmx
   ```

3. **Install dependencies (if not done already):**
   ```bash
   npm install
   ```

## Running Tests

### Interactive Menu

Run the interactive test suite:
```bash
node test.js
```

The interactive menu allows you to:
- Test individual operations (messages, folders, events, contacts, etc.)
- Run all tests at once
- Verify your connection and credentials

### Menu Options

```
[1]   Test Connection          - Verify AWS WorkMail connection
[2]   Get Messages             - Retrieve messages from Inbox
[3]   Send Message             - Send a test email
[4]   Reply to Message         - Reply to an existing message
[5]   Create Folder            - Create a test folder
[6]   Get Folders              - List all folders
[7]   Create Event             - Create a calendar event
[8]   Get Events               - List calendar events
[9]   Create Contact           - Create a test contact
[10]  Get Contacts             - List all contacts
[11]  Add Attachment           - Add attachment to message
[12]  Get Calendars            - Get calendar information

[all] Run ALL Tests            - Execute all tests sequentially
[q]   Quit                     - Exit the test suite
```

## Important Notes

### Credentials Security

⚠️ **NEVER commit your `.env` file!**

The `.env` file contains your password and is automatically excluded from git (via `.gitignore`).

### Test Email Recipient

By default, test emails are sent to your own email address (EWS_USERNAME). You can override this by setting `EWS_TEST_EMAIL` in `.env`:

```bash
EWS_TEST_EMAIL=another-email@example.com
```

### Cleanup

Most tests that create resources (folders, events, contacts) automatically clean up after themselves. Test messages are sent to Drafts and then deleted.

### Reply to Message Test

The "Reply to Message" test will only work if you have messages in your Inbox from OTHER senders. It cannot reply to messages sent by yourself (Exchange limitation).

## Troubleshooting

### "Missing configuration" Error

Make sure your `.env` file exists and contains all required fields:
- `EWS_USERNAME`
- `EWS_PASSWORD`
- `EWS_URL`

### "403 Forbidden" Error

Check:
1. Your credentials are correct
2. Your IP is not blocked by AWS WorkMail Access Control Rules
3. Basic Authentication is enabled for your organization

### "The specified server version is invalid"

Ensure your EWS URL uses the correct format:
```
https://ews.mail.{region}.awsapps.com/EWS/Exchange.asmx
```
(NOT `mobile.mail.`)

## Running Specific Tests Programmatically

You can also use the test suite programmatically:

```javascript
const { tests, runTest } = require('./test.js');

// Run a specific test
await runTest('Get Messages', tests.testGetMessages);
```
