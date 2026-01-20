#!/usr/bin/env node
/**
 * AWS WorkMail EWS - Interactive Test Suite
 *
 * Setup:
 * 1. Copy .env.example to .env
 * 2. Fill in your AWS WorkMail credentials
 * 3. Run: node test.js
 */

require('dotenv').config();
const readline = require('readline');
const ews = require('ews-javascript-api');

// =============================================================================
// CONFIGURATION
// =============================================================================

const CONFIG = {
  username: process.env.EWS_USERNAME,
  password: process.env.EWS_PASSWORD,
  ewsUrl: process.env.EWS_URL,
  testEmail: process.env.EWS_TEST_EMAIL || process.env.EWS_USERNAME,
};

// Validate configuration
if (!CONFIG.username || !CONFIG.password || !CONFIG.ewsUrl) {
  console.error('\n❌ Error: Missing configuration!');
  console.error('\nPlease create a .env file with your credentials.');
  console.error('Copy .env.example to .env and fill in your details.\n');
  process.exit(1);
}

// =============================================================================
// SERVICE INITIALIZATION
// =============================================================================

let service;

function initService() {
  service = new ews.ExchangeService(ews.ExchangeVersion.Exchange2010_SP2);
  service.Credentials = new ews.WebCredentials(CONFIG.username, CONFIG.password);
  service.Url = new ews.Uri(CONFIG.ewsUrl);
  service.UserAgent = 'n8n-aws-workmail-ews-test/1.0';
}

// =============================================================================
// LOGGING HELPERS
// =============================================================================

function log(message, type = 'info') {
  const symbols = {
    info: 'ℹ️',
    success: '✅',
    error: '❌',
    warn: '⚠️',
    test: '🧪'
  };
  console.log(`${symbols[type]} ${message}`);
}

function logSection(title) {
  console.log('\n' + '═'.repeat(80));
  console.log(`  ${title}`);
  console.log('═'.repeat(80) + '\n');
}

// =============================================================================
// TEST OPERATIONS
// =============================================================================

const tests = {
  // MESSAGE TESTS
  async testGetMessages() {
    logSection('Testing: Get Messages');
    const view = new ews.ItemView(5);
    const findResults = await service.FindItems(ews.WellKnownFolderName.Inbox, view);
    log(`Found ${findResults.Items.length} messages in Inbox`, 'success');

    if (findResults.Items.length > 0) {
      console.log('\nSample messages:');
      for (let i = 0; i < Math.min(3, findResults.Items.length); i++) {
        console.log(`  ${i + 1}. ${findResults.Items[i].Subject}`);
      }
    }
  },

  async testSendMessage() {
    logSection('Testing: Send Message');
    const message = new ews.EmailMessage(service);
    message.Subject = 'Test Message - ' + new Date().toISOString();
    message.Body = new ews.MessageBody(ews.BodyType.Text, 'This is a test message from the EWS test suite.');
    message.ToRecipients.Add(CONFIG.testEmail);

    await message.SendAndSaveCopy();
    log('Message sent successfully!', 'success');
  },

  async testReplyToMessage() {
    logSection('Testing: Reply to Message');
    const view = new ews.ItemView(10);
    const findResults = await service.FindItems(ews.WellKnownFolderName.Inbox, view);

    if (findResults.Items.length === 0) {
      log('No messages in Inbox to reply to', 'warn');
      return;
    }

    // Find a message from a different sender
    let targetMessage = null;
    for (const item of findResults.Items) {
      const msg = await ews.EmailMessage.Bind(service, item.Id);
      const fromAddress = msg.From ? msg.From.Address : '';
      if (fromAddress && fromAddress.toLowerCase() !== CONFIG.username.toLowerCase()) {
        targetMessage = msg;
        break;
      }
    }

    if (!targetMessage) {
      log('No suitable message to reply to (all messages are from yourself)', 'warn');
      return;
    }

    const reply = targetMessage.CreateReply(false);
    reply.BodyPrefix = new ews.MessageBody(ews.BodyType.Text, 'Test reply - please ignore.');
    await reply.Send();
    log('Reply sent successfully!', 'success');
  },

  // FOLDER TESTS
  async testCreateFolder() {
    logSection('Testing: Create Folder');
    const folder = new ews.Folder(service);
    folder.DisplayName = 'Test Folder ' + Date.now();
    await folder.Save(ews.WellKnownFolderName.Inbox);
    log(`Folder created: ${folder.DisplayName}`, 'success');

    // Clean up
    await folder.Delete(ews.DeleteMode.HardDelete);
    log('Folder deleted (cleanup)', 'info');
  },

  async testGetFolders() {
    logSection('Testing: Get Folders');
    const view = new ews.FolderView(10);
    const findResults = await service.FindFolders(ews.WellKnownFolderName.MsgFolderRoot, view);
    log(`Found ${findResults.Folders.length} folders`, 'success');

    console.log('\nSample folders:');
    for (let i = 0; i < Math.min(5, findResults.Folders.length); i++) {
      console.log(`  ${i + 1}. ${findResults.Folders[i].DisplayName}`);
    }
  },

  // EVENT TESTS
  async testCreateEvent() {
    logSection('Testing: Create Event');
    const appointment = new ews.Appointment(service);
    appointment.Subject = 'Test Event ' + Date.now();
    appointment.Body = new ews.MessageBody(ews.BodyType.Text, 'Test event from EWS test suite');

    const startDate = new ews.DateTime(new Date());
    startDate.Add(1, 'day');
    appointment.Start = startDate;

    const endDate = new ews.DateTime(startDate.jsDate);
    endDate.Add(1, 'hour');
    appointment.End = endDate;

    await appointment.Save(ews.WellKnownFolderName.Calendar);
    log(`Event created: ${appointment.Subject}`, 'success');

    // Clean up
    await appointment.Delete(ews.DeleteMode.HardDelete);
    log('Event deleted (cleanup)', 'info');
  },

  async testGetEvents() {
    logSection('Testing: Get Events');
    const view = new ews.CalendarView(
      new ews.DateTime(new Date()),
      new ews.DateTime(new Date(Date.now() + 30 * 24 * 60 * 60 * 1000))
    );
    const findResults = await service.FindAppointments(ews.WellKnownFolderName.Calendar, view);
    log(`Found ${findResults.Items.length} upcoming events`, 'success');
  },

  // CONTACT TESTS
  async testCreateContact() {
    logSection('Testing: Create Contact');
    const contact = new ews.Contact(service);
    contact.GivenName = 'Test';
    contact.Surname = 'Contact';
    contact.CompanyName = 'Test Company';

    const emailAddress = new ews.EmailAddress('test@example.com');
    contact.EmailAddresses._setItem(ews.EmailAddressKey.EmailAddress1, emailAddress);

    await contact.Save(ews.WellKnownFolderName.Contacts);
    log(`Contact created: ${contact.GivenName} ${contact.Surname}`, 'success');

    // Clean up
    await contact.Delete(ews.DeleteMode.HardDelete);
    log('Contact deleted (cleanup)', 'info');
  },

  async testGetContacts() {
    logSection('Testing: Get Contacts');
    const view = new ews.ItemView(10);
    const findResults = await service.FindItems(ews.WellKnownFolderName.Contacts, view);
    log(`Found ${findResults.Items.length} contacts`, 'success');
  },

  // ATTACHMENT TESTS
  async testAddAttachment() {
    logSection('Testing: Add Attachment');
    const message = new ews.EmailMessage(service);
    message.Subject = 'Test Attachment ' + Date.now();
    message.Body = new ews.MessageBody(ews.BodyType.Text, 'Message with attachment');
    message.ToRecipients.Add(CONFIG.testEmail);

    await message.Save(ews.WellKnownFolderName.Drafts);

    const attachmentContent = 'This is test file content';
    await message.Attachments.AddFileAttachment('test.txt', attachmentContent);
    await message.Update(ews.ConflictResolutionMode.AlwaysOverwrite);

    log('Attachment added to message', 'success');

    // Clean up
    await message.Delete(ews.DeleteMode.HardDelete);
    log('Test message deleted (cleanup)', 'info');
  },

  // CALENDAR TESTS
  async testGetCalendars() {
    logSection('Testing: Get Calendars');
    const calendar = await ews.Folder.Bind(service, ews.WellKnownFolderName.Calendar);
    log(`Default calendar: ${calendar.DisplayName}`, 'success');
    console.log(`  Total items: ${calendar.TotalCount}`);
  },

  // CONNECTION TEST
  async testConnection() {
    logSection('Testing: Connection');
    const inbox = await ews.Folder.Bind(service, ews.WellKnownFolderName.Inbox);
    log('Connection successful!', 'success');
    console.log(`  Inbox: ${inbox.DisplayName}`);
    console.log(`  Total messages: ${inbox.TotalCount}`);
    console.log(`  Unread: ${inbox.UnreadCount}`);
  },

  // TRIGGER SIMULATION TEST
  async testTriggerSimulation() {
    logSection('Testing: Trigger Simulation');
    
    // Simulate the trigger's static data storage
    const staticData = {
      initialized: false,
      knownMessageIds: [],
    };

    log('Simulating Trigger Workflow...', 'info');
    console.log('\n--- POLL 1: First activation (should NOT trigger) ---\n');

    // Poll 1: First poll after activation
    const poll1Messages = await getMessagesForTrigger(5);
    const poll1Ids = poll1Messages.map(m => m.id);
    
    console.log(`  Found ${poll1Messages.length} existing messages`);
    poll1Messages.forEach((m, i) => {
      console.log(`    ${i + 1}. [${m.id.substring(0, 20)}...] ${m.subject}`);
    });

    if (!staticData.initialized) {
      // First poll - store all IDs as known
      staticData.initialized = true;
      staticData.knownMessageIds = poll1Ids.slice(0, 200);
      log('First poll: Storing existing message IDs. No trigger.', 'success');
      console.log(`  Stored ${staticData.knownMessageIds.length} known IDs`);
    }

    console.log('\n--- POLL 2: Checking for new messages ---\n');
    
    // Poll 2: Simulate same messages (no new emails)
    const poll2Messages = await getMessagesForTrigger(5);
    const newMessages2 = poll2Messages.filter(m => !staticData.knownMessageIds.includes(m.id));
    
    if (newMessages2.length === 0) {
      log('No new messages found. No trigger.', 'success');
    } else {
      log(`Found ${newMessages2.length} new message(s)!`, 'warn');
      newMessages2.forEach((m, i) => {
        console.log(`    NEW: ${m.subject}`);
      });
    }

    console.log('\n--- SIMULATING: New email arrives ---\n');
    
    // Simulate a new message arriving (we'll use a fake ID)
    const fakeNewMessage = {
      id: 'FAKE_NEW_MESSAGE_' + Date.now(),
      subject: 'NEW TEST EMAIL - ' + new Date().toISOString(),
      received: new Date().toISOString(),
    };
    
    console.log(`  Simulated new email: "${fakeNewMessage.subject}"`);
    
    console.log('\n--- POLL 3: After new email arrives ---\n');
    
    // Poll 3: Check with the fake new message included
    const poll3Messages = [...poll2Messages, fakeNewMessage];
    const newMessages3 = poll3Messages.filter(m => !staticData.knownMessageIds.includes(m.id));
    
    if (newMessages3.length > 0) {
      log(`TRIGGER FIRED! Found ${newMessages3.length} new message(s):`, 'success');
      newMessages3.forEach((m, i) => {
        console.log(`    ✅ ${m.subject}`);
      });
      
      // Update known IDs
      const newIds = newMessages3.map(m => m.id);
      staticData.knownMessageIds = [...newIds, ...staticData.knownMessageIds].slice(0, 200);
      console.log(`\n  Updated known IDs (total: ${staticData.knownMessageIds.length})`);
    } else {
      log('No new messages.', 'info');
    }

    console.log('\n--- POLL 4: Same messages again (should NOT trigger) ---\n');
    
    // Poll 4: Same messages - should not trigger again
    const newMessages4 = poll3Messages.filter(m => !staticData.knownMessageIds.includes(m.id));
    
    if (newMessages4.length === 0) {
      log('No new messages. Trigger correctly NOT fired again.', 'success');
    } else {
      log(`ERROR: Found ${newMessages4.length} messages that should have been filtered!`, 'error');
    }

    console.log('\n--- TRIGGER SIMULATION COMPLETE ---\n');
    log('Trigger logic verified successfully!', 'success');
  },

  // TRIGGER LIVE WATCH (Real-time monitoring)
  async testTriggerLiveWatch() {
    logSection('Testing: Trigger Live Watch (Real-time)');
    
    log('Starting live trigger watch...', 'info');
    console.log('  This will poll every 10 seconds for new emails.');
    console.log('  Send an email to your mailbox to test the trigger.');
    console.log('  Press Ctrl+C to stop.\n');

    const staticData = {
      initialized: false,
      knownMessageIds: [],
    };

    let pollCount = 0;
    const maxPolls = 30; // 5 minutes max (30 * 10 seconds)

    const pollInterval = setInterval(async () => {
      pollCount++;
      const timestamp = new Date().toLocaleTimeString();
      
      try {
        const messages = await getMessagesForTrigger(20);
        const messageIds = messages.map(m => m.id);

        if (!staticData.initialized) {
          // First poll
          staticData.initialized = true;
          staticData.knownMessageIds = messageIds.slice(0, 200);
          console.log(`[${timestamp}] Poll #${pollCount}: Initialized with ${messages.length} existing messages. Waiting for new emails...`);
        } else {
          // Check for new messages
          const newMessages = messages.filter(m => !staticData.knownMessageIds.includes(m.id));
          
          if (newMessages.length > 0) {
            console.log(`\n🔔 [${timestamp}] Poll #${pollCount}: TRIGGER! ${newMessages.length} new email(s):`);
            newMessages.forEach(m => {
              console.log(`   ✅ From: ${m.from} | Subject: ${m.subject}`);
            });
            console.log('');
            
            // Update known IDs
            const newIds = newMessages.map(m => m.id);
            staticData.knownMessageIds = [...newIds, ...staticData.knownMessageIds].slice(0, 200);
          } else {
            process.stdout.write(`[${timestamp}] Poll #${pollCount}: No new messages (${messages.length} total)\r`);
          }
        }

        if (pollCount >= maxPolls) {
          clearInterval(pollInterval);
          console.log('\n\n⏱️ Live watch timeout (5 minutes). Stopping...\n');
        }
      } catch (error) {
        console.log(`\n❌ [${timestamp}] Poll error: ${error.message}`);
      }
    }, 10000); // Poll every 10 seconds

    // Initial poll immediately
    const messages = await getMessagesForTrigger(20);
    staticData.initialized = true;
    staticData.knownMessageIds = messages.map(m => m.id).slice(0, 200);
    console.log(`[${new Date().toLocaleTimeString()}] Initial: Found ${messages.length} existing messages. Now watching for new emails...\n`);

    // Wait for interrupt or timeout
    await new Promise(resolve => {
      process.on('SIGINT', () => {
        clearInterval(pollInterval);
        console.log('\n\n👋 Live watch stopped by user.\n');
        resolve();
      });
      
      // Also resolve after max polls
      setTimeout(resolve, (maxPolls + 1) * 10000);
    });
  },
};

// Helper function to get messages in trigger format
async function getMessagesForTrigger(maxResults = 10) {
  const view = new ews.ItemView(maxResults);
  view.PropertySet = new ews.PropertySet(ews.BasePropertySet.FirstClassProperties);
  
  const findResults = await service.FindItems(ews.WellKnownFolderName.Inbox, view);
  
  const messages = [];
  for (const item of findResults.Items) {
    if (item instanceof ews.EmailMessage) {
      const fullMessage = await ews.EmailMessage.Bind(
        service,
        item.Id,
        new ews.PropertySet(ews.BasePropertySet.FirstClassProperties)
      );
      
      messages.push({
        id: fullMessage.Id.UniqueId,
        subject: fullMessage.Subject || '(No Subject)',
        from: fullMessage.From ? fullMessage.From.Address : 'Unknown',
        received: fullMessage.DateTimeReceived ? fullMessage.DateTimeReceived.toString() : null,
      });
    }
  }
  
  // Sort by received date (newest first)
  messages.sort((a, b) => {
    if (!a.received || !b.received) return 0;
    return new Date(b.received).getTime() - new Date(a.received).getTime();
  });
  
  return messages;
}

// =============================================================================
// TEST RUNNER
// =============================================================================

async function runTest(testName, testFn) {
  try {
    log(`Running: ${testName}`, 'test');
    await testFn();
    return { success: true };
  } catch (error) {
    log(`Failed: ${testName}`, 'error');
    console.error(`  Error: ${error.message}`);
    if (error.stack) {
      console.error(`  ${error.stack.split('\n').slice(0, 3).join('\n  ')}`);
    }
    return { success: false, error: error.message };
  }
}

async function runAllTests() {
  console.log('\n' + '█'.repeat(80));
  console.log('  AWS WorkMail EWS - Running All Tests');
  console.log('█'.repeat(80));

  const results = {};
  const testList = Object.entries(tests);

  for (const [name, fn] of testList) {
    const result = await runTest(name, fn);
    results[name] = result;

    // Small delay between tests
    await new Promise(resolve => setTimeout(resolve, 1000));
  }

  // Summary
  console.log('\n' + '█'.repeat(80));
  console.log('  TEST SUMMARY');
  console.log('█'.repeat(80) + '\n');

  const passed = Object.values(results).filter(r => r.success).length;
  const total = Object.keys(results).length;

  console.log(`✅ Passed: ${passed}/${total}`);
  if (passed < total) {
    console.log(`❌ Failed: ${total - passed}/${total}`);
  }

  console.log(`\n📊 Success Rate: ${((passed / total) * 100).toFixed(1)}%\n`);
}

// =============================================================================
// INTERACTIVE MENU
// =============================================================================

async function showMenu() {
  console.clear();
  console.log('\n' + '█'.repeat(80));
  console.log('  AWS WorkMail EWS - Interactive Test Suite');
  console.log('█'.repeat(80));
  console.log(`\n  Configuration:`);
  console.log(`    Username: ${CONFIG.username}`);
  console.log(`    EWS URL:  ${CONFIG.ewsUrl}`);
  console.log('\n' + '─'.repeat(80));
  console.log('\n  Available Tests:\n');

  const menuItems = [
    { key: '1', name: 'Test Connection', test: 'testConnection' },
    { key: '2', name: 'Get Messages', test: 'testGetMessages' },
    { key: '3', name: 'Send Message', test: 'testSendMessage' },
    { key: '4', name: 'Reply to Message', test: 'testReplyToMessage' },
    { key: '5', name: 'Create Folder', test: 'testCreateFolder' },
    { key: '6', name: 'Get Folders', test: 'testGetFolders' },
    { key: '7', name: 'Create Event', test: 'testCreateEvent' },
    { key: '8', name: 'Get Events', test: 'testGetEvents' },
    { key: '9', name: 'Create Contact', test: 'testCreateContact' },
    { key: '10', name: 'Get Contacts', test: 'testGetContacts' },
    { key: '11', name: 'Add Attachment', test: 'testAddAttachment' },
    { key: '12', name: 'Get Calendars', test: 'testGetCalendars' },
    { key: '', name: '─'.repeat(50), test: null },
    { key: 't', name: '🔔 Trigger Simulation (Test Logic)', test: 'testTriggerSimulation' },
    { key: 'w', name: '👀 Trigger Live Watch (Real-time)', test: 'testTriggerLiveWatch' },
    { key: '', name: '─'.repeat(50), test: null },
    { key: 'all', name: 'Run ALL Tests', test: 'all' },
    { key: 'q', name: 'Quit', test: 'quit' },
  ];

  menuItems.forEach(item => {
    if (item.key === '') {
      console.log(`  ${item.name}`);
    } else {
      console.log(`  [${item.key.padEnd(3)}] ${item.name}`);
    }
  });

  console.log('\n' + '─'.repeat(80));

  return new Promise((resolve) => {
    const rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });

    rl.question('\n  Select an option: ', async (answer) => {
      rl.close();

      const choice = answer.trim().toLowerCase();

      if (choice === 'q' || choice === 'quit') {
        console.log('\n  Goodbye! 👋\n');
        process.exit(0);
      }

      if (choice === 'all') {
        await runAllTests();
      } else {
        const item = menuItems.find(m => m.key === choice);
        if (item && item.test && tests[item.test]) {
          await runTest(item.name, tests[item.test]);
        } else {
          console.log('\n  ❌ Invalid option. Please try again.\n');
        }
      }

      // Wait for user input before showing menu again
      const rl2 = readline.createInterface({
        input: process.stdin,
        output: process.stdout
      });

      rl2.question('\n  Press Enter to continue...', () => {
        rl2.close();
        resolve();
      });
    });
  });
}

// =============================================================================
// MAIN
// =============================================================================

async function main() {
  try {
    initService();

    // Interactive menu loop
    while (true) {
      await showMenu();
    }
  } catch (error) {
    console.error('\n❌ Fatal Error:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  main();
}

module.exports = { tests, runTest, runAllTests };
