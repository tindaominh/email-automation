// Constants are now in constant.js file
// Make sure to include constant.js in your Google Apps Script project

// Check current timezone and trigger settings
function checkTimezone() {
  const now = new Date();
  const vietnamTime = new Date().toLocaleString("en-US", {
    timeZone: "Asia/Ho_Chi_Minh",
  });
  const scriptTime = new Date().toLocaleString();

  console.log("Current time in script:", scriptTime);
  console.log("Current time in Vietnam:", vietnamTime);
  console.log("Your Google Account timezone:", Session.getScriptTimeZone());

  const ui = SpreadsheetApp.getUi();
  ui.alert(
    "Timezone Check",
    `Script timezone: ${Session.getScriptTimeZone()}\n` +
      `Script time: ${scriptTime}\n` +
      `Vietnam time: ${vietnamTime}\n\n` +
      `If script timezone shows "Asia/Ho_Chi_Minh" or "Asia/Bangkok", then TRIGGER_HOUR uses Vietnam time.`,
    ui.ButtonSet.OK
  );
}

// Helper function to find column by name
function findColumnIndex(headers, columnName) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase().trim() === columnName.toLowerCase().trim()) {
      return i;
    }
  }
  return -1;
}

// Simple email validation
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

// Get value from row by column name
function getColumnValue(row, columnMap, columnName, defaultValue = "") {
  const colIndex = columnMap[columnName];
  return colIndex !== undefined ? row[colIndex] || defaultValue : defaultValue;
}

// Validate a single row of data
function validateRowData(row, columnMap, rowIndex) {
  const email = String(
    getColumnValue(row, columnMap, COLUMN_HEADERS.EMAIL)
  ).trim();
  const subject = String(
    getColumnValue(row, columnMap, COLUMN_HEADERS.SUBJECT)
  ).trim();
  const content = String(
    getColumnValue(row, columnMap, COLUMN_HEADERS.CONTENT)
  ).trim();
  const attachment = String(
    getColumnValue(row, columnMap, COLUMN_HEADERS.ATTACHMENT)
  ).trim();
  const startDateValue = getColumnValue(
    row,
    columnMap,
    COLUMN_HEADERS.START_DATE
  );
  const deadlineValue = getColumnValue(row, columnMap, COLUMN_HEADERS.DEADLINE);
  const status = String(
    getColumnValue(row, columnMap, COLUMN_HEADERS.STATUS)
  ).trim();

  // Check required fields
  if (!email) {
    console.log(`Row ${rowIndex + 2}: Email is required`);
    return null;
  }

  if (!isValidEmail(email)) {
    console.log(`Row ${rowIndex + 2}: Invalid email format: ${email}`);
    return null;
  }

  if (!subject) {
    console.log(`Row ${rowIndex + 2}: Subject is required`);
    return null;
  }

  if (!content) {
    console.log(`Row ${rowIndex + 2}: Content is required`);
    return null;
  }

  // Validate start date
  const startDate = new Date(startDateValue);
  if (isNaN(startDate.getTime())) {
    console.log(`Row ${rowIndex + 2}: Invalid start date: ${startDateValue}`);
    return null;
  }

  // Validate deadline
  const deadline = new Date(deadlineValue);
  if (isNaN(deadline.getTime())) {
    console.log(`Row ${rowIndex + 2}: Invalid deadline date: ${deadlineValue}`);
    return null;
  }

  return {
    email: email,
    subject: subject,
    content: content,
    attachment: attachment,
    startDate: startDate,
    deadline: deadline,
    status: status,
    rowIndex: rowIndex + 2, // +2 for header row and 1-based indexing
  };
}

// Check if we should send email today
function shouldSendEmail(deadline, today, status) {
  // Don't send if already sent
  if (status && status.toLowerCase().includes("sent")) {
    return false;
  }

  // Calculate days until deadline
  const deadlineDate = new Date(deadline);
  deadlineDate.setHours(0, 0, 0, 0);
  const daysUntilDeadline = Math.floor(
    (deadlineDate - today) / (1000 * 60 * 60 * 24)
  );

  // Send if within configured days or overdue
  return daysUntilDeadline <= CONFIG.DAYS_BEFORE_DEADLINE;
}

// Create email content
function prepareEmailBody(content, startDate, deadline) {
  return content;
}

// Create HTML email using content from Sheet data
function formatHtmlBody(content, startDate, deadline) {
  // Convert content to HTML format, preserving line breaks
  const htmlContent = content.replace(/\n/g, "<br>");

  // Calculate days overdue
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const deadlineDate = new Date(deadline);
  deadlineDate.setHours(0, 0, 0, 0);
  const daysOverdue = Math.floor(
    (today - deadlineDate) / (1000 * 60 * 60 * 24)
  );

  // Format dates for display
  const formattedStartDate = formatVietnameseDate(startDate);
  const formattedDeadline = formatVietnameseDate(deadline);

  return `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 650px; margin: 0 auto; background-color: #f8f9fa;">
      <!-- Header -->
      <div style="background-color: #dc3545; color: white; padding: 15px; text-align: center; font-weight: bold; border-radius: 8px 8px 0 0;">
        ⚠️ THÔNG BÁO CÔNG VIỆC QUÁ DEADLINE
      </div>
      
      <!-- Main Content -->
      <div style="background-color: white; padding: 30px; border-radius: 0 0 8px 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
        <!-- Column C Content -->
        <div style="font-size: 15px; line-height: 1.6; color: #333; margin-bottom: 20px;">
          ${htmlContent}
        </div>
        
        <!-- Task Status Message -->
        <div style="color: #dc3545; font-weight: bold; margin: 15px 0;">
          Công việc dưới đây đã quá deadline ${daysOverdue} ngày và vẫn chưa được hoàn thành:
        </div>
        
        <!-- Task Details Box -->
        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 4px solid #dc3545; margin: 15px 0;">
          <div style="margin-bottom: 10px;">
            <span style="display: inline-block; width: 15px; height: 15px; background-color: #6c757d; margin-right: 8px;"></span>
            <strong>Tóm tắc sách</strong>
          </div>
          
          <div style="margin: 10px 0; color: #666;">
            <strong>Ngày bắt đầu:</strong> ${formattedStartDate}
          </div>
          
          <div style="margin: 10px 0; color: #dc3545;">
            <strong>Deadline:</strong> ${formattedDeadline}
          </div>
          
          <div style="margin: 10px 0; color: #dc3545;">
            <strong>Trạng thái:</strong> Chưa
          </div>
        </div>
        
        <!-- Footer Message -->
        <div style="margin-top: 20px; color: #666; font-size: 14px;">
          Vui lòng hoàn thành công việc này trong thời gian sớm nhất có thể và cập nhật trạng thái trong bảng kế hoạch.
        </div>
        
        <div style="margin-top: 15px; color: #666; font-size: 14px;">
          <strong>Lưu ý:</strong> Nếu không hoàn thành trong 24h, hệ thống sẽ tự động gọi điện để nhắc nhở.
        </div>
        
        <div style="margin-top: 20px; color: #666; font-size: 14px;">
          Trân trọng,<br>
          Hệ thống Quản lý Dự án Tự động
        </div>
      </div>
      
      <!-- System Footer -->
      <div style="text-align: center; padding: 15px; font-size: 11px; color: #7f8c8d; background-color: #ecf0f1; margin-top: 10px; border-radius: 6px;">
        Email này được gửi tự động từ hệ thống quản lý dự án. Vui lòng không trả lời email này.
      </div>
    </div>
  `;
}

// Format date as DD/MM/YYYY HH:MM
function formatDate(date) {
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();
  const hours = String(d.getHours()).padStart(2, "0");
  const minutes = String(d.getMinutes()).padStart(2, "0");

  return `${day}/${month}/${year} ${hours}:${minutes}`;
}

// Format Vietnamese date as YYYY-MM-DD
function formatVietnameseDate(date) {
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, "0");
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const year = d.getFullYear();

  return `${year}-${month}-${day}`;
}

// Update status in spreadsheet
function updateRowStatus(rowIndex, status, columnMap) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEET_NAME
    );
    const statusColumn = columnMap[COLUMN_HEADERS.STATUS];

    if (statusColumn !== undefined) {
      sheet.getRange(rowIndex, statusColumn + 1).setValue(status);
    } else {
      console.log("Status column not found, cannot update status");
    }
  } catch (error) {
    console.error("Error updating status:", error);
  }
}

// Send error notification to admin
function notifyAdmin(error) {
  const subject = "Error in automatic email system";
  const body = `An error occurred in the automatic email process:\n\n${error.toString()}`;

  try {
    MailApp.sendEmail(CONFIG.ADMIN_EMAIL, subject, body);
  } catch (e) {
    console.error("Unable to send error notification email:", e);
  }
}

function sendDailyEmails() {
  try {
    console.log("Starting email automation...");

    const sheetData = getSheetData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    let processedCount = 0;
    let sentCount = 0;
    let errorCount = 0;

    console.log(`Found ${sheetData.data.length} rows of data`);
    console.log(`Current time: ${new Date().toLocaleString()}`);
    console.log(
      `Checking emails due within ${CONFIG.DAYS_BEFORE_DEADLINE} days`
    );

    // Process each row
    for (let i = 0; i < sheetData.data.length; i++) {
      const row = sheetData.data[i];
      const emailData = validateRowData(row, sheetData.columnMap, i);

      if (!emailData) {
        continue; // Skip invalid rows
      }

      processedCount++;

      // Check if we should send email
      if (shouldSendEmail(emailData.deadline, today, emailData.status)) {
        try {
          // Prepare and send email
          const emailBody = prepareEmailBody(
            emailData.content,
            emailData.startDate,
            emailData.deadline
          );

          MailApp.sendEmail(emailData.email, emailData.subject, emailBody, {
            htmlBody: formatHtmlBody(
              emailData.content,
              emailData.startDate,
              emailData.deadline
            ),
          });

          // Update status
          updateRowStatus(
            emailData.rowIndex,
            `Sent - ${formatDate(today)}`,
            sheetData.columnMap
          );

          console.log(`✓ Email sent to: ${emailData.email}`);
          sentCount++;
        } catch (emailError) {
          console.error(`✗ Error sending to ${emailData.email}:`, emailError);
          updateRowStatus(
            emailData.rowIndex,
            `Send Error - ${formatDate(today)}`,
            sheetData.columnMap
          );
          errorCount++;
        }
      } else {
        console.log(`- Skipped ${emailData.email} (already sent or not due)`);
      }
    }

    console.log(`\nEmail automation complete!`);
    console.log(
      `Processed: ${processedCount}, Sent: ${sentCount}, Errors: ${errorCount}`
    );
  } catch (error) {
    console.error("Error in email automation:", error);
    notifyAdmin(error);
  }
}

// Setup daily trigger with custom time input
function setupDailyTrigger() {
  try {
    console.log("Setting up daily trigger...");

    const ui = SpreadsheetApp.getUi();

    // Ask user for time input
    const defaultHour = CONFIG.TRIGGER_HOUR;
    // Calculate minimum recommended hour
    const now = new Date();
    const currentHour = now.getHours();
    const currentMinutes = now.getMinutes();
    let minRecommendedHour = currentHour + 1;
    if (currentMinutes > 0) {
      minRecommendedHour = currentHour + 2;
    }
    if (minRecommendedHour > 23) {
      minRecommendedHour = minRecommendedHour - 24;
    }

    let triggerHour = defaultHour;
    let userInput = "";
    let validInput = false;
    let errorMessage = "";

    // Loop until valid input or user cancels
    while (!validInput) {
      // Update current time for each iteration
      const currentTime = new Date();
      const currentHour = currentTime.getHours();
      const currentMinutes = currentTime.getMinutes();
      let minRecommendedHour = currentHour + 1;
      if (currentMinutes > 0) {
        minRecommendedHour = currentHour + 2;
      }
      if (minRecommendedHour > 23) {
        minRecommendedHour = minRecommendedHour - 24;
      }

      const response = ui.prompt(
        "Setup Daily Email Sending",
        `${errorMessage}` +
          `Enter the time when you want emails to be sent daily.\n\n` +
          `⚠️ IMPORTANT LIMITATIONS:\n` +
          `• Google Apps Script can only trigger at whole hours (9:15 → 9:00)\n` +
          `• Time must be at least 1 hour from now for reliable execution\n\n` +
          `Current time: ${currentHour}:${String(currentMinutes).padStart(
            2,
            "0"
          )}\n` +
          `Minimum recommended: ${minRecommendedHour}:00 (${formatHourDisplay(
            minRecommendedHour
          )})\n` +
          `Script timezone: ${Session.getScriptTimeZone()}\n\n` +
          `Accepted formats:\n` +
          `• 9 or 9:00 = 9:00 AM\n` +
          `• 14 or 14:30 = 2:00 PM (ignores :30)\n` +
          `• 20 or 20:45 = 8:00 PM (ignores :45)\n\n` +
          `Leave empty to use default (${defaultHour}):`,
        ui.ButtonSet.OK_CANCEL
      );

      // Handle user response
      if (response.getSelectedButton() === ui.Button.CANCEL) {
        console.log("User cancelled trigger setup");
        return;
      }

      userInput = response.getResponseText().trim();

      // Validate user input
      if (userInput !== "") {
        let inputHour;
        let hasMinutes = false;
        let originalInput = userInput;

        // Handle HH:MM format
        if (userInput.includes(":")) {
          const parts = userInput.split(":");
          if (parts.length === 2) {
            inputHour = parseInt(parts[0]);
            const minutes = parseInt(parts[1]);

            if (!isNaN(minutes) && minutes > 0) {
              hasMinutes = true;
            }
          } else {
            inputHour = NaN;
          }
        } else {
          // Handle just hour number
          inputHour = parseInt(userInput);
        }

        if (isNaN(inputHour) || inputHour < 0 || inputHour > 23) {
          errorMessage =
            `❌ ERROR: "${userInput}" is not a valid time.\n` +
            `Please enter a number between 0-23 (e.g., 9, 14, 20) or time format HH:MM.\n\n`;
          continue; // Go back to prompt
        }

        // Validate time is at least 1 hour from now
        const now = new Date();
        const nowHour = now.getHours();
        const nowMinutes = now.getMinutes();

        // Calculate minimum required hour (at least 1 hour from now)
        let minRequiredHour = nowHour + 1;
        if (nowMinutes > 0) {
          minRequiredHour = nowHour + 2; // If past the hour mark, need +2 hours
        }

        // Handle day rollover (if min required hour > 23, it's tomorrow)
        const isNextDay = minRequiredHour > 23;
        if (isNextDay) {
          minRequiredHour = minRequiredHour - 24;
        }

        // Check if input hour meets minimum requirement
        const needsNextDay =
          inputHour < nowHour || (inputHour === nowHour && nowMinutes > 0);
        const isTooSoon = !needsNextDay && inputHour < minRequiredHour;

        if (isTooSoon) {
          const recommendedHour = minRequiredHour;
          errorMessage =
            `❌ TIME TOO SOON ERROR:\n` +
            `Current time: ${nowHour}:${String(nowMinutes).padStart(
              2,
              "0"
            )}\n` +
            `You entered: ${inputHour}:00\n` +
            `Minimum required: ${recommendedHour}:00 (${formatHourDisplay(
              recommendedHour
            )})\n\n` +
            `🚫 Please enter a time at least 1 hour from now.\n\n`;
          continue; // Go back to prompt
        }

        triggerHour = inputHour;

        if (hasMinutes) {
          console.log(
            `User entered: ${originalInput}, using hour: ${triggerHour} (minutes ignored)`
          );

          // Show warning about minutes being ignored
          const continueResponse = ui.alert(
            "Minutes Will Be Ignored",
            `You entered: ${originalInput}\n\n` +
              `⚠️ Google Apps Script limitation:\n` +
              `Triggers can only run at whole hours.\n\n` +
              `Your emails will be sent at: ${triggerHour}:00 (${formatHourDisplay(
                triggerHour
              )})\n\n` +
              `Continue with ${triggerHour}:00?`,
            ui.ButtonSet.YES_NO
          );

          if (continueResponse === ui.Button.NO) {
            console.log("User cancelled after minutes warning");
            errorMessage = `⚠️ Please enter a different time (without minutes or accept the whole hour).\n\n`;
            continue; // Go back to prompt
          }
        } else {
          console.log(`User specified custom hour: ${triggerHour}`);
        }

        validInput = true; // Valid input received
      } else {
        console.log(`Using default hour: ${triggerHour}`);
        triggerHour = defaultHour;
        validInput = true; // Using default is valid
      }
    } // End of while loop

    // Delete old triggers
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "sendDailyEmails") {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }
    console.log(`Deleted ${deletedCount} old triggers`);

    // Create new trigger with specified hour
    const newTrigger = ScriptApp.newTrigger("sendDailyEmails")
      .timeBased()
      .atHour(triggerHour)
      .everyDays(1)
      .create();

    console.log(`✓ Daily trigger created successfully!`);
    console.log(
      `Will run at ${triggerHour}:00 (${formatHourDisplay(triggerHour)})`
    );
    console.log(`Trigger ID: ${newTrigger.getUniqueId()}`);
    console.log(`Script timezone: ${Session.getScriptTimeZone()}`);

    // Show confirmation dialog
    ui.alert(
      "Daily Trigger Setup Complete!",
      `✅ SUCCESS!\n\n` +
        `Emails will be sent daily at:\n` +
        `${triggerHour}:00 (${formatHourDisplay(triggerHour)})\n\n` +
        `Timezone: ${Session.getScriptTimeZone()}\n` +
        `Trigger ID: ${newTrigger.getUniqueId().substring(0, 8)}...\n\n` +
        `Next steps:\n` +
        `1. Use "🧪 Test with first row" to test emails\n` +
        `2. Use "📋 View active triggers" to verify\n` +
        `3. Check your email tomorrow at ${formatHourDisplay(triggerHour)}!`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    console.error("Error setting up trigger:", error);
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Error!",
      `❌ Failed to set up daily trigger:\n\n${error.toString()}\n\nCheck the Apps Script logs for more details.`,
      ui.ButtonSet.OK
    );
    notifyAdmin(error);
  }
}

// Helper function to format hour display (24h to 12h format)
function formatHourDisplay(hour) {
  if (hour === 0) return "12:00 AM (Midnight)";
  if (hour === 12) return "12:00 PM (Noon)";
  if (hour < 12) return `${hour}:00 AM`;
  return `${hour - 12}:00 PM`;
}

// Test with first row
function testSendEmail() {
  try {
    console.log("Testing email with first row...");

    const sheetData = getSheetData();

    if (sheetData.data.length === 0) {
      console.error("No data rows found for testing");
      return;
    }

    const firstRow = sheetData.data[0];
    const emailData = validateRowData(firstRow, sheetData.columnMap, 0);

    if (!emailData) {
      console.error("First row data is invalid");
      return;
    }

    const emailBody = prepareEmailBody(
      emailData.content,
      emailData.startDate,
      emailData.deadline
    );

    MailApp.sendEmail(emailData.email, emailData.subject, emailBody, {
      htmlBody: formatHtmlBody(
        emailData.content,
        emailData.startDate,
        emailData.deadline
      ),
    });

    updateRowStatus(emailData.rowIndex, "Test - Sent", sheetData.columnMap);
    console.log("✓ Test email sent successfully!");
  } catch (error) {
    console.error("Error sending test email:", error);
    notifyAdmin(error);
  }
}

// Clear all status values
function resetStatus() {
  try {
    const sheetData = getSheetData();
    const statusColumn = sheetData.columnMap[COLUMN_HEADERS.STATUS];

    if (statusColumn === undefined) {
      console.log("Status column not found, cannot reset");
      return;
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
      CONFIG.SHEET_NAME
    );
    const lastRow = sheet.getLastRow();

    if (lastRow > 1) {
      sheet.getRange(2, statusColumn + 1, lastRow - 1, 1).clear();
      console.log("Status cleared for all rows");
    }
  } catch (error) {
    console.error("Error resetting status:", error);
    notifyAdmin(error);
  }
}

// Debug function to check why emails aren't being sent
function debugEmailSending() {
  try {
    console.log("=== EMAIL DEBUG ANALYSIS ===");

    const sheetData = getSheetData();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    console.log(`Found ${sheetData.data.length} rows of data`);
    console.log(`Today: ${today.toDateString()}`);
    console.log(`DAYS_BEFORE_DEADLINE: ${CONFIG.DAYS_BEFORE_DEADLINE}`);
    console.log("Column mapping:", sheetData.columnMap);

    // Check each row
    for (let i = 0; i < sheetData.data.length; i++) {
      console.log(`\n--- ROW ${i + 2} ANALYSIS ---`);
      const row = sheetData.data[i];

      // Show raw row data
      console.log("Raw row data:", row);

      // Try to validate
      const emailData = validateRowData(row, sheetData.columnMap, i);

      if (!emailData) {
        console.log("❌ Row validation FAILED");
        continue;
      }

      console.log("✅ Row validation PASSED");
      console.log("Email:", emailData.email);
      console.log("Subject:", emailData.subject);
      console.log("Start Date:", emailData.startDate);
      console.log("Deadline:", emailData.deadline);
      console.log("Status:", emailData.status);

      // Check if email should be sent
      const shouldSend = shouldSendEmail(
        emailData.deadline,
        today,
        emailData.status
      );
      console.log(`Should send email: ${shouldSend}`);

      if (!shouldSend) {
        const deadlineDate = new Date(emailData.deadline);
        deadlineDate.setHours(0, 0, 0, 0);
        const daysUntilDeadline = Math.floor(
          (deadlineDate - today) / (1000 * 60 * 60 * 24)
        );

        console.log(`Days until deadline: ${daysUntilDeadline}`);
        console.log(
          `Status contains 'sent': ${
            emailData.status && emailData.status.toLowerCase().includes("sent")
          }`
        );

        if (daysUntilDeadline > CONFIG.DAYS_BEFORE_DEADLINE) {
          console.log(
            `❌ Not sending: Deadline is ${daysUntilDeadline} days away (> ${CONFIG.DAYS_BEFORE_DEADLINE})`
          );
        }
        if (
          emailData.status &&
          emailData.status.toLowerCase().includes("sent")
        ) {
          console.log(`❌ Not sending: Status indicates already sent`);
        }
      } else {
        console.log("🚀 THIS EMAIL WOULD BE SENT!");
      }
    }

    console.log("\n=== DEBUG COMPLETE ===");

    // Show results to user
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Debug Complete",
      "Check the Apps Script logs (View > Logs) for detailed analysis of why emails are not being sent.",
      ui.ButtonSet.OK
    );
  } catch (error) {
    console.error("Debug error:", error);
    const ui = SpreadsheetApp.getUi();
    ui.alert("Debug Error", error.toString(), ui.ButtonSet.OK);
  }
}

// Show usage instructions
function showInstructions() {
  const html = `
    <div style="padding: 20px; font-family: Arial;">
      <h2>📋 How to Use Email Automation</h2>
      <h3>📊 Your Spreadsheet Columns:</h3>
      <ul>
        <li><strong>"${COLUMN_HEADERS.EMAIL}"</strong> - Email addresses to send to</li>
        <li><strong>"${COLUMN_HEADERS.SUBJECT}"</strong> - Email subject line</li>
        <li><strong>"${COLUMN_HEADERS.CONTENT}"</strong> - Email message content</li>
        <li><strong>"${COLUMN_HEADERS.ATTACHMENT}"</strong> - File attachments (optional)</li>
        <li><strong>"${COLUMN_HEADERS.DEADLINE}"</strong> - When task is due</li>
        <li><strong>"${COLUMN_HEADERS.STATUS}"</strong> - Automatically updated</li>
      </ul>
      
      <h3>⚙️ Setup:</h3>
      <ol>
        <li><strong>Setup Daily Sending:</strong> Emails sent automatically at ${CONFIG.TRIGGER_HOUR}:00 AM</li>
        <li><strong>Send Now:</strong> Send all due emails immediately</li>
        <li><strong>Test:</strong> Send test email using first row</li>
        <li><strong>Clear Status:</strong> Reset all status values to test again</li>
      </ol>
      
      <h3>📅 When Emails Are Sent:</h3>
      <p>Emails are sent when deadline is within <strong>${CONFIG.DAYS_BEFORE_DEADLINE} days</strong> or overdue.</p>
      
      <h3>🎨 Email Colors:</h3>
      <ul>
        <li>🔴 <strong>Red:</strong> Due today or overdue</li>
        <li>🟡 <strong>Yellow:</strong> Due within 3 days</li>
        <li>🟢 <strong>Green:</strong> Due within ${CONFIG.DAYS_BEFORE_DEADLINE} days</li>
      </ul>
    </div>
  `;

  const ui = SpreadsheetApp.getUi();
  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(600).setHeight(500),
    "Email Automation Instructions"
  );
}

// Create menu when sheet opens
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("📧 Email Automation")
      .addItem("▶️ Send emails now", "sendDailyEmails")
      .addItem("⏰ Setup daily sending", "setupDailyTrigger")
      .addItem("❌ Cancel daily sending", "cancelDailyTriggers")
      .addSeparator()
      .addItem("🧪 Test with first row", "testSendEmail")
      .addItem("🔄 Clear all status", "resetStatus")
      .addSeparator()
      .addItem("🔍 Debug email sending", "debugEmailSending")
      .addItem("🕐 Check timezone", "checkTimezone")
      .addItem("📋 View active triggers", "showActiveTriggers")
      .addItem("ℹ️ How to use", "showInstructions")
      .addToUi();
  } catch (error) {
    console.error("Error setting up menu:", error);
  }
}
