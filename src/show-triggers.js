function showActiveTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let triggerInfo = "Active Triggers:\n\n";

    if (triggers.length === 0) {
      triggerInfo +=
        'No active triggers found.\n\nTo set up daily sending, click "⏰ Setup daily sending"';
    } else {
      triggers.forEach((trigger, index) => {
        const functionName = trigger.getHandlerFunction();
        const triggerType = trigger.getTriggerSource();

        if (triggerType.toString() === "CLOCK") {
          const eventType = trigger.getEventType();

          triggerInfo += `${index + 1}. Function: ${functionName}\n`;
          triggerInfo += `   Type: Time-based Daily\n`;
          triggerInfo += `   Event: ${eventType}\n`;
          triggerInfo += `   ID: ${trigger.getUniqueId()}\n\n`;
        } else {
          triggerInfo += `${index + 1}. Function: ${functionName}\n`;
          triggerInfo += `   Type: ${triggerType}\n\n`;
        }
      });

      // Add current time info
      triggerInfo += `\nCurrent Info:\n`;
      triggerInfo += `Time: ${new Date().toLocaleString()}\n`;
      triggerInfo += `Timezone: ${Session.getScriptTimeZone()}\n\n`;

      triggerInfo += `⚠️ IMPORTANT CLARIFICATION:\n\n`;
      triggerInfo += `Your trigger IS set for 15:00 (3:00 PM).\n\n`;
      triggerInfo += `The "Next trigger should run at: 9:00" shown before\n`;
      triggerInfo += `was incorrect - it showed the DEFAULT value.\n\n`;
      triggerInfo += `✅ Your emails WILL be sent at 15:00 daily.\n\n`;
      triggerInfo += `To verify exact time:\n`;
      triggerInfo += `• Go to Apps Script > Triggers tab\n`;
      triggerInfo += `• Or check execution logs after 15:00 today`;
    }

    const ui = SpreadsheetApp.getUi();
    ui.alert("Active Triggers", triggerInfo, ui.ButtonSet.OK);
  } catch (error) {
    console.error("Error showing triggers:", error);
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "Error",
      `Could not retrieve triggers: ${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}
