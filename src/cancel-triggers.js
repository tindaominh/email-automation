// Cancel/Delete all email automation triggers
function cancelDailyTriggers() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get all triggers
    const triggers = ScriptApp.getProjectTriggers();
    const emailTriggers = triggers.filter(trigger => 
      trigger.getHandlerFunction() === 'sendDailyEmails'
    );
    
    if (emailTriggers.length === 0) {
      ui.alert('No Triggers Found', 
        'No daily email triggers are currently active.\n\n' +
        'Nothing to cancel.',
        ui.ButtonSet.OK);
      return;
    }
    
    // Show confirmation dialog
    const response = ui.alert(
      'Cancel Daily Email Triggers?',
      `Found ${emailTriggers.length} daily email trigger(s).\n\n` +
      '⚠️ This will STOP automatic email sending.\n\n' +
      'Are you sure you want to cancel all daily email triggers?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.NO) {
      console.log('User cancelled trigger deletion');
      return;
    }
    
    // Delete the triggers
    let deletedCount = 0;
    let errorCount = 0;
    
    emailTriggers.forEach(trigger => {
      try {
        ScriptApp.deleteTrigger(trigger);
        console.log(`Deleted trigger: ${trigger.getUniqueId()}`);
        deletedCount++;
      } catch (error) {
        console.error(`Error deleting trigger ${trigger.getUniqueId()}:`, error);
        errorCount++;
      }
    });
    
    // Show result
    if (errorCount === 0) {
      ui.alert('Triggers Cancelled Successfully', 
        `✅ Successfully cancelled ${deletedCount} daily email trigger(s).\n\n` +
        '📧 Automatic email sending has been stopped.\n\n' +
        'To set up automatic sending again:\n' +
        'Use "⏰ Setup daily sending" from the menu.',
        ui.ButtonSet.OK);
      
      console.log(`✓ Successfully deleted ${deletedCount} email triggers`);
    } else {
      ui.alert('Partial Success', 
        `⚠️ Results:\n` +
        `• Successfully cancelled: ${deletedCount} triggers\n` +
        `• Failed to cancel: ${errorCount} triggers\n\n` +
        'Check the Apps Script logs for error details.',
        ui.ButtonSet.OK);
      
      console.log(`Deleted ${deletedCount} triggers, ${errorCount} errors`);
    }
    
  } catch (error) {
    console.error('Error in cancelDailyTriggers:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 
      `❌ Failed to cancel triggers:\n\n${error.toString()}\n\n` +
      'Check the Apps Script logs for more details.',
      ui.ButtonSet.OK);
  }
}

// Cancel ALL triggers (not just email ones) - USE WITH CAUTION
function cancelAllTriggers() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const triggers = ScriptApp.getProjectTriggers();
    
    if (triggers.length === 0) {
      ui.alert('No Triggers Found', 
        'No triggers are currently active.\n\nNothing to cancel.',
        ui.ButtonSet.OK);
      return;
    }
    
    // Show warning
    const response = ui.alert(
      '⚠️ DANGER: Cancel ALL Triggers?',
      `Found ${triggers.length} total trigger(s).\n\n` +
      '🚨 This will delete ALL triggers in this project,\n' +
      'not just email triggers!\n\n' +
      'Are you absolutely sure?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.NO) {
      console.log('User cancelled ALL trigger deletion');
      return;
    }
    
    // Delete all triggers
    let deletedCount = 0;
    triggers.forEach(trigger => {
      try {
        ScriptApp.deleteTrigger(trigger);
        console.log(`Deleted trigger: ${trigger.getHandlerFunction()} - ${trigger.getUniqueId()}`);
        deletedCount++;
      } catch (error) {
        console.error(`Error deleting trigger:`, error);
      }
    });
    
    ui.alert('All Triggers Cancelled', 
      `✅ Successfully cancelled ${deletedCount} trigger(s).\n\n` +
      '🚨 ALL automatic functions have been stopped.',
      ui.ButtonSet.OK);
    
    console.log(`✓ Successfully deleted ${deletedCount} triggers`);
    
  } catch (error) {
    console.error('Error in cancelAllTriggers:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Error', 
      `❌ Failed to cancel triggers:\n${error.toString()}`,
      ui.ButtonSet.OK);
  }
}