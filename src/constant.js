const CONFIG = {
  // ADMIN_EMAIL: 'sangmt@bcons.com.vn',
  ADMIN_EMAIL: 'tin.daominh@gmail.com',
  TRIGGER_HOUR: 9, //
  DAYS_BEFORE_DEADLINE: 7,
  SHEET_NAME: 'file dữ liệu gửi mail tự động' 
};

// Column headers
const COLUMN_HEADERS = {
  EMAIL: 'email',
  SUBJECT: 'Tiêu đề', 
  CONTENT: 'nội dung',
  ATTACHMENT: 'File đính kèm',
  START_DATE: 'Ngày bắt đầu',  // Column E
  DEADLINE: 'Deadline',        // Column F
  STATUS: 'trạng thái'
};

// Email templates
const EMAIL_TEMPLATES = {
  URGENCY_MESSAGES: {
    OVERDUE: '⚠️ TASK IS {days} DAYS OVERDUE!',
    DUE_TODAY: '🔴 TASK IS DUE TODAY!',
    DUE_SOON: '⏰ {days} days remaining until deadline!',
    DUE_NORMAL: '📅 {days} days remaining until deadline.'
  },
  
  STYLES: {
    OVERDUE: 'background-color: #ff4444; color: white; padding: 10px; border-radius: 5px;',
    DUE_TODAY: 'background-color: #ff6b6b; color: white; padding: 10px; border-radius: 5px;',
    DUE_SOON: 'background-color: #ffc107; color: black; padding: 10px; border-radius: 5px;',
    DUE_NORMAL: 'background-color: #4CAF50; color: white; padding: 10px; border-radius: 5px;'
  }
};

// Status messages
const STATUS_MESSAGES = {
  SENT: 'Sent',
  TEST_SENT: 'Test - Sent',
  ERROR: 'Send Error',
  PROCESSING: 'Processing'
};