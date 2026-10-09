// TEST PUSH for Jorge


// ================================
// CONFIG / CONSTANTS
// ================================

const ngtlemail = 'ngtl@akamai.com';
const mailArchitect = 'pnajduch@akamai.com';
const mailPgM = 'kgrzeszc@akamai.com';
const ngtlsenderName = 'The NGTL Team'

// TEST MODE
const IS_TEST_MODE = false; // test TRUE; production FALSE
const TEST_EMAIL = mailPgM;
const TEST_CC_EMAIL = '';

const MSG_SENT = true;

// Flow Types
const FLOW_TYPE = {
  INVITATION: "invitation", // (for further development)
  CONFIRMATION: "confirmation",
  PLACEHOLDER: "placeholder",
  VP_SELECTION: "vp_selection",
  VP_REMINDER: "vp_reminder",
  CALENDAR_INVITE: "calendar_invite"
};

// Flow Sheets 
const FLOW_SHEETS = {
  [FLOW_TYPE.INVITATION]:   { name: "Presenters", dataStartRow: 2 }, // (for further development)
  [FLOW_TYPE.CONFIRMATION]: { name: "Presenters", dataStartRow: 2 },
  [FLOW_TYPE.PLACEHOLDER]: { name: "Presenters", dataStartRow: 2 },
  [FLOW_TYPE.VP_SELECTION]: { name: "Leadership", dataStartRow: 2 },
  [FLOW_TYPE.VP_REMINDER]:  { name: "Leadership", dataStartRow: 2 },
  [FLOW_TYPE.CALENDAR_INVITE]: { name: "Presenters", dataStartRow: 2 }
};

const LOG_SHEET_NAME = "Email LOG";

// Trainers / Presenters
const COL = {
  NO: 0,                           // col A
  SESSION_NAME: 1,                 // col B
  PRESENTER: 2,                    // col C
  COHORT: 3,                       // col D
  SESSION_DATE: 4,                 // col E
  SESSION_TIME: 5,                 // col F
  DURATION: 6,                     // col G
  PRESENTER_EMAIL: 7,              // col H
  PRESENTER_TIMEZONE: 9,           // col J
  PRESENTER_TIME: 10,              // col K
  WEBEX_LINK: 11,                  // col L
  INVITATION_CHECKBOX: 12,         // col M
  CONFIRMATION_CHECKBOX: 14,       // col O
  PLACEHOLDER_CHECKBOX: 16,        // col Q
  INVITE_CHECKBOX: 17               // col R
};

// VP Selection
const VP_COL = {
  ORG: 0,                  // col A (Department / Org)
  VP_NAME: 1,              // col B
  VP_EMAIL: 2,             // col C
  SPREADSHEET_LINK: 3,     // col D
  TOTAL_SEATS: 4,          // col E
  AMER_SEATS: 5,           // col F
  APJ_SEATS: 6,            // col G
  EMEA_SEATS: 7,           // col H
  WAITING_LIST: 8,         // col I
  WAITING_SEATS: 9,        // col J
  EMAIL_SENT_CHECKBOX: 10,  // col K 
  REMINDER_SENT_CHECKBOX: 11,  // col L
  PARTICIPANT_SELECTED: 12  // col M
};

// LOGS
const LOG_COL = {
  TIMESTAMP: 0,     // col A
  SENDER: 1,        // col B
  RECIPIENT: 2,     // col C
  EMAIL_TYPE: 3,    // col D
  COHORT: 4,        // col E
  SUBJECT: 5,       // col F
  SEND_STATUS: 6,   // col G
  NOTES: 7          // col H
};