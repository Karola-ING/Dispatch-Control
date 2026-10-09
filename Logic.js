// ================================
// CORE ENGINE
// ================================


// Main function that processes shipments from the UI
function processSending(selectedItems, flow) {
  let mailCount = 0;

  switch (flow) {
    case FLOW_TYPE.VP_SELECTION:
      mailCount = processVpSelectionFlow(selectedItems, flow);
      break;

    case FLOW_TYPE.VP_REMINDER:
      mailCount = processVpReminderFlow(selectedItems, flow);
      break;
      
    case FLOW_TYPE.CONFIRMATION:
    case FLOW_TYPE.INVITATION:
      mailCount = processPresenterFlows(selectedItems, flow);
      break;

    case FLOW_TYPE.PLACEHOLDER:
      mailCount = processPlaceholderFlow(selectedItems, flow);
      break;

    case FLOW_TYPE.CALENDAR_INVITE: // 
      mailCount = processCalendarInviteFlow(selectedItems, flow);
      break;
      
    default:
      throw new Error(`Unsupported flow type: ${flow}`);
  }

  return `Sending complete! You have sent ${mailCount} email(s).`;
}


// ================================
// DEDICATED FLOW PROCESSORS
// ================================


// Retrieves data and groups it by Presenter and Session
function processPresenterFlows(selectedItems, flow) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = FLOW_SHEETS[flow];
  const presenterSheet = spreadsheet.getSheetByName(sheetConfig.name);
  
  if (!presenterSheet) throw new Error(`Sheet not found: ${sheetConfig.name}`);
  
  const data = presenterSheet.getDataRange().getValues();
  const senderEmail = Session.getActiveUser().getEmail();
  let count = 0;

  selectedItems.forEach(group => {
    const sessionRows = group.rowsIndices.map(rowIndex => data[rowIndex - 1]);
    if (sessionRows.length === 0) return;

    let mailMetadata = null;
    let isSuccess = false;
    let errorMessage = "";

    try {
      if (flow === FLOW_TYPE.CONFIRMATION) {
        mailMetadata = executeConfirmationFlow(sessionRows);
      } else if (flow === FLOW_TYPE.INVITATION) {
        mailMetadata = executeInvitationFlow(sessionRows);
      }
      isSuccess = true;
    } catch (error) {
      isSuccess = false;
      errorMessage = error.toString();
      mailMetadata = {
        recipient: sessionRows[0][COL.PRESENTER_EMAIL],
        subject: `[FAILED] Flow: ${flow}`,
        cohorts: "Error Status"
      };
    }

    appendEmailLog({
      sender: senderEmail,
      recipient: mailMetadata.recipient,
      flowType: flow,
      cohorts: mailMetadata.cohorts,
      subject: mailMetadata.subject,
      status: isSuccess,
      notes: errorMessage
    });

    if (isSuccess) {
      group.rowsIndices.forEach(rowIndex => {
        presenterSheet.getRange(rowIndex, COL.CONFIRMATION_CHECKBOX + 1).setValue(MSG_SENT);
      });
      count++;
    }
  });

  return count;
}


// VPs Selection Processor
function processVpSelectionFlow(selectedItems, flow) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = FLOW_SHEETS[flow];
  const vpSheet = spreadsheet.getSheetByName(sheetConfig.name);
  
  if (!vpSheet) throw new Error(`Sheet not found: ${sheetConfig.name}`);
  
  const allVpData = vpSheet.getDataRange().getValues();
  const senderEmail = Session.getActiveUser().getEmail();
  
  // Benchmark table
  const fullDepartmentTableRaw = allVpData.slice(sheetConfig.dataStartRow - 1).map(row => ({
    org: row[VP_COL.ORG],
    total: row[VP_COL.TOTAL_SEATS],
    apj: row[VP_COL.APJ_SEATS],
    emea: row[VP_COL.EMEA_SEATS],
    amer: row[VP_COL.AMER_SEATS]
  }));

  let count = 0;
  selectedItems.forEach(item => {
    const row = vpSheet.getRange(item.rowIndex, 1, 1, vpSheet.getLastColumn()).getValues()[0];
    
    let mailMetadata = null;
    let isSuccess = false;
    let errorMessage = "";

    try {
      mailMetadata = executeVpSelectionFlow(row, fullDepartmentTableRaw);
      isSuccess = true;
    } catch (error) {
      isSuccess = false;
      errorMessage = error.toString();
      mailMetadata = {
        recipient: row[VP_COL.VP_EMAIL],
        subject: "[FAILED] VP Selection Flow",
        cohorts: ""
      };
    }

    appendEmailLog({
      sender: senderEmail,
      recipient: mailMetadata.recipient,
      flowType: flow,
      cohorts: mailMetadata.cohorts, // Będzie puste zgodnie z wytyczną
      subject: mailMetadata.subject,
      status: isSuccess,
      notes: errorMessage
    });

    if (isSuccess) {
      vpSheet.getRange(item.rowIndex, VP_COL.EMAIL_SENT_CHECKBOX + 1).setValue(MSG_SENT);
      count++;
    }
  });

  return count;
}

// VPs Reminder Processor
function processVpReminderFlow(selectedItems, flow) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = FLOW_SHEETS[flow];
  const vpSheet = spreadsheet.getSheetByName(sheetConfig.name);
  
  if (!vpSheet) throw new Error(`Sheet not found: ${sheetConfig.name}`);
  
  const allVpData = vpSheet.getDataRange().getValues();
  const senderEmail = Session.getActiveUser().getEmail();
  
  // Benchmark table
  const fullDepartmentTableRaw = allVpData.slice(sheetConfig.dataStartRow - 1).map(row => ({
    org: row[VP_COL.ORG],
    total: row[VP_COL.TOTAL_SEATS],
    apj: row[VP_COL.APJ_SEATS],
    emea: row[VP_COL.EMEA_SEATS],
    amer: row[VP_COL.AMER_SEATS]
  }));

  let count = 0;
  selectedItems.forEach(item => {
    const row = vpSheet.getRange(item.rowIndex, 1, 1, vpSheet.getLastColumn()).getValues()[0];
    
    let mailMetadata = null;
    let isSuccess = false;
    let errorMessage = "";

    try {
      mailMetadata = executeVpReminderFlow(row, fullDepartmentTableRaw);
      isSuccess = true;
    } catch (error) {
      isSuccess = false;
      errorMessage = error.toString();
      mailMetadata = {
        recipient: row[VP_COL.VP_EMAIL],
        subject: "[FAILED] VP Reminder Flow",
        cohorts: "N/A"
      };
    }

    appendEmailLog({
      sender: senderEmail,
      recipient: mailMetadata.recipient,
      flowType: flow,
      cohorts: mailMetadata.cohorts, 
      subject: mailMetadata.subject,
      status: isSuccess,
      notes: errorMessage
    });

    if (isSuccess) {
      vpSheet.getRange(item.rowIndex, VP_COL.REMINDER_SENT_CHECKBOX + 1).setValue(MSG_SENT);
      count++;
    }
  });

  return count;
}


// Placeholder creation flow
function processPlaceholderFlow(selectedItems, flow) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = FLOW_SHEETS[flow];
  const presenterSheet = spreadsheet.getSheetByName(sheetConfig.name);
  
  if (!presenterSheet) throw new Error(`Sheet not found: ${sheetConfig.name}`);
  
  const data = presenterSheet.getDataRange().getValues();
  const senderEmail = Session.getActiveUser().getEmail();
  let count = 0;

  selectedItems.forEach(group => {
    const sessionRows = group.rowsIndices.map(rowIndex => data[rowIndex - 1]);
    if (sessionRows.length === 0) return;

    let mailMetadata = null;
    let isSuccess = false;
    let errorMessage = "";

    try {
      mailMetadata = executeCalendarPlaceholderFlow(sessionRows);
      isSuccess = true;
    } catch (error) {
      isSuccess = false;
      errorMessage = error.toString();
      mailMetadata = {
        recipient: sessionRows[0][COL.PRESENTER_EMAIL],
        subject: `[FAILED ICAL] Placeholder: ${sessionRows[0][COL.SESSION_NAME]}`,
        cohorts: "Error"
      };
    }

    appendEmailLog({
      sender: senderEmail,
      recipient: mailMetadata.recipient,
      flowType: flow,
      cohorts: mailMetadata.cohorts,
      subject: mailMetadata.subject,
      status: isSuccess,
      notes: errorMessage
    });

    if (isSuccess) {
      group.rowsIndices.forEach(rowIndex => {
        presenterSheet.getRange(rowIndex, COL.PLACEHOLDER_CHECKBOX + 1).setValue(MSG_SENT);
      });
      count++;
    }
  });

  return count;
}


// Prawdziwe zaproszenia z Webex i wylistowaniem kohorty
function processCalendarInviteFlow(selectedItems, flow) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetConfig = FLOW_SHEETS[flow];
  const presenterSheet = spreadsheet.getSheetByName(sheetConfig.name);
  
  if (!presenterSheet) throw new Error(`Sheet not found: ${sheetConfig.name}`);
  
  const data = presenterSheet.getDataRange().getValues();
  const senderEmail = Session.getActiveUser().getEmail();
  let count = 0;

  selectedItems.forEach(group => {
    const sessionRows = group.rowsIndices.map(rowIndex => data[rowIndex - 1]);
    if (sessionRows.length === 0) return;

    let mailMetadata = null;
    let isSuccess = false;
    let errorMessage = "";

    try {
      // Wywołuje funkcję z Emails.js
      mailMetadata = executeCalendarInviteFlow(sessionRows);
      isSuccess = true;
    } catch (error) {
      isSuccess = false;
      errorMessage = error.toString();
      mailMetadata = {
        recipient: "Cohort Participants",
        subject: `[FAILED INVITE] Session: ${sessionRows[0][COL.SESSION_NAME]}`,
        cohorts: "Error"
      };
    }

    appendEmailLog({
      sender: senderEmail,
      recipient: mailMetadata.recipient,
      flowType: flow,
      cohorts: mailMetadata.cohorts,
      subject: mailMetadata.subject,
      status: isSuccess,
      notes: errorMessage
    });

    if (isSuccess) {
      group.rowsIndices.forEach(rowIndex => {
        // Zaznaczamy Checkbox "INVITE" w kolumnie R
        presenterSheet.getRange(rowIndex, COL.INVITE_CHECKBOX + 1).setValue(MSG_SENT);
      });
      count++;
    }
  });

  return count;
}


function getGroupedInviteData() {
  const sheetConfig = FLOW_SHEETS[FLOW_TYPE.CALENDAR_INVITE];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetConfig.name);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < sheetConfig.dataStartRow) return [];

  const data = sheet.getRange(sheetConfig.dataStartRow, 1, lastRow - (sheetConfig.dataStartRow - 1), sheet.getLastColumn()).getValues();
  const grouped = {};

  data.forEach((row, index) => {
    const isSent = row[COL.INVITE_CHECKBOX]; // Filtrujemy po kolumnie R (INVITE_CHECKBOX)
    const sessionName = row[COL.SESSION_NAME];
    const presenter = row[COL.PRESENTER];
    const presenterEmail = row[COL.PRESENTER_EMAIL];

    if (!sessionName || isSent === MSG_SENT) return;

    const groupKey = `${presenter}|||${sessionName}`;

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        presenter: presenter,
        presenterEmail: presenterEmail,
        sessionName: sessionName,
        rowsIndices: []
      };
    }
    grouped[groupKey].rowsIndices.push(index + sheetConfig.dataStartRow);
  });

  return Object.keys(grouped).map(key => grouped[key]);
}

// ================================
// DEDICATED DATA FETCHERS (UI)
// ================================


function getGroupedSessionsData() {
  const sheetConfig = FLOW_SHEETS[FLOW_TYPE.CONFIRMATION];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetConfig.name);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < sheetConfig.dataStartRow) return [];

  const data = sheet.getRange(sheetConfig.dataStartRow, 1, lastRow - (sheetConfig.dataStartRow - 1), sheet.getLastColumn()).getValues();
  const grouped = {};

  data.forEach((row, index) => {
    const isSent = row[COL.CONFIRMATION_CHECKBOX];
    const sessionName = row[COL.SESSION_NAME];
    const presenter = row[COL.PRESENTER];
    const presenterEmail = row[COL.PRESENTER_EMAIL];

    if (!sessionName || isSent === MSG_SENT) return;

    const groupKey = `${presenter}|||${sessionName}`;

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        presenter: presenter,
        presenterEmail: presenterEmail,
        sessionName: sessionName,
        rowsIndices: []
      };
    }
    // Calculate the exact line number based on the starting line from the configuration
    grouped[groupKey].rowsIndices.push(index + sheetConfig.dataStartRow);
  });

  return Object.keys(grouped).map(key => grouped[key]);
}

function getVpSelectionData() {
  const sheetConfig = FLOW_SHEETS[FLOW_TYPE.VP_SELECTION];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetConfig.name);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < sheetConfig.dataStartRow) return [];

  const data = sheet.getRange(sheetConfig.dataStartRow, 1, lastRow - (sheetConfig.dataStartRow - 1), sheet.getLastColumn()).getValues();
  const vps = [];

  data.forEach((row, index) => {
    const isSent = row[VP_COL.EMAIL_SENT_CHECKBOX];
    const vpName = row[VP_COL.VP_NAME];
    const org = row[VP_COL.ORG];

    if (!vpName || isSent === MSG_SENT) return;

    vps.push({
      vpName: vpName,
      org: org,
      total: row[VP_COL.TOTAL_SEATS],
      apj: row[VP_COL.APJ_SEATS],
      amer: row[VP_COL.AMER_SEATS],
      emea: row[VP_COL.EMEA_SEATS],
      hasWaiting: row[VP_COL.WAITING_LIST] === true,
      waitingSeats: row[VP_COL.WAITING_SEATS],
      rowIndex: index + sheetConfig.dataStartRow
    });
  });

  return vps;
}

function getVpReminderData() {
  const sheetConfig = FLOW_SHEETS[FLOW_TYPE.VP_REMINDER];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetConfig.name);
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < sheetConfig.dataStartRow) return [];

  const data = sheet.getRange(sheetConfig.dataStartRow, 1, lastRow - (sheetConfig.dataStartRow - 1), sheet.getLastColumn()).getValues();
  const vps = [];

  data.forEach((row, index) => {
    const isParticipantSelected = row[VP_COL.PARTICIPANT_SELECTED];
    const vpName = row[VP_COL.VP_NAME];
    const org = row[VP_COL.ORG];

    if (!vpName) return;

    // Only those who didn't chose participants
    if (isParticipantSelected === true) return;

    vps.push({
      vpName: vpName,
      org: org,
      total: row[VP_COL.TOTAL_SEATS],
      apj: row[VP_COL.APJ_SEATS],
      amer: row[VP_COL.AMER_SEATS],
      emea: row[VP_COL.EMEA_SEATS],
      hasWaiting: row[VP_COL.WAITING_LIST] === true,
      waitingSeats: row[VP_COL.WAITING_SEATS],
      rowIndex: index + sheetConfig.dataStartRow
    });
  });

  return vps;
}


function getGroupedPlaceholderData() {
  const sheetConfig = FLOW_SHEETS[FLOW_TYPE.PLACEHOLDER];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetConfig.name);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < sheetConfig.dataStartRow) return [];

  const data = sheet.getRange(sheetConfig.dataStartRow, 1, lastRow - (sheetConfig.dataStartRow - 1), sheet.getLastColumn()).getValues();
  const grouped = {};

  data.forEach((row, index) => {
    const isSent = row[COL.PLACEHOLDER_CHECKBOX]; // Filtrujemy po kolumnie Q
    const sessionName = row[COL.SESSION_NAME];
    const presenter = row[COL.PRESENTER];
    const presenterEmail = row[COL.PRESENTER_EMAIL];

    if (!sessionName || isSent === MSG_SENT) return;

    const groupKey = `${presenter}|||${sessionName}`;

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        presenter: presenter,
        presenterEmail: presenterEmail,
        sessionName: sessionName,
        rowsIndices: []
      };
    }
    grouped[groupKey].rowsIndices.push(index + sheetConfig.dataStartRow);
  });

  return Object.keys(grouped).map(key => grouped[key]);
}

// ================================
// ENGINE HELPERS
// ================================


// Presenters' names
function getPresenterGreetingNames(fullPresentersString) {
  if (!fullPresentersString) return "Presenter";

  // More than one presenter in PRESENTER cell 
  const presenters = fullPresentersString.split(",").map(p => p.trim());
  const resolvedNames = [];

  presenters.forEach(presenter => {
    const parts = presenter.split(/\s+/); // Divide by spaces
    
    // Find variable that is not an initial (longer than 2 letters)
    let firstName = parts.find(part => part.length > 2 && !part.includes("."));
    
    // If not found, take the first part of the string
    if (parts.length > 1) {
      // Usuwamy tylko ostatni element (nazwisko) i łączymy resztę z powrotem
      const firstName = parts.slice(0, -1).join(" ");
      resolvedNames.push(firstName);
    } else {
      // Jeśli to tylko jedno słowo, bierzemy je w całości
      resolvedNames.push(presenter);
    }
  });

  // Combine names based on the number of presenters
  if (resolvedNames.length === 0) return "Presenter";
  if (resolvedNames.length === 1) return resolvedNames[0];
  if (resolvedNames.length === 2) return `${resolvedNames[0]} and ${resolvedNames[1]}`;
  
  // For more than 2 presenters
  return resolvedNames.slice(0, -1).join(", ") + " and " + resolvedNames[resolvedNames.length - 1];
}


// Log printing
function appendEmailLog(logObject) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = spreadsheet.getSheetByName(LOG_SHEET_NAME);
  
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(["Timestamp", "Sender email", "Recipient email", "Email type", "Cohort", "Subject", "Send", "Notes"]);
  }

  // Przygotowanie pustego wiersza o długości zdefiniowanej w konfiguracji mappingu kolumn
  const columnsCount = Object.keys(LOG_COL).length;
  const newRowValues = new Array(columnsCount).fill("");

  // Wypełnianie komórek danymi na bazie stałych indeksów z Config.gs
  newRowValues[LOG_COL.TIMESTAMP]   = new Date(); // Generuje aktualną datę i czas triggeru
  newRowValues[LOG_COL.SENDER]      = logObject.sender;
  newRowValues[LOG_COL.RECIPIENT]   = logObject.recipient;
  newRowValues[LOG_COL.EMAIL_TYPE]  = logObject.flowType;
  newRowValues[LOG_COL.COHORT]      = logObject.cohorts;
  newRowValues[LOG_COL.SUBJECT]     = logObject.subject;
  newRowValues[LOG_COL.SEND_STATUS] = logObject.status; // Zapisze TRUE / FALSE (jako natywny typ logiczny arkusza)
  newRowValues[LOG_COL.NOTES]       = logObject.notes;

  logSheet.appendRow(newRowValues);
}

// Time Calculation
function calculateEndTime(startTimeStr, durationMinutes) {
  try {
    // Łapiemy cyfry, AM/PM oraz opcjonalną strefę czasową (np. ET, PT)
    const match = startTimeStr.match(/(\d+):(\d+)\s*(AM|PM)(?:\s*([A-Z]{2,4}))?/i);
    if (!match) return startTimeStr; 
    
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const ampm = match[3].toUpperCase();
    const timezone = match[4] ? match[4].toUpperCase() : ""; 
    
    // Tworzymy czysty czas startu bez strefy czasowej
    const paddedStartHours = hours < 10 ? `0${hours}` : hours;
    const paddedStartMinutes = minutes < 10 ? `0${minutes}` : minutes;
    const cleanStartTime = `${paddedStartHours}:${paddedStartMinutes} ${ampm}`;
    
    // Konwersja na format 24-godzinny do obliczeń
    if (ampm === "PM" && hours < 12) hours += 12;
    if (ampm === "AM" && hours === 12) hours = 0;
    
    // Obliczanie nowego czasu
    const date = new Date();
    date.setHours(hours, minutes + parseInt(durationMinutes, 10));
    
    let endHours = date.getHours();
    let endMinutes = date.getMinutes();
    const endAmpm = endHours >= 12 ? "PM" : "AM";
    
    // Powrót do formatu 12-godzinnego
    endHours = endHours % 12;
    if (endHours === 0) endHours = 12;
    if (endMinutes < 10) endMinutes = `0${endMinutes}`;
    
    const paddedEndHours = endHours < 10 ? `0${endHours}` : endHours;
    
    // Tworzymy czas końca z doklejoną strefą na samym końcu
    const endTimeWithZone = `${paddedEndHours}:${endMinutes} ${endAmpm}${timezone ? ' ' + timezone : ''}`;
    
    // ZWRACAMY GOTOWY STRING (np. "11:00 AM - 12:30 PM ET")
    return `${cleanStartTime} - ${endTimeWithZone}`;
    
  } catch(e) {
    return startTimeStr;
  }
}

// iCal time format
function convertTo24hICal(timeStr) {
  try {
    const match = timeStr.match(/(\d+):(\d+)\s*(AM|PM)/i);
    if (!match) return "120000"; // fallback na południe w razie błędu parsowania
    
    let hours = parseInt(match[1], 10);
    const minutes = parseInt(match[2], 10);
    const ampm = match[3].toUpperCase();
    
    if (ampm === "PM" && hours < 12) hours += 12;
    if (ampm === "AM" && hours === 12) hours = 0;
    
    const paddedHours = hours < 10 ? `0${hours}` : hours;
    const paddedMinutes = minutes < 10 ? `0${minutes}` : minutes;
    
    return `${paddedHours}${paddedMinutes}00`; // Dodajemy 00 sekund na końcu
  } catch(e) {
    return "120000";
  }
}


// FLOW: OUTLOOK CALENDAR INVITE (PRAWDZIWE ZAPROSZENIE Z WEBEXEM I LOGO)
function executeCalendarInviteFlow(sessionRows, cohortEmails) {
  const fullPresenterName = sessionRows[0][COL.PRESENTER];
  const presenterEmail = sessionRows[0][COL.PRESENTER_EMAIL];
  const sessionName = sessionRows[0][COL.SESSION_NAME];
  const sessionNo = sessionRows[0][COL.NO];

  const activeUserEmail = Session.getActiveUser().getEmail(); 
  const iCalOrganizerEmail = IS_TEST_MODE ? "ngtl-system@akamai.com" : ngtlemail;

  // Główny tytuł do logowania
  const subject = `NGTL ${sessionNo}: ${sessionName}`;

  // Pobranie logo jako Blob
  const logoUrl = "https://akamai-university.akamaized.net/NGTL_logo.png";
  let logoBlob;
  try {
    logoBlob = UrlFetchApp.fetch(logoUrl).getBlob().setName("logo.png");
  } catch(e) {
    // W razie gdyby akamai padło, wrzucamy pusty blob by skrypt się nie wywalił
    logoBlob = Utilities.newBlob(" ", "image/png", "logo.png");
  }

  const cohortNames = [];
  let eventsIcs = "";

  sessionRows.forEach(row => {
    const cohortName = row[COL.COHORT];
    const webexLink = row[COL.WEBEX_LINK]; // Link do Webex z kolumny L – osobny dla każdej kohorty
    if (!cohortNames.includes(cohortName)) cohortNames.push(cohortName);

    const sessionDateRaw = row[COL.SESSION_DATE]; 
    const presenterTimeStr = row[COL.PRESENTER_TIME]; 
    const durationMinutes = parseInt(row[COL.DURATION], 10);
    
    let rawTimezone = row[COL.PRESENTER_TIMEZONE] || "UTC";
    
    let iCalTimezone = "America/New_York"; 
    if (rawTimezone.includes("PT")) iCalTimezone = "America/Los_Angeles";
    if (rawTimezone.includes("CET")) iCalTimezone = "Europe/Warsaw";
    if (rawTimezone.includes("GMT")) iCalTimezone = "Europe/London";

    const cleanTimeMatch = presenterTimeStr.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    let parsedStartDate = new Date(sessionDateRaw);
    
    if (cleanTimeMatch) {
      let hours = parseInt(cleanTimeMatch[1], 10);
      const minutes = parseInt(cleanTimeMatch[2], 10);
      const ampm = cleanTimeMatch[3].toUpperCase();

      if (ampm === "PM" && hours < 12) hours += 12;
      if (ampm === "AM" && hours === 12) hours = 0;

      parsedStartDate.setHours(hours, minutes, 0, 0);
    }

    let parsedEndDate = new Date(parsedStartDate.getTime() + durationMinutes * 60000);

    const iCalStartStr = Utilities.formatDate(parsedStartDate, Session.getScriptTimeZone(), "yyyyMMdd'T'HHmmss");
    const iCalEndStr = Utilities.formatDate(parsedEndDate, Session.getScriptTimeZone(), "yyyyMMdd'T'HHmmss");

    const formattedDateText = Utilities.formatDate(parsedStartDate, Session.getScriptTimeZone(), "E, MMM d, yyyy");

    // Zgodnie z wytycznymi - tytuł taki sam jak placeholder, tylko bez [PLACEHOLDER]
    const iterationSubject = `NGTL ${sessionNo}: ${sessionName} (${cohortName})`;

    // ICS będzie zawierał tych samych gości
    // Dodajemy Webex do lokalizacji
    eventsIcs = 
      "BEGIN:VCALENDAR\r\n" +
      "VERSION:2.0\r\n" +
      "PRODID:-//Google Inc//Google Apps Script//EN\r\n" +
      "METHOD:REQUEST\r\n" + 
      "BEGIN:VEVENT\r\n" +
      "UID:NGTL-" + Utilities.getUuid() + "\r\n" +
      "ORGANIZER;CN=" + ngtlsenderName + ":MAILTO:" + iCalOrganizerEmail + "\r\n";
      
    // Doklejamy uczestników z kohorty do pliku ICS, aby się wyświetlali
    if (IS_TEST_MODE) {
      // W trybie testowym dodajemy adres testowy tylko jeden raz
      eventsIcs += "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=" + mailPgM + ":MAILTO:" + mailPgM + "\r\n";
    } else {
      // W trybie produkcyjnym dodajemy prawdziwych uczestników po kolei
      cohortEmails.forEach(email => {
        let participantEmail = email.trim();
        eventsIcs += "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=" + participantEmail + ":MAILTO:" + participantEmail + "\r\n";
      });
    }

    eventsIcs +=
      "DTSTART;TZID=" + iCalTimezone + ":" + iCalStartStr + "\r\n" + 
      "DTEND;TZID=" + iCalTimezone + ":" + iCalEndStr + "\r\n" +
      "SUMMARY:" + iterationSubject + "\r\n" +
      "LOCATION:" + webexLink + "\r\n" +
      "DESCRIPTION:You're invited to an NGTL session:\\n\\n" + sessionName + " – presented by " + fullPresenterName + "\\n\\nWebex Link: " + webexLink + "\\r\n" +
      "SEQUENCE:0\r\n" +
      "STATUS:CONFIRMED\r\n" +
      "TRANSP:OPAQUE\r\n" + 
      "END:VEVENT\r\n" +
      "END:VCALENDAR";

    // Budujemy dedykowaną treść HTML zgodnie z wytycznymi
    const htmlBody = `
      <html>
        <body style="line-height: 1.5; padding: 0 20px; font-family: Arial, sans-serif; color: #333;">
          <p>You're invited to an NGTL session:</p>
          <p><strong>${sessionName}</strong> – presented by <strong>${fullPresenterName}</strong></p>
          <img src="cid:ngtlLogo" alt="NGTL Logo" style="max-width: 200px; margin-bottom: 20px;" />
        </body>
      </html>
    `;

    // Wysyłanie e-maila
    // Używam MailApp zamiast mailSenderAppWithCalendar, aby móc dodać inlineImages
    const recipientList = IS_TEST_MODE ? TEST_EMAIL : cohortEmails.join(',');
    const ccrecipients = `${ngtlemail}, ${mailArchitect}, ${mailPgM}`;
    const ccList = IS_TEST_MODE ? TEST_CC_EMAIL : ccrecipients;
    const finalSubject = IS_TEST_MODE ? `[TEST CALENDAR] ${iterationSubject}` : iterationSubject;

    const icsAttachment = Utilities.newBlob(eventsIcs, 'text/calendar; method=REQUEST; charset=UTF-8', 'invite.ics');

    MailApp.sendEmail({
      name: ngtlsenderName,
      replyTo: ngtlemail,
      to: recipientList,
      cc: ccList,
      subject: finalSubject,
      htmlBody: htmlBody,
      inlineImages: {
        ngtlLogo: logoBlob
      },
      attachments: [icsAttachment]
    });
  });

  return {
    recipient: "Cohort Participants", // Lub możesz wstawić `cohortEmails.join(", ")`
    subject: subject,
    cohorts: cohortNames.join(", ")
  };
}



// ================================
// PARTICIPANTS DATA FETCHER
// ================================

/**
 * Funkcja pomocnicza: Pobiera uczestników z zakładki "Participants" 
 * i grupuje ich w kohorty (Americas, APJ, EMEA).
 */
function getCohortEmailsMap_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Participants");
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const map = {};

  // Zakładamy, że wiersz 0 to nagłówki, więc zaczynamy od i = 1
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][1]).toLowerCase().trim(); // Kolumna B (indeks 1) - Status
    const email = String(data[i][3]).trim();                // Kolumna D (indeks 3) - Email
    let group = String(data[i][9]).trim();                  // Kolumna J (indeks 9) - Projekt/Grupa

    // Bierzemy tylko "approved" i sprawdzamy, czy mail w ogóle istnieje
    if (status !== "approved" || !email) continue;

    // Przekształcamy "Americas 1" -> "Americas", rozbijając po spacji i biorąc pierwsze słowo
    let cohort = group.split(" ")[0]; 

    // Wyjątek: Ashwini Saket zawsze trafia do Americas
    if (email.toLowerCase() === "asaket@akamai.com") {
      cohort = "Americas";
    }

    if (!map[cohort]) {
      map[cohort] = [];
    }
    map[cohort].push(email);
  }
  
  return map;
}

// ================================
// FLOW: OUTLOOK CALENDAR INVITE
// ================================

function executeCalendarInviteFlow(sessionRows) {
  const fullPresenterName = sessionRows[0][COL.PRESENTER];
  const sessionName = sessionRows[0][COL.SESSION_NAME];
  const sessionNo = sessionRows[0][COL.NO];

  const activeUserEmail = Session.getActiveUser().getEmail(); 
  const iCalOrganizerEmail = IS_TEST_MODE ? "ngtl-system@akamai.com" : ngtlemail;

  // Główny tytuł do logowania
  const subject = `NGTL ${sessionNo}: ${sessionName}`;

  // Pobranie logo jako Blob
  const logoUrl = "https://akamai-university.akamaized.net/NGTL_logo.png";
  let logoBlob;
  try {
    logoBlob = UrlFetchApp.fetch(logoUrl).getBlob().setName("logo.png");
  } catch(e) {
    // Awaryjnie, gdyby serwer akamai nie odpowiedział
    logoBlob = Utilities.newBlob(" ", "image/png", "logo.png");
  }

  const cohortNames = [];
  let eventsIcs = "";

  // 1. Pobranie zgrupowanych maili (tylko approved)
  const cohortEmailsMap = getCohortEmailsMap_();

  sessionRows.forEach(row => {
    const cohortName = row[COL.COHORT];
    const webexLink = row[COL.WEBEX_LINK]; // Link do Webex z kolumny L – osobny dla każdej kohorty
    if (!cohortNames.includes(cohortName)) cohortNames.push(cohortName);

    // 2. Wyciągamy maile dla akurat procesowanej kohorty (np. "Americas")
    const cohortEmails = cohortEmailsMap[cohortName] || [];

    const sessionDateRaw = row[COL.SESSION_DATE]; 
    const presenterTimeStr = row[COL.PRESENTER_TIME]; 
    const durationMinutes = parseInt(row[COL.DURATION], 10);
    
    let rawTimezone = row[COL.PRESENTER_TIMEZONE] || "UTC";
    
    let iCalTimezone = "America/New_York"; 
    if (rawTimezone.includes("PT")) iCalTimezone = "America/Los_Angeles";
    if (rawTimezone.includes("CET")) iCalTimezone = "Europe/Warsaw";
    if (rawTimezone.includes("GMT")) iCalTimezone = "Europe/London";

    const cleanTimeMatch = presenterTimeStr.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
    let parsedStartDate = new Date(sessionDateRaw);
    
    if (cleanTimeMatch) {
      let hours = parseInt(cleanTimeMatch[1], 10);
      const minutes = parseInt(cleanTimeMatch[2], 10);
      const ampm = cleanTimeMatch[3].toUpperCase();

      if (ampm === "PM" && hours < 12) hours += 12;
      if (ampm === "AM" && hours === 12) hours = 0;

      parsedStartDate.setHours(hours, minutes, 0, 0);
    }

    let parsedEndDate = new Date(parsedStartDate.getTime() + durationMinutes * 60000);

    const iCalStartStr = Utilities.formatDate(parsedStartDate, Session.getScriptTimeZone(), "yyyyMMdd'T'HHmmss");
    const iCalEndStr = Utilities.formatDate(parsedEndDate, Session.getScriptTimeZone(), "yyyyMMdd'T'HHmmss");
    const formattedDateText = Utilities.formatDate(parsedStartDate, Session.getScriptTimeZone(), "E, MMM d, yyyy");

    // Tytuł zaproszenia i maila
    const iterationSubject = `NGTL ${sessionNo}: ${sessionName} (${cohortName})`;

    eventsIcs = 
      "BEGIN:VCALENDAR\r\n" +
      "VERSION:2.0\r\n" +
      "PRODID:-//Google Inc//Google Apps Script//EN\r\n" +
      "METHOD:REQUEST\r\n" + 
      "BEGIN:VEVENT\r\n" +
      "UID:NGTL-" + Utilities.getUuid() + "\r\n" +
      "ORGANIZER;CN=" + ngtlsenderName + ":MAILTO:" + iCalOrganizerEmail + "\r\n";
      
    // Doklejamy uczestników z kohorty do pliku ICS, aby Outlook/Gmail zapisał ich na wydarzeniu
    cohortEmails.forEach(email => {
      let participantEmail = IS_TEST_MODE ? mailPgM : email.trim();
      eventsIcs += "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=" + participantEmail + ":MAILTO:" + participantEmail + "\r\n";
    });

    eventsIcs +=
      "DTSTART;TZID=" + iCalTimezone + ":" + iCalStartStr + "\r\n" + 
      "DTEND;TZID=" + iCalTimezone + ":" + iCalEndStr + "\r\n" +
      "SUMMARY:" + iterationSubject + "\r\n" +
      "LOCATION:" + webexLink + "\r\n" +
      "DESCRIPTION:You're invited to an NGTL session:\\n\\n" + sessionName + " – presented by " + fullPresenterName + "\\n\\nWebex Link: " + webexLink + "\\r\n" +
      "SEQUENCE:0\r\n" +
      "STATUS:CONFIRMED\r\n" +
      "TRANSP:OPAQUE\r\n" + 
      "END:VEVENT\r\n" +
      "END:VCALENDAR";

    // 3. Dodanie banera testowego z listą maili nad treścią wiadomości, jeśli IS_TEST_MODE = true
    let testModeBanner = "";
    if (IS_TEST_MODE) {
      testModeBanner = `
        <div style="background-color: #fee; border: 1px solid #fcc; padding: 12px; margin-bottom: 20px; font-size: 12px; color: #b71c1c; border-radius: 4px;">
          <strong>[TEST MODE] Zamiast do grupy docelowej, mail poszedł na adres testowy.</strong><br>
          <span style="display:block; margin-top: 5px;">Oryginalni odbiorcy z kohorty <strong>${cohortName}</strong> (${cohortEmails.length} osób):</span>
          <span style="font-family: monospace;">${cohortEmails.join(", ")}</span>
        </div>`;
    }

    const htmlBody = `
      <html>
        <body style="line-height: 1.5; padding: 0 20px; font-family: Arial, sans-serif; color: #333;">
          ${testModeBanner}
          <p>You're invited to an NGTL session:</p>
          <p><strong>${sessionName}</strong> – presented by <strong>${fullPresenterName}</strong></p>
          <img src="cid:ngtlLogo" alt="NGTL Logo" style="max-width: 200px; margin-bottom: 20px;" />
        </body>
      </html>
    `;

    const ccrecipients = `${ngtlemail}, ${mailArchitect}, ${mailPgM}`;

    // Ustalanie odbiorców
    const recipientList = IS_TEST_MODE ? TEST_EMAIL : cohortEmails.join(',');
    const ccList = IS_TEST_MODE ? TEST_CC_EMAIL : ccrecipients; 
    const finalSubject = IS_TEST_MODE ? `[TEST CALENDAR] ${iterationSubject}` : iterationSubject;

    const icsAttachment = Utilities.newBlob(eventsIcs, 'text/calendar; method=REQUEST; charset=UTF-8', 'invite.ics');

    // Wysyłamy, tylko jeśli kohorta ma przynajmniej jednego zatwierdzonego uczestnika (lub test)
    if (cohortEmails.length > 0 || IS_TEST_MODE) {
      MailApp.sendEmail({
        name: ngtlsenderName,
        replyTo: ngtlemail,
        to: recipientList,
        cc: ccList,
        subject: finalSubject,
        htmlBody: htmlBody,
        inlineImages: {
          ngtlLogo: logoBlob
        },
        attachments: [icsAttachment]
      });
    }
  });

  // Zwracamy obiekt metadanych do logera (appendEmailLog)
  return {
    recipient: `Cohort(s): ${cohortNames.join(", ")}`, 
    subject: subject,
    cohorts: cohortNames.join(", ")
  };
}