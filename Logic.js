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