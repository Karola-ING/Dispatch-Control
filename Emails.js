// ================================
// MAIL SENDER APP
// ================================

function mailSenderApp(mail, subject, htmlBody) {

  const ccrecipients = `${ngtlemail}, ${mailArchitect}, ${mailPgM}`;

  MailApp.sendEmail({
    name: ngtlsenderName,
    replyTo: ngtlemail,
    to: IS_TEST_MODE ? TEST_EMAIL : mail, // Automatic recipient switch
    cc: IS_TEST_MODE ? TEST_CC_EMAIL : ccrecipients, // Automatic CC recipient switch
    subject: IS_TEST_MODE ? `[TEST] ${subject}` : subject, // Prefix in Test mode 
    htmlBody: htmlBody
  });
}

// ================================
// CALENDAR SENDER (NATIVE MIME VERSION)
// ================================
function mailSenderAppWithCalendar(mail, subject, htmlBody, icsContent) {

  const ccrecipients = `${ngtlemail}, ${mailArchitect}, ${mailPgM}`;

  const recipient = IS_TEST_MODE ? TEST_EMAIL : mail; 
  const cc = IS_TEST_MODE ? TEST_CC_EMAIL : ccrecipients; 
  const finalSubject = IS_TEST_MODE ? `[TEST CALENDAR] ${subject}` : subject; 

  const icsAttachment = Utilities.newBlob(icsContent, 'text/calendar; method=REQUEST; charset=UTF-8', 'invite.ics');

  MailApp.sendEmail({
    name: ngtlsenderName,
    replyTo: ngtlemail,
    to: recipient,
    cc: cc,
    subject: finalSubject,
    htmlBody: htmlBody,
    attachments: [icsAttachment]
  });
}


// ================================
// FLOW EXECUTIONS & HTML TEMPLATES
// ================================

// FLOW: NGTL INVITATION (in development)
function executeInvitationFlow(sessionRows) {
  const presenterEmail = sessionRows[0][COL.PRESENTER_EMAIL];
  const presenterFirstName = sessionRows[0][COL.PRESENTER].split(" ")[0];

  const cohortNames = [];
  
  const subject = "[Invitation] New NGTL Edition Placeholder";
  const htmlBody = `<p>Hello ${presenterFirstName},</p><p>This is a placeholder for Flow 1 (New Edition Invitation).</p>`;
  
  mailSenderApp(presenterEmail, subject, htmlBody)

  return {
    recipient: presenterEmail,
    subject: subject,
    cohorts: cohortNames.join(", ")
  };
}


// FLOW: SESSION CONFIRMATION
function executeConfirmationFlow(sessionRows) {
  const fullPresenterName = sessionRows[0][COL.PRESENTER];
  const presenterFirstName = getPresenterGreetingNames(fullPresenterName);
  const presenterEmail = sessionRows[0][COL.PRESENTER_EMAIL].split(",").map(email => email.trim()).join(",");
  const sessionName = sessionRows[0][COL.SESSION_NAME];
  const durationMinutes = sessionRows[0][COL.DURATION];
  
  const months = [];
  const cohortsDetailsList = [];
  const cohortNames = [];
  
  sessionRows.forEach(row => {
    const sessionDate = new Date(row[COL.SESSION_DATE]);
    const cohortName = row[COL.COHORT];
    const startTimeStr = row[COL.PRESENTER_TIME]; 
    const timezone = row[COL.PRESENTER_TIMEZONE]; 

    if (!cohortNames.includes(cohortName)) {
      cohortNames.push(cohortName);
    }
    
    // Subject month
    const monthName = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), "MMMM");
    if (!months.includes(monthName)) months.push(monthName);
    
    // Data format: Tue, Sep 15
    const formattedDate = Utilities.formatDate(sessionDate, Session.getScriptTimeZone(), "E, MMM d");
    const timeFrame = calculateEndTime(startTimeStr, durationMinutes);
    
    cohortsDetailsList.push(`<li><strong>${cohortName} cohort:</strong> ${formattedDate}; ${timeFrame}</li>`);
  });
  
  const subject = `[Action Needed] NGTL: Confirm your availability for NGTL sessions in ${months.join("/")}`;
  const htmlBody = getConfirmationHtml(presenterFirstName, sessionName, durationMinutes, cohortsDetailsList.join(""));

  const metadata = {
    recipient: presenterEmail,
    subject: subject,
    cohorts: cohortNames.join(", ")
  }
  
  mailSenderApp(presenterEmail, subject, htmlBody)

  return metadata
}


// FLOW: VPS SECLECTION
function executeVpSelectionFlow(vpRow, fullDepartmentTableRaw) {
  const currentOrg = vpRow[VP_COL.ORG];
  const fullVpName = vpRow[VP_COL.VP_NAME];
  const vpEmail = vpRow[VP_COL.VP_EMAIL];
  const spreadsheetLink = vpRow[VP_COL.SPREADSHEET_LINK];
  
  const totalSeats = vpRow[VP_COL.TOTAL_SEATS];
  const apjSeats = vpRow[VP_COL.APJ_SEATS];
  const emeaSeats = vpRow[VP_COL.EMEA_SEATS];
  const amerSeats = vpRow[VP_COL.AMER_SEATS];

  const hasWaitingList = vpRow[VP_COL.WAITING_LIST] === true;
  const waitingSeats = vpRow[VP_COL.WAITING_SEATS];

  // Parsing the name: we only remove the last word (last name) and keep the rest, e.g., “R. H.” or “John”
  const nameParts = fullVpName.trim().split(" ");
  const vpFirstName = nameParts.length > 1 ? nameParts.slice(0, -1).join(" ") : fullVpName;

  // Generating an HTML table with a benchmark
  const tableHtml = generateDepartmentTableHtml(fullDepartmentTableRaw, currentOrg);

  const subject = "[Action Needed] NGTL 2026 Participant Selection – Time for your final decision!";
  const htmlBody = getVpSelectionHtml(vpFirstName, totalSeats, apjSeats, emeaSeats, amerSeats, tableHtml, spreadsheetLink, hasWaitingList, waitingSeats);

  mailSenderApp(vpEmail, subject, htmlBody)

  return {
    recipient: vpEmail,
    subject: subject,
    cohorts: "N/A"
  };
}

// FLOW: VP REMINDER EXECUTION
function executeVpReminderFlow(vpRow, fullDepartmentTableRaw) {
  const currentOrg = vpRow[VP_COL.ORG];
  const fullVpName = vpRow[VP_COL.VP_NAME];
  const vpEmail = vpRow[VP_COL.VP_EMAIL];
  const spreadsheetLink = vpRow[VP_COL.SPREADSHEET_LINK];
  
  const totalSeats = vpRow[VP_COL.TOTAL_SEATS];
  const apjSeats = vpRow[VP_COL.APJ_SEATS];
  const emeaSeats = vpRow[VP_COL.EMEA_SEATS];
  const amerSeats = vpRow[VP_COL.AMER_SEATS];

  const hasWaitingList = vpRow[VP_COL.WAITING_LIST] === true;
  const waitingSeats = vpRow[VP_COL.WAITING_SEATS];

  // Wyodrębnienie imienia VP
  const nameParts = fullVpName.trim().split(" ");
  const vpFirstName = nameParts.length > 1 ? nameParts.slice(0, -1).join(" ") : fullVpName;

  // Wygenerowanie tabeli HTML z benchmarkiem
  const tableHtml = generateDepartmentTableHtml(fullDepartmentTableRaw, currentOrg);

  const subject = "[Action Needed] Reminder: NGTL 2026 Participant Selection (Due July 17th)";
  const htmlBody = getVpReminderHtml(vpFirstName, totalSeats, apjSeats, emeaSeats, amerSeats, tableHtml, spreadsheetLink, hasWaitingList, waitingSeats);

  mailSenderApp(vpEmail, subject, htmlBody);

  return {
    recipient: vpEmail,
    subject: subject,
    cohorts: "N/A"
  };
}


// Create Seat Allocation Table
function generateDepartmentTableHtml(departments, targetOrg) {
  let rowsHtml = "";
  
  departments.forEach(dept => {
    //  Use a distinctive background for VPs department
    const isTarget = dept.org === targetOrg;
    const rowBg = isTarget ? "#00A4EB" : ""; 
    const fontStyle = isTarget ? "font-weight: bold; color: #F5F5F5;" : "";

    rowsHtml += `
      <tr style="background-color: ${rowBg}; ${fontStyle}">
        <td style="padding: 8px; border-bottom: 1px solid #DEDEDE;">${dept.org}</td>
        <td style="padding: 8px; border-bottom: 1px solid #DEDEDE; text-align: center;">${dept.amer}</td>
        <td style="padding: 8px; border-bottom: 1px solid #DEDEDE; text-align: center;">${dept.apj}</td>
        <td style="padding: 8px; border-bottom: 1px solid #DEDEDE; text-align: center;">${dept.emea}</td>
        <td style="padding: 8px; border-bottom: 1px solid #DEDEDE; text-align: center;">${dept.total}</td>
      </tr>
    `;
  });

  const totalRowHtml = `
    <tr style="font-weight: bold; background-color: #F5F5F5; border-top: 2px solid #818181;">
      <td style="padding: 10px;">Cohort Total</td>
      <td style="padding: 10px;text-align: center;">24</td>
      <td style="padding: 10px;text-align: center;">24</td>
      <td style="padding: 10px;text-align: center;">12</td>
      <td style="padding: 10px;text-align: center;">60</td>
    </tr>
  `;

  return `
    <table style="border-collapse: collapse; width: 100%; max-width: 600px;">
      <thead>
        <tr style="font-weight: bold; background: #DEDEDE; font-size: 1.2rem;">
          <th style="padding: 10px; border-bottom: 2px solid #818181; text-align: left;">Department / Org</th>
          <th style="padding: 10px; border-bottom: 2px solid #818181; ">AMER</th>
          <th style="padding: 10px; border-bottom: 2px solid #818181; ">APJ</th>
          <th style="padding: 10px; border-bottom: 2px solid #818181; ">EMEA</th>
          <th style="padding: 10px; border-bottom: 2px solid #818181; ">Allocated Total</th>
        </tr>
      </thead>
      <tbody>
        ${rowsHtml}
        ${totalRowHtml}
      </tbody>
    </table>
  `;
}


// FLOW: OUTLOOK CALENDAR PLACEHOLDER (STREFE PREZENTERA & POPRAWNE RSVP)
function executeCalendarPlaceholderFlow(sessionRows) {
  const fullPresenterName = sessionRows[0][COL.PRESENTER];
  const presenterEmail = sessionRows[0][COL.PRESENTER_EMAIL];
  const sessionName = sessionRows[0][COL.SESSION_NAME];
  const sessionNo = sessionRows[0][COL.NO];
  const cohort = sessionRows[0][COL.COHORT];

  const iCalAttendeeEmail = IS_TEST_MODE ? mailPgM : presenterEmail;
  const iCalAttendeeName = IS_TEST_MODE ? "Test User" : fullPresenterName.replace(/"/g, "");

  const activeUserEmail = Session.getActiveUser().getEmail(); 
  const iCalOrganizerEmail = IS_TEST_MODE ? "ngtl-system@akamai.com" : ngtlemail;

  // ✅ ROZWIĄZANIE: Definiujemy subject na poziomie całej funkcji, żeby return na dole go widział
  const subject = `NGTL ${sessionNo}: ${sessionName} [PLACEHOLDER]`;

  const cohortNames = [];
  let eventsIcs = "";
  let emailListHtml = ""; 

  sessionRows.forEach(row => {
    const cohortName = row[COL.COHORT];
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
    emailListHtml += `<li><strong>${cohortName} cohort:</strong> ${formattedDateText} at ${presenterTimeStr}</li>`;

    // Dla każdego wysyłanego maila generujemy dedykowany temat kohortowy
    const iterationSubject = `NGTL ${sessionNo}: ${sessionName} (${cohortName}) [PLACEHOLDER]`;

    eventsIcs = 
      "BEGIN:VCALENDAR\r\n" +
      "VERSION:2.0\r\n" +
      "PRODID:-//Google Inc//Google Apps Script//EN\r\n" +
      "METHOD:REQUEST\r\n" + 
      "BEGIN:VEVENT\r\n" +
      "UID:NGTL-" + Utilities.getUuid() + "\r\n" +
      "ORGANIZER;CN=" + ngtlsenderName + ":MAILTO:" + iCalOrganizerEmail + "\r\n" +
      "ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=" + iCalAttendeeName + ":MAILTO:" + iCalAttendeeEmail + "\r\n" +
      "DTSTART;TZID=" + iCalTimezone + ":" + iCalStartStr + "\r\n" + 
      "DTEND;TZID=" + iCalTimezone + ":" + iCalEndStr + "\r\n" +
      "SUMMARY:NGTL 2026: " + sessionName + " (" + cohortName + ")\r\n" +
      "DESCRIPTION:Dear " + iCalAttendeeName + ",\\n\\nThis is a calendar placeholder.\\n\\nSession: " + sessionName + "\\nCohort: " + cohortName + "\\r\n" +
      "SEQUENCE:0\r\n" +
      "STATUS:CONFIRMED\r\n" +
      "TRANSP:OPAQUE\r\n" + 
      "END:VEVENT\r\n" +
      "END:VCALENDAR";

    // Budujemy dedykowaną treść HTML
    const htmlBody = `
      <html>
        <body style="line-height: 1.5; padding: 0 20px;">
          <p>Hello ${iCalAttendeeName.split(" ")[0]},</p>
          <p>This is an automated calendar placeholder for your NGTL session: <strong>${sessionName}</strong>.</p>
          <p>The time has been blocked in your calendar for the following cohort session:</p>
          <ul style="margin: 15px 0; padding-left: 20px;">
            <li><strong>${cohortName} cohort:</strong> ${formattedDateText} at ${presenterTimeStr}</li>
          </ul>
          <p>Please accept the attached invitation to add this event to your Outlook schedule. No further action is required from your side.</p>
          <p>Thank you,<br/><strong>The NGTL Team</strong></p>
        </body>
      </html>
    `;

    // Używamy tematu dedykowanego pod konkretną pętlę
    mailSenderAppWithCalendar(presenterEmail, iterationSubject, htmlBody, eventsIcs);
  });

  // ✅ Zwracamy ogólny, widoczny już dla funkcji "subject" do logów
  return {
    recipient: presenterEmail,
    subject: subject,
    cohorts: cohortNames.join(", ")
  };
}


// ================================
// HTML BODY GENERATORS
// ================================

// HTML BODY: SESSION CONFIRMATION
function getConfirmationHtml(name, sessionName, duration, cohortsLinesHtml) {
  return `
    <html>
      <body style="line-height: 1.5; padding: 0 20px;">
        <p>Hello ${name},</p>
        <p>I hope you're having a great week!</p>
        <p>I’m reaching out to check your availability for your upcoming NGTL session: <strong>${sessionName}</strong>.</p>
        <p>Could you please let me know if the following ${duration}-minute slot(s) work for you?</p>
        
        <ul style="list-style-type: disc; margin: 20px 0 20px 25px; padding: 0;;">
          ${cohortsLinesHtml}
        </ul>
        
        <p>If none of these times work, no worries at all – just let me know what suits you better and we’ll find an alternative. Once we have the dates locked in, I’ll send over the calendar placeholders.</p>
        <p>Thanks so much for your help with this!</p>
        <p>Thank you,<br/>
        <strong>The NGTL Team</strong>
        </p>
      </body>
    </html>
  `;
}


// HTML BODY: VP SELECTION
function getVpSelectionHtml(vpName, total, apj, emea, amer, tableHtml, link, hasWaitingList, waitingSeats) {
  const regionalLines = [];

  function isValueActive(val) {
    if (!val || val === "-" || val === "0" || val === 0) return false;
    return true;
  }

  if (isValueActive(apj)) regionalLines.push(`<li>APJ: <strong>${apj}</strong></li>`);
  if (isValueActive(emea)) regionalLines.push(`<li>EMEA: <strong>${emea}</strong></li>`);
  if (isValueActive(amer)) regionalLines.push(`<li>AMER: <strong>${amer}</strong></li>`);

  const regionalTargetsHtml = regionalLines.join("\n");

  // Conditional HTML block: appears only for those with waiting list enabled
  const waitingListParagraph = hasWaitingList 
    ? `<p>Additionally, this year you have the option to select up to <strong>${waitingSeats} backup candidates</strong> for the waiting list, in case any of the primary nominees are unable to participate. Please use comments column to share any notes or feedback that you think would be helpful for the NGTL Team.</p>`
    : '';

  if (vpName === "Sareen") {
    return `
    <html>
      <body style="line-height: 1.5; padding: 0 20px;">
        <p>Hello ${vpName},</p>
        
        <p>Nominations for NGTL 2026 are officially closed!</p>
        <p>Now comes the challenging but exciting part: deciding who will secure a spot in this year's cohort. We are looking for our future trail-blazers—high-performing individuals who have already demonstrated exceptional dedication and impact as individual contributors.</p>
        
        <p>For the <strong>Customer Success</strong> department, only one candidate has been nominated: <strong>Adrienne Hatlevig</strong>. Here are her details for your review:</p>
        <ul>
          <li><strong>Geo: </strong>Americas</li>
          <li><strong>Job title: </strong>Senior Customer Success Manager</li>
          <li><strong>Tenure: </strong> 6.5 years</li>
        </ul>
        
        <p>Helena nominated her with the following words:</p>
        <p style="margin: 0 1.5em; font-style: italic;">Adrienne Hatlevig has established herself as a transformative force within the Customer Success organization, earning the title of Commercial MVP for Q1 2026 and currently ranking number one in pipeline opportunity with a staggering $232,000 in MRR across her 15-account territory. Adrienne’s impact is defined by her role as a trailblazer in value realization; she was the first to develop and share success plans using our new frameworks and innovated the ""Offload Report Gem,"" a critical tool now used broadly by the team to demonstrate tangible value to stakeholders. Her thought leadership was further highlighted when she presented her success planning methodology at a global session, setting a new standard for the organization.</p>
        <p style="margin: 0 1.5em; font-style: italic;">Beyond her framework contributions, Adrienne is a master of turning technical friction into commercial growth- she has done a great job on her aligned Commercial and CTM accounts over the years. She successfully supported the closure of a $34,000 GMRR upgrade for Mazda, moving them from a WAF-only posture to a robust suite including Bot Manager Premier and API Security. Her ""Adoption Strategy""—which leverages data-backed reporting on consumption gaps—has been instrumental in overcoming PS pushback at accounts like RealPage and Vivid Seats, transforming skeptics into high-value advocates. Whether she is leading high-stakes onsite reviews for Corsair, managing intense security crises at Michael’s, or onboarding complex APR/BMP solutions, Adrienne consistently demonstrates the perfect blend of technical expertise and commercial instinct. She created product books so customers can have a holistic understanding of our offerings around Compute, Performance and Security. She has mentored several Customer Success Managers and innovated in the kinds of customer reports she creates i.e. Mid-Year reviews continues to raise the bar in enabling the field through providing customer success trainings around success planning and relationship management best practices. It's been incredible to see her growth at Akamai and she makes the perfect candidate for NGTL!"</p>

        <p>We look forward to your decision on whether you would like to accept Adrienne's nomination and secure her spot in this year's NGTL cohort.</p>

        <p>If not, below is the breakdown of the seat allocations across other teams and regions for your reference:</p>
        
        <div style="margin: 20px 0;">
          ${tableHtml}
        </div>
        
        <p><strong>Please let us know your decision by July 17th.</strong></p>
        
        <p>If you have any questions or need a sounding board while making your decisions, please don't hesitate to reach out. We’re here to help!</p>
        
        <p>Thank you,<br/>
        <strong>The NGTL Team</strong></p>
      </body>
    </html>
  `;
  }
  else {
    return `
    <html>
      <body style="line-height: 1.5; padding: 0 20px;">
        <p>Hello ${vpName},</p>
        
        <p>Nominations for NGTL 2026 are officially closed!</p>
        <p>Now comes the challenging but exciting part: deciding who will secure a spot in this year's cohort. We are looking for our future trail-blazers—high-performing individuals who have already demonstrated exceptional dedication and impact as individual contributors.</p>
        
        <p>Your leadership team has put forward fantastic nominees. From this list, please select <strong>${total} people</strong>, while keeping our regional targets in mind:</p>
        <ul>
          ${regionalTargetsHtml}
        </ul>
        
        <p>For your reference, here is the breakdown of the nominations across teams and regions:</p>
        
        <div style="margin: 20px 0;">
          ${tableHtml}
        </div>
        
        <p><strong>Next Steps:</strong></p>
        <p>Please head over to the <a href="${link}" style="font-weight: bold;">NGTL 2026 Participant Selection</a> sheet and update the statuses in column A to allocate all of your available seats based on the guidelines above.</p>
        ${waitingListParagraph}
        
        <p><strong>Please complete your selection by July 17th.</strong></p>
        
        <p>If you have any questions or need a sounding board while making your decisions, please don't hesitate to reach out. We’re here to help!</p>
        
        <p>Thank you,<br/>
        <strong>The NGTL Team</strong></p>
      </body>
    </html>
  `;
  }  
}

// HTML BODY: VP REMINDER
function getVpReminderHtml(vpName, total, apj, emea, amer, tableHtml, link, hasWaitingList, waitingSeats) {
  const regionalLines = [];

  function isValueActive(val) {
    if (!val || val === "-" || val === "0" || val === 0) return false;
    return true;
  }

  if (isValueActive(apj)) regionalLines.push(`<li>APJ: <strong>${apj}</strong></li>`);
  if (isValueActive(emea)) regionalLines.push(`<li>EMEA: <strong>${emea}</strong></li>`);
  if (isValueActive(amer)) regionalLines.push(`<li>AMER: <strong>${amer}</strong></li>`);

  const regionalTargetsHtml = regionalLines.join("\n");

  const waitingListParagraph = hasWaitingList 
    ? `<p>You can also select up to <strong>${waitingSeats} backup candidates</strong> for the waiting list and leave any notes in the comments column that you think would be helpful for the NGTL Team.</p>`
    : '';

  return `
    <html>
      <body style="line-height: 1.5; padding: 0 20px;">
        <p>Hello ${vpName},</p>

        <p>Hope you’re having a great week!</p>
        
        <p>Just a gentle reminder to finalize your participant selection for the <strong>NGTL 2026</strong> cohort by <strong>this Friday, July 17th</strong>.</p>
        
        <p>Please select <strong>${total} people</strong> from your nominee list, keeping your regional targets in mind:</p>
        <ul>
          ${regionalTargetsHtml}
        </ul>
        
        <p>For your reference, here is the breakdown of the nominations across teams and regions:</p>
        
        <div style="margin: 20px 0;">
          ${tableHtml}
        </div>
        
        <p><strong>Next Steps:</strong></p>
        <p>Go to the <a href="${link}" style="font-weight: bold;">NGTL 2026 Participant Selection</a> sheet and update the statuses in Column A to allocate your seats.</p>
        ${waitingListParagraph}
                
        <p>If you have any questions or need a sounding board, please let us know. We’re here to help!</p>
        
        <p>Thank you,<br/>
        <strong>The NGTL Team</strong></p>
      </body>
    </html>
  `;
}