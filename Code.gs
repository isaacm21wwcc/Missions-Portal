// ═══════════════════════════════════════════════════════════
//  GC YOUTH MISSIONS TRIP — APPS SCRIPT BACKEND
//  Code.gs  —  Auth, email, Sheet read/write
//
//  SETUP:
//  1. Replace SHEET_ID with your Google Sheet ID
//  2. Replace ADMIN_PASSWORD with your chosen password
//  3. Replace NOREPLY_NAME with how emails should appear
//  4. Replace YOUTH_COORDINATOR_EMAIL with your fallback address
//  5. Run initSheets() once to create Sheet tabs
//  6. Deploy as Web App (Execute as: Me, Access: Anyone)
// ═══════════════════════════════════════════════════════════

const SHEET_ID               = 'https://docs.google.com/spreadsheets/d/10YQHHT9ON-kFfBtQNUn532-FfKJHy1DnE4_yiz93Z4M/edit?usp=sharing';
const ADMIN_PASSWORD         = 'gcyouth2027';
const NOREPLY_NAME           = 'GC Youth Missions';
const YOUTH_COORDINATOR_EMAIL = 'isaac.mitchell@generocitychurch.com';   // fallback if no leader assigned
const PORTAL_URL             = 'https://missionsportal.isaacmitchell0803.workers.dev/'; // e.g. https://gcyouth.netlify.app

const TRIP_NAME    = 'GC Youth × Horizon Project';
const TRIP_DATE    = 'September 2027';
const TRIP_TOTAL   = 3450;
const DEPOSIT_AMT  = 500;
const FINAL_CUTOFF = '1 Jul 2027';

const TAB = {
  participants: 'Participants',
  passwords:    'Passwords',
  payments:     'Payments',
  profile:      'ProfileData',
  messages:     'Messages',
};


// ═══════════════════════════════════════════════════════════
//  ROUTING
// ═══════════════════════════════════════════════════════════

// Handle CORS preflight requests
function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doGet(e) {
  const page = (e.parameter && e.parameter.page) || '';
  const fn   = (e.parameter && e.parameter.fn)   || '';

  // ── API call from Netlify frontend ──
  if (fn) {
    const callback = (e.parameter && e.parameter.callback) || '';
    try {
      const args   = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      const result = callFunction(fn, args);
      const json   = JSON.stringify(result);
      // JSONP response if callback provided, plain JSON otherwise
      const body   = callback ? callback + '(' + json + ')' : json;
      return ContentService
        .createTextOutput(body)
        .setMimeType(callback
          ? ContentService.MimeType.JAVASCRIPT
          : ContentService.MimeType.JSON);
    } catch(err) {
      const json = JSON.stringify({ error: err.message });
      const body = callback ? callback + '(' + json + ')' : json;
      return ContentService
        .createTextOutput(body)
        .setMimeType(callback
          ? ContentService.MimeType.JAVASCRIPT
          : ContentService.MimeType.JSON);
    }
  }

  // ── Default response for direct browser access ──
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'GC Youth Missions API is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Routes function name string to the actual function
function callFunction(name, args) {
  const allowed = [
    'checkEmailRegistered','createAccount','studentLogin',
    'getStudentData','getTripConfig',
    'submitProfileSection','requestSectionEdit',
    'makePayment',
    'adminLogin','getAllParticipants','registerParticipant',
    'sendInviteEmail','sendReminderEmail',
    'approveProfileSection','rejectProfileSection','getPendingProfileChanges',
    'recordPaymentAdmin','saveAdminNotes',
  ];
  if (!allowed.includes(name)) throw new Error('Function not allowed: ' + name);
  const fns = {
    checkEmailRegistered,createAccount,studentLogin,
    getStudentData,getTripConfig,
    submitProfileSection,requestSectionEdit,
    makePayment,
    adminLogin,getAllParticipants,registerParticipant,
    sendInviteEmail,sendReminderEmail,
    approveProfileSection,rejectProfileSection,getPendingProfileChanges,
    recordPaymentAdmin,saveAdminNotes,
  };
  return fns[name](...args);
}


// ═══════════════════════════════════════════════════════════
//  SHEET HELPERS
// ═══════════════════════════════════════════════════════════

function getSheet(tabName) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(tabName);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function findRowByCol(sheet, colHeader, value) {
  const data  = sheet.getDataRange().getValues();
  const headers = data[0];
  const col   = headers.indexOf(colHeader);
  if (col < 0) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col]).toLowerCase() === String(value).toLowerCase()) return i + 1;
  }
  return -1;
}

function getColIndex(sheet, header) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(header) + 1; // 1-based
}

function setCell(sheet, rowHeader, colHeader, value) {
  const row = findRowByCol(sheet, 'ID', rowHeader);
  if (row < 0) return false;
  const col = getColIndex(sheet, colHeader);
  if (col < 1) return false;
  sheet.getRange(row, col).setValue(value);
  return true;
}


// ═══════════════════════════════════════════════════════════
//  SHEET INITIALISATION  —  run once manually
// ═══════════════════════════════════════════════════════════

function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const tabs = {
    [TAB.participants]: [
      'ID','FirstName','LastName','Email','Mobile','DOB','Gender','Address',
      'YouthContactName','YouthContactEmail',
      'InviteSent','AccountActive','ProfileComplete','DepositPaid',
      'TotalPaid','Status','CreatedAt','Notes'
    ],
    [TAB.passwords]: [
      'Email','PasswordHash','ParticipantID','CreatedAt','LastLogin'
    ],
    [TAB.payments]: [
      'ParticipantID','Description','Amount','Method','Reference','Date','AddedBy'
    ],
    [TAB.profile]: [
      'ParticipantID','Section','Field','Value','Status','SubmittedAt','ApprovedAt'
    ],
    [TAB.messages]: [
      'ParticipantID','Title','Body','Type','Date','Read'
    ],
  };

  Object.entries(tabs).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a4fa0')
      .setFontColor('#ffffff');
  });

  return { success: true, message: 'All sheets initialised.' };
}


// ═══════════════════════════════════════════════════════════
//  ADMIN AUTH
// ═══════════════════════════════════════════════════════════

function adminLogin(password) {
  return password === ADMIN_PASSWORD
    ? { success: true }
    : { success: false, error: 'Incorrect password.' };
}


// ═══════════════════════════════════════════════════════════
//  STUDENT AUTH
// ═══════════════════════════════════════════════════════════

// Called from the invite link — checks email is registered before showing setup screen
function checkEmailRegistered(email) {
  try {
    const ps   = getSheet(TAB.participants);
    const rows = sheetToObjects(ps);
    const p    = rows.find(r => r.Email && r.Email.toLowerCase() === email.toLowerCase());

    if (!p) return { registered: false, error: 'This email address is not registered for the trip. Please contact your youth coordinator.' };
    if (p.AccountActive === 'Yes') return { registered: true, alreadyActive: true, firstName: p.FirstName };

    return {
      registered:    true,
      alreadyActive: false,
      firstName:     p.FirstName,
      lastName:      p.LastName,
      tripName:      TRIP_NAME,
      tripDate:      TRIP_DATE,
      tripTotal:     TRIP_TOTAL,
      depositAmt:    DEPOSIT_AMT,
      finalCutoff:   FINAL_CUTOFF,
    };
  } catch(e) {
    return { error: e.message };
  }
}

// Called when student submits the create-account form
function createAccount(email, password) {
  try {
    const ps   = getSheet(TAB.participants);
    const pws  = getSheet(TAB.passwords);
    const rows = sheetToObjects(ps);
    const p    = rows.find(r => r.Email && r.Email.toLowerCase() === email.toLowerCase());

    if (!p) return { success: false, error: 'Email not registered. Please contact your youth coordinator.' };
    if (p.AccountActive === 'Yes') return { success: false, error: 'An account already exists for this email. Please sign in instead.' };

    // Store hashed password
    const hash = Utilities.base64Encode(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      email.toLowerCase() + '::' + password
    ));

    const now = new Date().toLocaleDateString('en-AU');
    pws.appendRow([email.toLowerCase(), hash, p.ID, now, '']);

    // Mark account active
    const row = findRowByCol(ps, 'Email', email);
    ps.getRange(row, getColIndex(ps, 'AccountActive')).setValue('Yes');

    // Send welcome email to student
    sendWelcomeEmail(p);

    // Notify youth contact
    notifyYouthContact(p);

    // Add welcome message in portal
    addMessage(p.ID,
      'Welcome to your participant portal!',
      'Hi ' + p.FirstName + '! Your account is active. Please complete your profile and pay your $' + DEPOSIT_AMT + ' deposit to confirm your place on the trip. We\'re so excited you\'re joining us!',
      'welcome'
    );

    return {
      success:       true,
      participantId: p.ID,
      firstName:     p.FirstName,
      lastName:      p.LastName,
      email:         p.Email,
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Called on every subsequent login
function studentLogin(email, password) {
  try {
    const pws  = getSheet(TAB.passwords);
    const ps   = getSheet(TAB.participants);
    const rows = sheetToObjects(pws);

    const hash = Utilities.base64Encode(Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      email.toLowerCase() + '::' + password
    ));

    const record = rows.find(r =>
      r.Email && r.Email.toLowerCase() === email.toLowerCase() &&
      r.PasswordHash === hash
    );

    if (!record) return { success: false, error: 'Incorrect email or password.' };

    // Update last login timestamp
    const row = findRowByCol(pws, 'Email', email);
    pws.getRange(row, getColIndex(pws, 'LastLogin')).setValue(new Date().toLocaleDateString('en-AU'));

    // Get participant details
    const parts = sheetToObjects(ps);
    const p     = parts.find(r => String(r.ID) === String(record.ParticipantID));
    if (!p) return { success: false, error: 'Participant record not found.' };

    return {
      success:       true,
      participantId: p.ID,
      firstName:     p.FirstName,
      lastName:      p.LastName,
      email:         p.Email,
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}


// ═══════════════════════════════════════════════════════════
//  EMAIL SENDING
// ═══════════════════════════════════════════════════════════

function sendWelcomeEmail(participant) {
  try {
    const subject = 'Your GC Youth missions trip account is ready!';
    const name    = participant.FirstName;
    const portalUrl = PORTAL_URL;

    const html = `
      <div style="font-family:'DM Sans',sans-serif;max-width:560px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">
        <div style="background:#1a4fa0;padding:28px 32px;">
          <div style="font-size:22px;font-weight:700;color:#ffffff;">✈️ ${TRIP_NAME}</div>
          <div style="font-size:14px;color:#bfdbfe;margin-top:4px;">${TRIP_DATE}</div>
        </div>
        <div style="padding:32px;">
          <div style="font-size:18px;font-weight:600;color:#111827;margin-bottom:8px;">Hi ${name}, your account is active!</div>
          <p style="font-size:14px;color:#6b7280;line-height:1.6;margin-bottom:20px;">
            Your participant portal for the ${TRIP_NAME} missions trip is ready. Log in to complete your profile, view your payment schedule, and pay your deposit to confirm your place.
          </p>
          <div style="background:#f3f4f6;border-radius:10px;padding:16px 20px;margin-bottom:24px;font-size:13px;color:#374151;">
            <div style="display:flex;justify-content:space-between;padding:4px 0;"><span>Trip total</span><strong>$${TRIP_TOTAL.toLocaleString()}</strong></div>
            <div style="display:flex;justify-content:space-between;padding:4px 0;"><span>Deposit required</span><strong style="color:#b91c1c;">$${DEPOSIT_AMT}</strong></div>
            <div style="display:flex;justify-content:space-between;padding:4px 0;"><span>Final payment cutoff</span><strong>${FINAL_CUTOFF}</strong></div>
          </div>
          <a href="${portalUrl}" style="display:inline-block;background:#1a4fa0;color:#ffffff;font-weight:600;font-size:14px;padding:12px 28px;border-radius:8px;text-decoration:none;">
            Go to my portal →
          </a>
          <p style="font-size:12px;color:#9ca3af;margin-top:24px;line-height:1.6;">
            If you have any questions, contact your youth coordinator. This email was sent from a no-reply address — please don't reply directly.
          </p>
        </div>
      </div>`;

    GmailApp.sendEmail(participant.Email, subject, '', {
      htmlBody: html,
      name: NOREPLY_NAME,
    });

    return true;
  } catch(e) {
    Logger.log('Welcome email error: ' + e.message);
    return false;
  }
}

function notifyYouthContact(participant) {
  try {
    const contactEmail = participant.YouthContactEmail || YOUTH_COORDINATOR_EMAIL;
    const contactName  = participant.YouthContactName  || 'Youth Coordinator';
    const studentName  = participant.FirstName + ' ' + participant.LastName;

    const subject = studentName + ' has activated their missions trip portal';
    const html = `
      <div style="font-family:'DM Sans',sans-serif;max-width:560px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">
        <div style="background:#1a4fa0;padding:28px 32px;">
          <div style="font-size:20px;font-weight:700;color:#ffffff;">GC Youth Admin Notification</div>
          <div style="font-size:14px;color:#bfdbfe;margin-top:4px;">${TRIP_NAME} — ${TRIP_DATE}</div>
        </div>
        <div style="padding:32px;">
          <div style="font-size:16px;font-weight:600;color:#111827;margin-bottom:8px;">Hi ${contactName},</div>
          <p style="font-size:14px;color:#6b7280;line-height:1.6;margin-bottom:20px;">
            <strong style="color:#111827;">${studentName}</strong> has just activated their participant portal account for the ${TRIP_NAME} missions trip. They've been prompted to complete their profile and pay their deposit.
          </p>
          <div style="background:#e6f4ea;border-radius:10px;padding:16px 20px;margin-bottom:24px;font-size:13px;color:#2d7a3a;border:1px solid #86efac;">
            <strong>✓ Account activated</strong> — ${studentName} can now log in and manage their profile and payments.
          </div>
          <p style="font-size:13px;color:#6b7280;line-height:1.6;margin-bottom:20px;">
            You can view their profile and track their progress in the admin portal.
          </p>
          <a href="${PORTAL_URL}?page=admin" style="display:inline-block;background:#1a4fa0;color:#ffffff;font-weight:600;font-size:14px;padding:12px 28px;border-radius:8px;text-decoration:none;">
            Open admin portal →
          </a>
          <p style="font-size:12px;color:#9ca3af;margin-top:24px;">
            This is an automated notification. Do not reply to this email.
          </p>
        </div>
      </div>`;

    GmailApp.sendEmail(contactEmail, subject, '', {
      htmlBody: html,
      name: NOREPLY_NAME,
    });

    return true;
  } catch(e) {
    Logger.log('Youth contact notification error: ' + e.message);
    return false;
  }
}

// Send invite email (called from admin when registering a student)
function sendInviteEmail(participantId) {
  try {
    const ps   = getSheet(TAB.participants);
    const rows = sheetToObjects(ps);
    const p    = rows.find(r => String(r.ID) === String(participantId));
    if (!p) return { success: false, error: 'Participant not found.' };

    const subject = 'You\'re registered for the ' + TRIP_NAME + ' missions trip!';
    const html = `
      <div style="font-family:'DM Sans',sans-serif;max-width:560px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">
        <div style="background:#1a4fa0;padding:28px 32px;">
          <div style="font-size:22px;font-weight:700;color:#ffffff;">✈️ You're going on a missions trip!</div>
          <div style="font-size:14px;color:#bfdbfe;margin-top:4px;">${TRIP_NAME} — ${TRIP_DATE}</div>
        </div>
        <div style="padding:32px;">
          <div style="font-size:18px;font-weight:600;color:#111827;margin-bottom:8px;">Hi ${p.FirstName}!</div>
          <p style="font-size:14px;color:#6b7280;line-height:1.6;margin-bottom:20px;">
            You've been registered for the <strong style="color:#111827;">${TRIP_NAME}</strong> missions trip in ${TRIP_DATE}. Set up your participant account to complete your profile and manage your payments.
          </p>
          <div style="background:#f3f4f6;border-radius:10px;padding:16px 20px;margin-bottom:24px;font-size:13px;color:#374151;">
            <div style="display:flex;justify-content:space-between;padding:4px 0;"><span>Trip total</span><strong>$${TRIP_TOTAL.toLocaleString()}</strong></div>
            <div style="display:flex;justify-content:space-between;padding:4px 0;"><span>Deposit required</span><strong style="color:#b91c1c;">$${DEPOSIT_AMT} — due ASAP</strong></div>
            <div style="display:flex;justify-content:space-between;padding:4px 0;"><span>Your email</span><strong>${p.Email}</strong></div>
          </div>
          <a href="${PORTAL_URL}" style="display:inline-block;background:#1a4fa0;color:#ffffff;font-weight:600;font-size:14px;padding:14px 32px;border-radius:8px;text-decoration:none;font-size:15px;">
            Set up my account →
          </a>
          <p style="font-size:13px;color:#6b7280;margin-top:20px;line-height:1.6;">
            Use the email address above (<strong>${p.Email}</strong>) to create your account. If you have any questions, contact your youth coordinator${p.YouthContactName ? ' <strong>' + p.YouthContactName + '</strong>' : ''}.
          </p>
          <p style="font-size:12px;color:#9ca3af;margin-top:16px;">
            This email was sent from a no-reply address. Please don't reply directly.
          </p>
        </div>
      </div>`;

    GmailApp.sendEmail(p.Email, subject, '', {
      htmlBody: html,
      name: NOREPLY_NAME,
    });

    // Mark invite sent
    const row = findRowByCol(ps, 'ID', participantId);
    ps.getRange(row, getColIndex(ps, 'InviteSent')).setValue('Yes');

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Send a reminder email to a student
function sendReminderEmail(participantId, subject, body) {
  try {
    const ps   = getSheet(TAB.participants);
    const rows = sheetToObjects(ps);
    const p    = rows.find(r => String(r.ID) === String(participantId));
    if (!p) return { success: false, error: 'Participant not found.' };

    const html = `
      <div style="font-family:'DM Sans',sans-serif;max-width:560px;margin:0 auto;background:#ffffff;border-radius:12px;overflow:hidden;border:1px solid #e5e7eb;">
        <div style="background:#1a4fa0;padding:24px 32px;">
          <div style="font-size:18px;font-weight:700;color:#ffffff;">${TRIP_NAME}</div>
          <div style="font-size:13px;color:#bfdbfe;margin-top:2px;">${TRIP_DATE}</div>
        </div>
        <div style="padding:32px;">
          <p style="font-size:14px;color:#374151;line-height:1.7;">${body.replace(/\n/g,'<br>')}</p>
          <div style="margin-top:24px;">
            <a href="${PORTAL_URL}" style="display:inline-block;background:#1a4fa0;color:#ffffff;font-weight:600;font-size:14px;padding:11px 24px;border-radius:8px;text-decoration:none;">
              Open my portal →
            </a>
          </div>
          <p style="font-size:12px;color:#9ca3af;margin-top:20px;">This email was sent from a no-reply address.</p>
        </div>
      </div>`;

    GmailApp.sendEmail(p.Email, subject, '', { htmlBody: html, name: NOREPLY_NAME });
    addMessage(p.ID, subject, body, 'reminder');

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}


// ═══════════════════════════════════════════════════════════
//  PARTICIPANT DATA
// ═══════════════════════════════════════════════════════════

function getAllParticipants() {
  try {
    const ps   = getSheet(TAB.participants);
    const pmts = getSheet(TAB.payments);
    const parts    = sheetToObjects(ps);
    const payments = sheetToObjects(pmts);

    return parts.map(p => {
      const paid = payments
        .filter(pm => String(pm.ParticipantID) === String(p.ID))
        .reduce((sum, pm) => sum + Number(pm.Amount || 0), 0);
      return {
        id:               p.ID,
        name:             (p.FirstName + ' ' + p.LastName).trim(),
        firstName:        p.FirstName,
        lastName:         p.LastName,
        email:            p.Email,
        mobile:           p.Mobile,
        dob:              p.DOB,
        address:          p.Address,
        youthContactName: p.YouthContactName  || '',
        youthContactEmail:p.YouthContactEmail || '',
        inviteSent:       p.InviteSent    === 'Yes',
        accountActive:    p.AccountActive === 'Yes',
        profileComplete:  p.ProfileComplete === 'Yes',
        depositPaid:      paid >= DEPOSIT_AMT,
        totalPaid:        paid,
        remaining:        Math.max(0, TRIP_TOTAL - paid),
        status:           p.Status || 'pending',
        notes:            p.Notes || '',
        createdAt:        p.CreatedAt || '',
      };
    });
  } catch(e) {
    return { error: e.message };
  }
}

function registerParticipant(data) {
  try {
    const ps  = getSheet(TAB.participants);
    const id  = 'P' + Date.now();
    const now = new Date().toLocaleDateString('en-AU');

    ps.appendRow([
      id,
      data.firstName    || '',
      data.lastName     || '',
      data.email        || '',
      data.mobile       || '',
      data.dob          || '',
      data.gender       || '',
      data.address      || '',
      data.youthContactName  || '',
      data.youthContactEmail || '',
      'No',       // InviteSent
      'No',       // AccountActive
      'No',       // ProfileComplete
      'No',       // DepositPaid
      0,          // TotalPaid
      'pending',  // Status
      now,        // CreatedAt
      data.notes || '',
    ]);

    return { success: true, id: id };
  } catch(e) {
    return { error: e.message };
  }
}

function getStudentData(participantId) {
  try {
    const ps   = getSheet(TAB.participants);
    const pmts = getSheet(TAB.payments);
    const pd   = getSheet(TAB.profile);
    const msg  = getSheet(TAB.messages);

    const parts = sheetToObjects(ps);
    const p = parts.find(r => String(r.ID) === String(participantId));
    if (!p) return { error: 'Participant not found.' };

    const payments = sheetToObjects(pmts)
      .filter(r => String(r.ParticipantID) === String(participantId));
    const totalPaid = payments.reduce((sum, r) => sum + Number(r.Amount || 0), 0);

    const profileRows = sheetToObjects(pd)
      .filter(r => String(r.ParticipantID) === String(participantId));

    const sections = ['parent','emergency','medical','passport','church'];
    const sectionStatus = {};
    sections.forEach(s => {
      const rows = profileRows.filter(r => r.Section === s);
      if (!rows.length)                               sectionStatus[s] = 'incomplete';
      else if (rows.some(r => r.Status === 'approved')) sectionStatus[s] = 'approved';
      else if (rows.some(r => r.Status === 'pending'))  sectionStatus[s] = 'pending';
      else                                              sectionStatus[s] = 'incomplete';
    });

    const messages = sheetToObjects(msg)
      .filter(r => String(r.ParticipantID) === String(participantId))
      .sort((a, b) => new Date(b.Date) - new Date(a.Date));

    return {
      success:       true,
      id:            p.ID,
      firstName:     p.FirstName,
      lastName:      p.LastName,
      email:         p.Email,
      mobile:        p.Mobile,
      dob:           p.DOB,
      address:       p.Address,
      depositPaid:   totalPaid >= DEPOSIT_AMT,
      totalPaid:     totalPaid,
      remaining:     Math.max(0, TRIP_TOTAL - totalPaid),
      tripTotal:     TRIP_TOTAL,
      depositAmt:    DEPOSIT_AMT,
      finalCutoff:   FINAL_CUTOFF,
      tripName:      TRIP_NAME,
      tripDate:      TRIP_DATE,
      payments:      payments,
      sectionStatus: sectionStatus,
      messages:      messages,
    };
  } catch(e) {
    return { error: e.message };
  }
}

function getTripConfig() {
  return {
    tripName:    TRIP_NAME,
    tripDate:    TRIP_DATE,
    tripTotal:   TRIP_TOTAL,
    depositAmt:  DEPOSIT_AMT,
    finalCutoff: FINAL_CUTOFF,
  };
}


// ═══════════════════════════════════════════════════════════
//  PROFILE SECTIONS
// ═══════════════════════════════════════════════════════════

function submitProfileSection(participantId, section, fields) {
  try {
    const pd  = getSheet(TAB.profile);
    const now = new Date().toLocaleDateString('en-AU');
    const data = pd.getDataRange().getValues();
    const headers = data[0];
    const pidCol  = headers.indexOf('ParticipantID');
    const secCol  = headers.indexOf('Section');
    const statCol = headers.indexOf('Status');

    // Remove previous pending/rejected rows for this section
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][pidCol]) === String(participantId) &&
          data[i][secCol] === section &&
          ['pending','rejected'].includes(data[i][statCol])) {
        pd.deleteRow(i + 1);
      }
    }

    // Write each field
    Object.entries(fields).forEach(([field, value]) => {
      pd.appendRow([participantId, section, field, value, 'pending', now, '']);
    });

    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

function approveProfileSection(participantId, section) {
  try {
    const pd  = getSheet(TAB.profile);
    const now = new Date().toLocaleDateString('en-AU');
    const data = pd.getDataRange().getValues();
    const headers  = data[0];
    const pidCol   = headers.indexOf('ParticipantID');
    const secCol   = headers.indexOf('Section');
    const statCol  = headers.indexOf('Status');
    const appCol   = headers.indexOf('ApprovedAt');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][pidCol]) === String(participantId) &&
          data[i][secCol] === section &&
          data[i][statCol] === 'pending') {
        pd.getRange(i + 1, statCol + 1).setValue('approved');
        pd.getRange(i + 1, appCol   + 1).setValue(now);
      }
    }

    // Notify student
    addMessage(participantId,
      section.charAt(0).toUpperCase() + section.slice(1) + ' details approved',
      'Your ' + section + ' details have been reviewed and approved. They are now locked to your profile. To make changes, use the Request edit button.',
      'approval'
    );

    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

function rejectProfileSection(participantId, section, reason) {
  try {
    const pd  = getSheet(TAB.profile);
    const data = pd.getDataRange().getValues();
    const headers = data[0];
    const pidCol  = headers.indexOf('ParticipantID');
    const secCol  = headers.indexOf('Section');
    const statCol = headers.indexOf('Status');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][pidCol]) === String(participantId) &&
          data[i][secCol] === section &&
          data[i][statCol] === 'pending') {
        pd.getRange(i + 1, statCol + 1).setValue('rejected');
      }
    }

    addMessage(participantId,
      section + ' — please resubmit',
      'Your ' + section + ' details need to be updated before they can be approved. ' + (reason || 'Please check and resubmit.'),
      'rejection'
    );

    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}

function getPendingProfileChanges() {
  try {
    const pd   = getSheet(TAB.profile);
    const ps   = getSheet(TAB.participants);
    const rows = sheetToObjects(pd).filter(r => r.Status === 'pending');
    const parts = sheetToObjects(ps);

    const grouped = {};
    rows.forEach(r => {
      const key = r.ParticipantID + '::' + r.Section;
      if (!grouped[key]) {
        const p = parts.find(x => String(x.ID) === String(r.ParticipantID));
        grouped[key] = {
          participantId:   r.ParticipantID,
          participantName: p ? (p.FirstName + ' ' + p.LastName) : r.ParticipantID,
          section:         r.Section,
          submittedAt:     r.SubmittedAt,
          fields:          [],
        };
      }
      grouped[key].fields.push({ field: r.Field, value: r.Value });
    });

    return Object.values(grouped);
  } catch(e) {
    return { error: e.message };
  }
}

function requestSectionEdit(participantId, section, reason) {
  try {
    // Notify admin via a message (in production, also email admin)
    addMessage('ADMIN',
      'Edit request: ' + section,
      'Participant ' + participantId + ' requested to edit their ' + section + '. Reason: ' + (reason || 'Not provided.'),
      'edit-request'
    );
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}


// ═══════════════════════════════════════════════════════════
//  PAYMENTS
// ═══════════════════════════════════════════════════════════

function makePayment(participantId, amount, method) {
  try {
    const ref = addPaymentRow(participantId, 'Payment', amount, method, 'student');
    updateParticipantPaidTotal(participantId);
    const total = getParticipantTotalPaid(participantId);
    const today = new Date().toLocaleDateString('en-AU');
    addMessage(participantId,
      'Payment received — $' + Number(amount).toLocaleString(),
      'We received your payment of $' + Number(amount).toLocaleString() + ' on ' + today + '. Your total paid is now $' + total.toLocaleString() + '.',
      'payment'
    );
    return { success: true, reference: ref, totalPaid: total };
  } catch(e) {
    return { error: e.message };
  }
}

function recordPaymentAdmin(participantId, amount, description, method) {
  try {
    const ref = addPaymentRow(participantId, description || 'Payment', amount, method, 'admin');
    updateParticipantPaidTotal(participantId);
    return { success: true, reference: ref };
  } catch(e) {
    return { error: e.message };
  }
}

function addPaymentRow(participantId, description, amount, method, addedBy) {
  const pmts = getSheet(TAB.payments);
  const ref  = 'GCH-' + Math.floor(1000 + Math.random() * 9000);
  const now  = new Date().toLocaleDateString('en-AU');
  pmts.appendRow([participantId, description, Number(amount), method || 'card', ref, now, addedBy]);
  return ref;
}

function getParticipantTotalPaid(participantId) {
  const pmts = getSheet(TAB.payments);
  return sheetToObjects(pmts)
    .filter(r => String(r.ParticipantID) === String(participantId))
    .reduce((sum, r) => sum + Number(r.Amount || 0), 0);
}

function updateParticipantPaidTotal(participantId) {
  const ps    = getSheet(TAB.participants);
  const total = getParticipantTotalPaid(participantId);
  const row   = findRowByCol(ps, 'ID', participantId);
  if (row < 0) return;
  ps.getRange(row, getColIndex(ps, 'TotalPaid')).setValue(total);
  if (total >= DEPOSIT_AMT) {
    ps.getRange(row, getColIndex(ps, 'DepositPaid')).setValue('Yes');
    ps.getRange(row, getColIndex(ps, 'Status')).setValue('active');
  }
}

function saveAdminNotes(participantId, notes) {
  try {
    const ps  = getSheet(TAB.participants);
    const row = findRowByCol(ps, 'ID', participantId);
    if (row < 0) return { error: 'Participant not found.' };
    ps.getRange(row, getColIndex(ps, 'Notes')).setValue(notes);
    return { success: true };
  } catch(e) {
    return { error: e.message };
  }
}


// ═══════════════════════════════════════════════════════════
//  MESSAGES
// ═══════════════════════════════════════════════════════════

function addMessage(participantId, title, body, type) {
  const msg = getSheet(TAB.messages);
  const now = new Date().toLocaleDateString('en-AU');
  msg.appendRow([participantId, title, body, type || 'info', now, 'No']);
}
