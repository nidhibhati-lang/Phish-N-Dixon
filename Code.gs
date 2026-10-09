function doGet(e) {
  const token = e && e.parameter ? String(e.parameter.token || "").trim() : "";
  if (!token) {
    return ContentService.createTextOutput("No tracking token provided.");
  }

  const worksheet = getCampaignWorksheet_();
  const values = worksheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return ContentService.createTextOutput("Tracking token not found.");
  }

  const columns = getColumns_(values[0]);
  const statusColumn = columns.status;
  const tokenColumn = columns.trackingToken;
  if (statusColumn < 0 || tokenColumn < 0) {
    throw new Error('The worksheet must contain "Status" and "Tracking Token" columns.');
  }

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    if (values[rowIndex][tokenColumn] !== token) continue;

    const statusCell = worksheet.getRange(rowIndex + 1, statusColumn + 1);
    if (statusCell.getDisplayValue().trim().toLowerCase() === "pending") {
      statusCell.setValue("Clicked");
    }
    return ContentService.createTextOutput("Recorded.");
  }

  return ContentService.createTextOutput("Tracking token not found.");
}

function sendPendingCampaign() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const properties = PropertiesService.getScriptProperties();
    const baseUrl = properties.getProperty("BASE_URL");
    if (!baseUrl || !/^https:\/\/[^/?#]+(?:\/[^?#]*)?$/.test(baseUrl)) {
      throw new Error(
        'Set the BASE_URL script property to your HTTPS landing-page URL, without query parameters or a fragment.'
      );
    }

    const worksheet = getCampaignWorksheet_();
    let values = worksheet.getDataRange().getDisplayValues();
    if (!values.length) {
      throw new Error("The tracking worksheet is empty.");
    }

    let columns = getColumns_(values[0]);
    if (columns.email < 0 || columns.status < 0) {
      throw new Error('The worksheet must contain an "Email" (or "Email ID") column and a "Status" column.');
    }

    if (columns.trackingToken < 0) {
      columns.trackingToken = worksheet.getLastColumn() + 1;
      worksheet.getRange(1, columns.trackingToken).setValue("Tracking Token");
    }
    if (columns.emailSentAt < 0) {
      columns.emailSentAt = worksheet.getLastColumn() + 1;
      worksheet.getRange(1, columns.emailSentAt).setValue("Email Sent At");
    }

    values = worksheet.getDataRange().getDisplayValues();
    const existingTokens = new Set();
    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const token = (values[rowIndex][columns.trackingToken] || "").trim();
      if (token && existingTokens.has(token)) {
        throw new Error("Duplicate tracking tokens exist in the worksheet; resolve them before sending.");
      }
      if (token) existingTokens.add(token);
    }

    const recipients = [];
    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const row = values[rowIndex];
      const email = (row[columns.email] || "").trim();
      const status = (row[columns.status] || "").trim().toLowerCase();
      const sentAt = (row[columns.emailSentAt] || "").trim();
      if (status !== "pending" || sentAt || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) continue;

      let token = (row[columns.trackingToken] || "").trim();
      if (!token) {
        token = createTrackingToken_();
        while (existingTokens.has(token)) token = createTrackingToken_();
        existingTokens.add(token);
        row[columns.trackingToken] = token;
      }
      recipients.push({ rowNumber: rowIndex + 1, email: email, token: token });
    }

    if (!recipients.length) {
      Logger.log("No unsent employees with Pending status and valid email addresses were found.");
      return;
    }

    if (MailApp.getRemainingDailyQuota() < recipients.length) {
      throw new Error(
        "Not enough remaining daily email quota for this campaign. No campaign emails were sent."
      );
    }

    const tokenValues = [];
    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      tokenValues.push([values[rowIndex][columns.trackingToken] || ""]);
    }
    worksheet.getRange(2, columns.trackingToken + 1, tokenValues.length, 1).setValues(tokenValues);

    let sentCount = 0;
    for (const recipient of recipients) {
      const trackingUrl = baseUrl + "?token=" + encodeURIComponent(recipient.token);
      MailApp.sendEmail({
        to: recipient.email,
        subject: "Mandatory IT Security Verification",
        body:
          "Dear User,\n\nPlease open this link for the internal security awareness exercise:\n" +
          trackingUrl +
          "\n\nRegards,\nIT Infrastructure Team",
        htmlBody:
          '<p>Dear User,</p>' +
          '<p>Please open this link for the internal security awareness exercise:</p>' +
          '<p><a href="' +
          escapeHtml_(trackingUrl) +
          '" style="background-color:#0d6efd;color:white;padding:10px 15px;text-decoration:none;border-radius:5px">Open the exercise</a></p>' +
          "<p>Regards,<br>IT Infrastructure Team</p>"
      });
      worksheet.getRange(recipient.rowNumber, columns.emailSentAt + 1).setValue(new Date());
      sentCount++;
      Logger.log("Campaign email sent to " + recipient.email);
    }

    Logger.log("Campaign completed. Emails sent: " + sentCount);
  } finally {
    lock.releaseLock();
  }
}

function getCampaignWorksheet_() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    throw new Error("Set the SPREADSHEET_ID script property before using the campaign.");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const worksheetName = properties.getProperty("WORKSHEET_NAME");
  const worksheet = worksheetName
    ? spreadsheet.getSheetByName(worksheetName)
    : spreadsheet.getSheets()[0];
  if (!worksheet) {
    throw new Error("The configured tracking worksheet does not exist.");
  }
  return worksheet;
}

function getColumns_(headerRow) {
  const headers = headerRow.map(value => String(value).trim().toLowerCase());
  return {
    email: findColumn_(headers, ["email", "email id", "email address"]),
    status: headers.indexOf("status"),
    trackingToken: headers.indexOf("tracking token"),
    emailSentAt: headers.indexOf("email sent at")
  };
}

function findColumn_(headers, acceptedNames) {
  for (const name of acceptedNames) {
    const index = headers.indexOf(name);
    if (index >= 0) return index;
  }
  return -1;
}

function createTrackingToken_() {
  return Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
}

function escapeHtml_(value) {
  return value.replace(/[&<>"']/g, character => {
    return {
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      '"': "&quot;",
      "'": "&#39;"
    }[character];
  });
}
