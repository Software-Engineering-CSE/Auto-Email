/**
 * Safely return sheet reference
 */
function getSheet_() {
  var sheetName = "Sheet1"; // 👈 update if renamed
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error("❌ Sheet '" + sheetName + "' not found.");
  return sheet;
}

/**
 * Fetch and clean recipient list
 */
function getRecipients_() {
  var sheet = getSheet_();
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) throw new Error("❌ No data rows found.");

  var headers = data[0];
  var requiredCols = ["CompanyName", "Email", "Address", "LOI Link", "Status", "LastSent"];
  requiredCols.forEach(function (col) {
    if (headers.indexOf(col) === -1) {
      throw new Error("❌ Missing required column: " + col);
    }
  });

  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var rowObj = {};
    headers.forEach(function (h, j) {
      rowObj[h] = data[i][j];
    });
    rowObj._rowIndex = i + 1;

    if (!rowObj.Email) {
      Logger.log("⚠️ Skipping row " + (i + 1) + " → Missing Email");
      continue;
    }
    if (!rowObj.Status) rowObj.Status = "Pending";
    rows.push(rowObj);
  }

  return { sheet: sheet, headers: headers, rows: rows };
}

/**
 * Render template with token replacement
 */
function renderTemplate(data) {
  var template = HtmlService.createTemplateFromFile("index");
  var html = template.getRawContent();

  html = html.replace(/{{(.*?)}}/g, function (match, token) {
    token = token.trim();
    return (data && token in data) ? data[token] : "";
  });

  return html;
}

/**
 * Send one email
 */
function sendEmail_(recipient, body) {
  MailApp.sendEmail({
    to: recipient.Email,
    subject: "Proposal from " + recipient.CompanyName,
    htmlBody: body,
    name: "Shanley Clarence Lacanlale – Project Manager"
  });
}

/**
 * Update sheet log (Status, LastSent, LOI Link)
 */
function updateLog_(sheet, headers, rowIndex, status, error, loiLink) {
  var statusCol = headers.indexOf("Status") + 1;
  var lastSentCol = headers.indexOf("LastSent") + 1;
  var loiCol = headers.indexOf("LOI Link") + 1;

  if (status.startsWith("Sent")) {
    sheet.getRange(rowIndex, statusCol).setValue("Sent");
    sheet.getRange(rowIndex, lastSentCol).setValue(new Date());
    if (loiLink) {
      sheet.getRange(rowIndex, loiCol).setValue(loiLink);
    }
  } else {
    sheet.getRange(rowIndex, statusCol).setValue("Error: " + error);
  }
}

/**
 * Master function – always send emails, update LastSent & LOI link
 */
function sendAllEmails() {
  try {
    var result = getRecipients_();
    var sheet = result.sheet;
    var headers = result.headers;

    result.rows.forEach(function (row) {
      try {
        // Generate unique LOI and get link
        var loiLink = generateLOI_(row);

        var data = Object.assign({}, row, {
          ProjectName: "University of Santo Tomas Software Engineering Team",
          "University Name": "University of Santo Tomas",
          LOI_Link: loiLink, // token used in index.html
          YourName: "Shanley Clarence Lacanlale",
          Role: "Project Manager",
          Email: "shanleyclarence.lacanlale.cics@ust.edu.ph",
          ContactNumber: "0960 293 7255",
          Date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy")
        });

        var body = renderTemplate(data);
        sendEmail_(row, body);

        // Always overwrite status + update timestamp + LOI link
        updateLog_(sheet, headers, row._rowIndex, "Sent", null, loiLink);
        Logger.log("📩 Sent to " + row.Email);

      } catch (err) {
        Logger.log("❌ Error for " + row.Email + ": " + err.message);
        updateLog_(sheet, headers, row._rowIndex, "Error", err.message);
      }
    });

  } catch (bulkErr) {
    Logger.log("🔥 Bulk send failed: " + bulkErr.message);
    try {
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: "🚨 Auto Email System Failure",
        body: "Bulk email run failed:\n\n" + bulkErr.message
      });
    } catch (notifyErr) {
      Logger.log("⚠️ Could not notify sender: " + notifyErr.message);
    }
  }
}
