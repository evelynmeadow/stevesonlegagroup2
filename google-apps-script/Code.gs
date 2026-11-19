/**
 * Stevenson Legal Group - Consultation form receiver
 *
 * IMPORTANT:
 * Open this Apps Script project from the target Google Sheet:
 * Google Sheet > Extensions > Apps Script.
 * This script uses the active spreadsheet, so it must be bound to the
 * exact spreadsheet that contains the "Consultation" tab.
 */

const CONSULTATION_TAB_NAME = "Consultation";

const CONSULTATION_HEADERS = [
  "Submission Date",
  "Full Name",
  "Email",
  "Phone",
  "Contact Method",
  "Service",
  "Location",
  "Preferred Date",
  "Preferred Time",
  "Opposing Party",
  "Deadline",
  "Summary",
  "Consent"
];

function doGet() {
  return jsonResponse({
    ok: true,
    service: "Stevenson Legal Group consultation receiver",
    message: "Web app is running. Submit the consultation form to add a row."
  });
}

function doPost(e) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      throw new Error(
        "No active spreadsheet. Open this Apps Script from the target Google Sheet using Extensions > Apps Script."
      );
    }

    const sheet = spreadsheet.getSheetByName(CONSULTATION_TAB_NAME);
    if (!sheet) {
      throw new Error(
        'The spreadsheet does not contain a tab named "' +
        CONSULTATION_TAB_NAME + '".'
      );
    }

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("The request body is empty.");
    }

    const data = JSON.parse(e.postData.contents);

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    try {
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(CONSULTATION_HEADERS);
      }

      sheet.appendRow([
        new Date(),
        data.fullName || "",
        data.email || "",
        data.phone || "",
        data.contactMethod || "",
        data.service || "",
        data.location || "",
        data.preferredDate || "",
        data.preferredTime || "",
        data.opposingParty || "",
        data.deadline || "",
        data.summary || "",
        data.consent ? "Yes" : "No"
      ]);

      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }

    return jsonResponse({ ok: true, saved: true });
  } catch (error) {
    console.error("Consultation submission error:", error);
    return jsonResponse({
      ok: false,
      saved: false,
      error: String(error && error.message ? error.message : error)
    });
  }
}

function jsonResponse(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
