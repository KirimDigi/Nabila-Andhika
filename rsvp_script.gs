// --- GOOGLE APPS SCRIPT FOR RSVP & GUEST WISHES ---
// Spreadsheet ID: 1qUDD6d6l-FaU_6JewIFVa-3E6OanFxYVD_E9HM_fqpI
// Sheet Name: Sheet1

const SPREADSHEET_ID = "1qUDD6d6l-FaU_6JewIFVa-3E6OanFxYVD_E9HM_fqpI";
const SHEET_NAME = "Sheet1";

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const rows = [];
    
    // Skip header row (index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[1]) continue; // Skip if Name is empty
      rows.push({
        timestamp: row[0],
        name: row[1],
        comment: row[2],
        attendance: row[3],
        guest: row[4]
      });
    }
    
    // Sort by timestamp descending (newest first)
    rows.reverse();
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    let name = "";
    let comment = "";
    let attendance = "";
    let guest = "1";
    
    if (e.postData && e.postData.contents) {
      // Try parsing as JSON
      try {
        const payload = JSON.parse(e.postData.contents);
        name = payload.name || "";
        comment = payload.comment || "";
        attendance = payload.attendance || "";
        guest = payload.guest || "1";
      } catch (err) {
        // Fallback to URL parameters
        name = e.parameter.name || "";
        comment = e.parameter.comment || "";
        attendance = e.parameter.attendance || "";
        guest = e.parameter.guest || "1";
      }
    } else {
      name = e.parameter.name || "";
      comment = e.parameter.comment || "";
      attendance = e.parameter.attendance || "";
      guest = e.parameter.guest || "1";
    }
    
    if (!name || !comment) {
      throw new Error("Nama dan ucapan wajib diisi");
    }
    
    // Format attendance text for spreadsheet clarity
    let attendanceText = "";
    if (attendance === "present" || attendance.toLowerCase() === "hadir") {
      attendanceText = "Hadir";
    } else if (attendance === "notpresent" || attendance.toLowerCase() === "tidak hadir") {
      attendanceText = "Tidak Hadir";
    } else {
      attendanceText = attendance || "Hadir";
    }
    
    const timestamp = new Date();
    sheet.appendRow([timestamp, name, comment, attendanceText, guest]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
