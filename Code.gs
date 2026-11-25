/**
 * Google Apps Script for SABTMHSS THAYINERI Student Transport Portal
 * Educator: Rajula P
 *
 * Capabilities:
 * 1. doPost:
 *    - Standard student survey submission with duplicate phone number check.
 *    - Delete single record (action: "delete") by phone number / student name.
 *    - Clear all test data (action: "clearAll") keeping the header row intact.
 * 2. doGet:
 *    - Serves all student data in JSON format to admin.html dashboard.
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    
    // -------------------------------------------------------------
    // ACTION: DELETE A SPECIFIC RECORD
    // -------------------------------------------------------------
    if (data.action === "delete") {
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return ContentService
          .createTextOutput(JSON.stringify({ "result": "error", "error": "No records to delete." }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var phoneCol = -1;
      var nameCol = -1;
      var timeCol = -1;
      
      for (var h = 0; h < headers.length; h++) {
        var hLower = headers[h].toString().toLowerCase();
        if (hLower.indexOf("phone") !== -1) phoneCol = h;
        if (hLower.indexOf("name") !== -1) nameCol = h;
        if (hLower.indexOf("timestamp") !== -1) timeCol = h;
      }
      
      var targetPhone = (data.phoneNo || "").toString().replace(/\D/g, "");
      if (targetPhone.length > 10) targetPhone = targetPhone.slice(-10);
      var targetName = (data.studentName || "").toString().trim().toLowerCase();
      var targetTime = (data.timestamp || "").toString().trim();
      
      var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      var deleted = false;
      
      // Traverse bottom-up so row deletion does not disturb indexing
      for (var r = values.length - 1; r >= 0; r--) {
        var rPhone = (phoneCol !== -1 && values[r][phoneCol]) ? values[r][phoneCol].toString().replace(/\D/g, "") : "";
        if (rPhone.length > 10) rPhone = rPhone.slice(-10);
        var rName = (nameCol !== -1 && values[r][nameCol]) ? values[r][nameCol].toString().trim().toLowerCase() : "";
        var rTime = (timeCol !== -1 && values[r][timeCol]) ? values[r][timeCol].toString().trim() : "";
        
        var match = false;
        if (targetPhone && targetPhone.length === 10 && rPhone === targetPhone) {
          match = true;
        } else if (targetName && targetTime && rName === targetName && rTime === targetTime) {
          match = true;
        } else if (targetName && targetPhone && rName === targetName && rPhone === targetPhone) {
          match = true;
        }
        
        if (match) {
          sheet.deleteRow(r + 2); // +2 for 1-based index and header row
          deleted = true;
          break;
        }
      }
      
      if (deleted) {
        return ContentService
          .createTextOutput(JSON.stringify({ "result": "success", "message": "Record deleted successfully" }))
          .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService
          .createTextOutput(JSON.stringify({ "result": "error", "error": "Matching record not found in sheet" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // -------------------------------------------------------------
    // ACTION: CLEAR ALL DATA ROWS (KEEP HEADER)
    // -------------------------------------------------------------
    if (data.action === "clearAll") {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ "result": "success", "message": "All data cleared successfully" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // -------------------------------------------------------------
    // DEFAULT ACTION: NEW STUDENT SUBMISSION
    // -------------------------------------------------------------
    var headers = [];
    if (sheet.getLastRow() === 0) {
      headers = ["Timestamp", "Student Name", "Roll Number", "Phone Number", "Class / Grade", "Division", "Transport Mode", "Bus Number"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#1fc1d1").setFontColor("#ffffff");
    } else {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var phoneHeaderIdx = -1;
      for (var h = 0; h < headers.length; h++) {
        if (headers[h].toString().toLowerCase().indexOf("phone") !== -1) {
          phoneHeaderIdx = h;
          break;
        }
      }
      if (phoneHeaderIdx === -1) {
        sheet.insertColumnAfter(3);
        sheet.getRange(1, 4).setValue("Phone Number").setFontWeight("bold").setBackground("#1fc1d1").setFontColor("#ffffff");
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      }
    }
    
    // Clean submitted phone number (last 10 digits)
    var inputPhoneDigits = (data.phoneNo || "").toString().replace(/\D/g, "");
    if (inputPhoneDigits.length > 10) inputPhoneDigits = inputPhoneDigits.slice(-10);
    
    // Duplicate Phone Number Check
    if (inputPhoneDigits.length === 10) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        var phoneColIdx = 3;
        for (var p = 0; p < headers.length; p++) {
          if (headers[p].toString().toLowerCase().indexOf("phone") !== -1) {
            phoneColIdx = p;
            break;
          }
        }
        
        var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        for (var i = 0; i < values.length; i++) {
          var existingPhoneVal = values[i][phoneColIdx] ? values[i][phoneColIdx].toString().replace(/\D/g, "") : "";
          if (existingPhoneVal.length > 10) existingPhoneVal = existingPhoneVal.slice(-10);
          
          if (existingPhoneVal.length === 10 && existingPhoneVal === inputPhoneDigits) {
            return ContentService
              .createTextOutput(JSON.stringify({ "result": "error", "error": "Duplicate phone number: " + data.phoneNo }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
    }
    
    var rowToAppend = new Array(headers.length).fill("");
    for (var j = 0; j < headers.length; j++) {
      var headerName = headers[j].toString().toLowerCase();
      if (headerName.indexOf("timestamp") !== -1) rowToAppend[j] = data.timestamp || new Date().toLocaleString();
      else if (headerName.indexOf("student name") !== -1 || headerName.indexOf("name") !== -1) rowToAppend[j] = data.studentName || "";
      else if (headerName.indexOf("roll") !== -1) rowToAppend[j] = data.rollNo || "N/A";
      else if (headerName.indexOf("phone") !== -1) rowToAppend[j] = data.phoneNo || "";
      else if (headerName.indexOf("class") !== -1 || headerName.indexOf("grade") !== -1) rowToAppend[j] = data.classGrade || "";
      else if (headerName.indexOf("division") !== -1) rowToAppend[j] = data.division || "";
      else if (headerName.indexOf("transport") !== -1 || headerName.indexOf("mode") !== -1) rowToAppend[j] = data.transportMode || "";
      else if (headerName.indexOf("bus") !== -1) rowToAppend[j] = data.busNumber || "N/A";
    }
    
    sheet.appendRow(rowToAppend);
    
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    var rows = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = rows[0].map(function(h) { return h.toString().trim(); });
    var result = [];
    
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      var item = {
        timestamp: "",
        studentName: "",
        rollNo: "N/A",
        phoneNo: "-",
        classGrade: "",
        division: "",
        transportMode: "",
        busNumber: "N/A"
      };
      
      for (var j = 0; j < headers.length; j++) {
        var val = row[j];
        var hName = headers[j].toLowerCase();
        
        if (hName.indexOf("timestamp") !== -1) item.timestamp = val;
        else if (hName.indexOf("student name") !== -1 || hName.indexOf("name") !== -1) item.studentName = val;
        else if (hName.indexOf("roll") !== -1) item.rollNo = val || "N/A";
        else if (hName.indexOf("phone") !== -1) item.phoneNo = val || "-";
        else if (hName.indexOf("class") !== -1 || hName.indexOf("grade") !== -1) item.classGrade = val;
        else if (hName.indexOf("division") !== -1) item.division = val;
        else if (hName.indexOf("transport") !== -1 || hName.indexOf("mode") !== -1) item.transportMode = val;
        else if (hName.indexOf("bus") !== -1) item.busNumber = val || "N/A";
      }
      
      if (!item.studentName && row[1]) item.studentName = row[1];
      if ((!item.phoneNo || item.phoneNo === "-") && row.length >= 8 && row[3]) item.phoneNo = row[3];
      
      result.push(item);
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
