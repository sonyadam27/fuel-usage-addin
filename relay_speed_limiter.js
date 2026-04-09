/**
 * MyGeotab Relay Speed Limiter
 * Database: alamui_indonesia
 *
 * This Google Apps Script:
 * 1. Receives POST from MyGeotab rule notification when speed > 100 km/h
 * 2. Disables the vehicle relay by removing AddInData (NFC key authorization)
 * 3. When speed drops below 100 km/h, re-enables the relay by adding AddInData back
 * 4. Logs all events to a Google Sheet
 * 5. Sends email notifications
 *
 * Action detection (both supported):
 *   - Query parameter: ?action=disable or ?action=enable
 *   - Rule name keywords: "over"/"exceed"/"above" = disable; "under"/"below"/"normal" = enable
 *
 * AUTHENTICATION: Uses username/password (BasicAuthentication service account)
 */

// ==================== CONFIG ====================
var CONFIG = {
  server: "my.geotab.com",
  database: "alamui_indonesia",

  // Service account credentials (no token expiry)
  username: "api.adapter@alamui.service",
  password: "Fl33tD@t@!2026x",

  // AddIn ID for relay control (from your MyGeotab setup)
  addInId: "aS_lt7cUYYEutQoXZGoQPZq",

  // Google Sheet ID for logging (CREATE A NEW SHEET and paste its ID here)
  // To get Sheet ID: open the sheet URL, copy the part between /d/ and /edit
  sheetId: "1RRNt7mXQVGeOQQW1j6i9k6QpSoprZR2FoHGicMHZaNA",
  logSheetName: "Relay Control Log",
  stateSheetName: "Saved State",

  // Email notification
  notifyEmail: "sonyadam273@gmail.com",
  sendEmail: true,

  // Speed threshold
  speedLimit: 100
};
// ================================================

/**
 * Handle POST requests from MyGeotab web request notifications
 */
function doPost(e) {
  try {
    // Parse POST data
    var postData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        postData = e.parameter || {};
      }
    } else if (e && e.parameter) {
      postData = e.parameter;
    }

    // Extract device info
    var deviceId = postData.DeviceId || postData.deviceId || postData["Device ID"] || "";
    var deviceName = postData.DeviceName || postData.deviceName || postData["Device name"] || postData["Device"] || "";
    var driverName = postData.DriverName || postData.driverName || postData["Device With Driver Name"] || "";
    var ruleName = postData.RuleName || postData.ruleName || postData["Rule name"] || "";
    var date = postData.Date || postData.date || new Date().toISOString();
    var address = postData.Address || postData.address || "";

    // Determine action: query parameter first, then rule name fallback
    var action = "";

    // Method 1: Query parameter
    if (e && e.parameter && e.parameter.action) {
      action = e.parameter.action.toLowerCase();
    }

    // Method 2: Rule name detection (fallback)
    if (!action && ruleName) {
      var ruleNameLower = ruleName.toLowerCase();
      if (ruleNameLower.indexOf("over") >= 0 || ruleNameLower.indexOf("exceed") >= 0 ||
          ruleNameLower.indexOf("above") >= 0 || ruleNameLower.indexOf("lebih") >= 0) {
        action = "disable";
      } else if (ruleNameLower.indexOf("under") >= 0 || ruleNameLower.indexOf("below") >= 0 ||
                 ruleNameLower.indexOf("normal") >= 0 || ruleNameLower.indexOf("kurang") >= 0) {
        action = "enable";
      }
    }

    if (!action) {
      action = "disable"; // default to disable for safety
    }

    // Get current vehicle speed
    var speed = getDeviceSpeed(deviceId);

    // Log initial entry
    logToSheet({
      timestamp: new Date().toISOString(),
      deviceId: deviceId,
      deviceName: deviceName,
      driverName: driverName,
      ruleName: ruleName,
      action: action,
      speed: speed,
      status: "Processing..."
    });

    var result = {};

    if (action === "disable") {
      result = disableRelay(deviceId, deviceName);
    } else if (action === "enable") {
      result = enableRelay(deviceId, deviceName);
    } else {
      result = { success: false, message: "Unknown action: " + action };
    }

    // Send email notification
    if (CONFIG.sendEmail && CONFIG.notifyEmail) {
      sendNotificationEmail({
        deviceName: deviceName,
        driverName: driverName,
        ruleName: ruleName,
        action: action,
        speed: speed,
        date: date,
        address: address,
        result: result
      });
    }

    // Update log status
    updateLastLogStatus(result.success ? "Success - " + action : "Failed - " + (result.message || "unknown"));

    return ContentService.createTextOutput(JSON.stringify({
      success: result.success,
      action: action,
      deviceId: deviceId,
      deviceName: deviceName,
      message: result.message
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error in doPost: " + err.message);
    updateLastLogStatus("Error: " + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle GET requests (health check)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "Relay Speed Limiter is running. Use POST to trigger.",
    speedLimit: CONFIG.speedLimit,
    timestamp: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

// ==================== MyGeotab API ====================

/**
 * Authenticate with MyGeotab
 */
function authenticate() {
  var url = "https://" + CONFIG.server + "/apiv1";

  var payload = {
    method: "Authenticate",
    params: {
      database: CONFIG.database,
      userName: CONFIG.username,
      password: CONFIG.password
    }
  };

  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());

  if (result.error) {
    throw new Error("Auth failed: " + JSON.stringify(result.error));
  }

  return result.result.credentials;
}

/**
 * Call MyGeotab API
 */
function callApi(method, params) {
  var credentials = authenticate();

  var url = "https://" + CONFIG.server + "/apiv1";

  var allParams = {};
  if (params) {
    for (var key in params) {
      allParams[key] = params[key];
    }
  }
  allParams.credentials = credentials;

  var payload = {
    method: method,
    params: allParams
  };

  var response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var result = JSON.parse(response.getContentText());

  if (result.error) {
    Logger.log("API error (" + method + "): " + JSON.stringify(result.error));
    throw new Error("API " + method + " failed: " + JSON.stringify(result.error));
  }

  return result.result;
}

// ==================== Relay Control ====================

/**
 * Disable relay: query AddInData for the device, save state, then remove
 */
function disableRelay(deviceId, deviceName) {
  Logger.log("Disabling relay for device: " + deviceId + " (" + deviceName + ")");

  // Check if already disabled (state exists)
  var existingState = getSavedState(deviceId);
  if (existingState) {
    Logger.log("Relay already disabled for device: " + deviceId);
    return { success: true, message: "Already disabled" };
  }

  // Query AddInData for this device
  var addInDataList;
  try {
    addInDataList = callApi("Get", {
      typeName: "AddInData",
      search: {
        addInId: CONFIG.addInId,
        deviceSearch: { id: deviceId }
      }
    });
  } catch (err) {
    // Try without device filter if the search doesn't support it
    Logger.log("Filtered search failed, trying broader search: " + err.message);
    addInDataList = callApi("Get", {
      typeName: "AddInData",
      search: {
        addInId: CONFIG.addInId
      }
    });

    // Filter manually by device
    if (addInDataList && addInDataList.length > 0) {
      addInDataList = addInDataList.filter(function(item) {
        return item.details && item.details.vehicle === deviceId;
      });
    }
  }

  if (!addInDataList || addInDataList.length === 0) {
    Logger.log("No AddInData found for device: " + deviceId);
    return { success: false, message: "No AddInData found for device" };
  }

  // Save state for each AddInData entry (so we can restore later)
  var removedCount = 0;
  for (var i = 0; i < addInDataList.length; i++) {
    var addInData = addInDataList[i];

    // Save to state sheet
    saveState(deviceId, addInData.id, JSON.stringify(addInData));

    // Remove the AddInData
    try {
      callApi("Remove", {
        typeName: "AddInData",
        entity: {
          id: addInData.id
        }
      });
      removedCount++;
      Logger.log("Removed AddInData: " + addInData.id);
    } catch (err) {
      Logger.log("Failed to remove AddInData " + addInData.id + ": " + err.message);
    }
  }

  return {
    success: removedCount > 0,
    message: "Removed " + removedCount + " AddInData entries"
  };
}

/**
 * Enable relay: read saved state and add AddInData back
 */
function enableRelay(deviceId, deviceName) {
  Logger.log("Enabling relay for device: " + deviceId + " (" + deviceName + ")");

  // Read saved state
  var savedStates = getAllSavedStates(deviceId);

  if (!savedStates || savedStates.length === 0) {
    Logger.log("No saved state found for device: " + deviceId);
    return { success: false, message: "No saved state to restore" };
  }

  var addedCount = 0;
  for (var i = 0; i < savedStates.length; i++) {
    var state = savedStates[i];
    try {
      var addInData = JSON.parse(state.json);

      // Build the Add entity from saved data
      var entity = {
        addInId: addInData.addInId || CONFIG.addInId,
        details: addInData.details
      };

      // Update the date to current time
      if (entity.details) {
        entity.details.date = new Date().toISOString();
      }

      callApi("Add", {
        typeName: "AddInData",
        entity: entity
      });

      addedCount++;
      Logger.log("Restored AddInData for device: " + deviceId);

      // Clear the saved state row
      clearSavedState(state.row);

    } catch (err) {
      Logger.log("Failed to restore AddInData: " + err.message);
    }
  }

  return {
    success: addedCount > 0,
    message: "Restored " + addedCount + " AddInData entries"
  };
}

/**
 * Get current device speed from DeviceStatusInfo
 */
function getDeviceSpeed(deviceId) {
  try {
    var result = callApi("Get", {
      typeName: "DeviceStatusInfo",
      search: {
        deviceSearch: { id: deviceId }
      }
    });

    if (result && result.length > 0) {
      return result[0].speed || 0;
    }
  } catch (err) {
    Logger.log("Could not get device speed: " + err.message);
  }
  return 0;
}

// ==================== State Management (Google Sheets) ====================

/**
 * Save AddInData state before removal
 */
function saveState(deviceId, addInDataId, jsonStr) {
  try {
    var sheet = getOrCreateSheet(CONFIG.stateSheetName, ["Device ID", "AddInData ID", "AddInData JSON", "Saved At"]);
    sheet.appendRow([deviceId, addInDataId, jsonStr, new Date().toISOString()]);
  } catch (err) {
    Logger.log("Save state error: " + err.message);
  }
}

/**
 * Get saved state for a device
 */
function getSavedState(deviceId) {
  try {
    var sheet = getOrCreateSheet(CONFIG.stateSheetName, ["Device ID", "AddInData ID", "AddInData JSON", "Saved At"]);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deviceId) {
        return { row: i + 1, deviceId: data[i][0], addInDataId: data[i][1], json: data[i][2] };
      }
    }
  } catch (err) {
    Logger.log("Get state error: " + err.message);
  }
  return null;
}

/**
 * Get all saved states for a device
 */
function getAllSavedStates(deviceId) {
  try {
    var sheet = getOrCreateSheet(CONFIG.stateSheetName, ["Device ID", "AddInData ID", "AddInData JSON", "Saved At"]);
    var data = sheet.getDataRange().getValues();
    var states = [];

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === deviceId) {
        states.push({ row: i + 1, deviceId: data[i][0], addInDataId: data[i][1], json: data[i][2] });
      }
    }
    return states;
  } catch (err) {
    Logger.log("Get all states error: " + err.message);
  }
  return [];
}

/**
 * Clear a saved state row after restoration
 */
function clearSavedState(rowNumber) {
  try {
    var sheet = getOrCreateSheet(CONFIG.stateSheetName, ["Device ID", "AddInData ID", "AddInData JSON", "Saved At"]);
    sheet.deleteRow(rowNumber);
  } catch (err) {
    Logger.log("Clear state error: " + err.message);
  }
}

// ==================== Logging ====================

/**
 * Get or create a sheet with headers
 */
function getOrCreateSheet(sheetName, headers) {
  var ss = SpreadsheetApp.openById(CONFIG.sheetId);
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Log event to the relay control log sheet
 */
function logToSheet(data) {
  try {
    var sheet = getOrCreateSheet(CONFIG.logSheetName, [
      "Timestamp", "Device ID", "Device Name", "Driver",
      "Rule Name", "Action", "Speed (km/h)", "Status"
    ]);

    sheet.appendRow([
      data.timestamp,
      data.deviceId,
      data.deviceName,
      data.driverName,
      data.ruleName,
      data.action,
      data.speed,
      data.status
    ]);
  } catch (err) {
    Logger.log("Sheet logging error: " + err.message);
  }
}

/**
 * Update status of the last log entry
 */
function updateLastLogStatus(status) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.sheetId);
    var sheet = ss.getSheetByName(CONFIG.logSheetName);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(lastRow, 8).setValue(status);
      }
    }
  } catch (err) {
    Logger.log("Status update error: " + err.message);
  }
}

// ==================== Email Notification ====================

function sendNotificationEmail(data) {
  var actionLabel = data.action === "disable" ? "DISABLED (Speed Over " + CONFIG.speedLimit + " km/h)" : "ENABLED (Speed Normal)";
  var actionColor = data.action === "disable" ? "#dc3545" : "#28a745";

  var subject = "Relay " + (data.action === "disable" ? "DISABLED" : "ENABLED") + ": " + (data.deviceName || "Unknown Vehicle");

  var body =
    "Relay " + actionLabel + "\n\n" +
    "Vehicle: " + (data.deviceName || "N/A") + "\n" +
    "Driver: " + (data.driverName || "N/A") + "\n" +
    "Rule: " + (data.ruleName || "N/A") + "\n" +
    "Speed: " + (data.speed || "N/A") + " km/h\n" +
    "Date: " + (data.date || "N/A") + "\n" +
    "Address: " + (data.address || "N/A") + "\n" +
    "Result: " + (data.result.message || "N/A") + "\n" +
    "\n---\nSent by MyGeotab Relay Speed Limiter (alamui_indonesia)";

  var htmlBody =
    "<div style='font-family:Arial,sans-serif;max-width:600px;'>" +
    "<h2 style='color:" + actionColor + ";'>Relay " + actionLabel + "</h2>" +
    "<table style='border-collapse:collapse;width:100%;'>" +
    "<tr><td style='padding:8px 12px;font-weight:bold;border-bottom:1px solid #eee;'>Vehicle:</td>" +
    "<td style='padding:8px 12px;border-bottom:1px solid #eee;'>" + (data.deviceName || "N/A") + "</td></tr>" +
    "<tr><td style='padding:8px 12px;font-weight:bold;border-bottom:1px solid #eee;'>Driver:</td>" +
    "<td style='padding:8px 12px;border-bottom:1px solid #eee;'>" + (data.driverName || "N/A") + "</td></tr>" +
    "<tr><td style='padding:8px 12px;font-weight:bold;border-bottom:1px solid #eee;'>Rule:</td>" +
    "<td style='padding:8px 12px;border-bottom:1px solid #eee;'>" + (data.ruleName || "N/A") + "</td></tr>" +
    "<tr><td style='padding:8px 12px;font-weight:bold;border-bottom:1px solid #eee;'>Speed:</td>" +
    "<td style='padding:8px 12px;border-bottom:1px solid #eee;'><strong>" + (data.speed || "N/A") + " km/h</strong></td></tr>" +
    "<tr><td style='padding:8px 12px;font-weight:bold;border-bottom:1px solid #eee;'>Date:</td>" +
    "<td style='padding:8px 12px;border-bottom:1px solid #eee;'>" + (data.date || "N/A") + "</td></tr>" +
    "<tr><td style='padding:8px 12px;font-weight:bold;border-bottom:1px solid #eee;'>Address:</td>" +
    "<td style='padding:8px 12px;border-bottom:1px solid #eee;'>" + (data.address || "N/A") + "</td></tr>" +
    "<tr><td style='padding:8px 12px;font-weight:bold;'>Result:</td>" +
    "<td style='padding:8px 12px;'>" + (data.result.message || "N/A") + "</td></tr>" +
    "</table>" +
    "<hr style='margin-top:20px;'>" +
    "<p style='color:#888;font-size:12px;'>Sent by MyGeotab Relay Speed Limiter (alamui_indonesia)</p>" +
    "</div>";

  try {
    MailApp.sendEmail({
      to: CONFIG.notifyEmail,
      subject: subject,
      body: body,
      htmlBody: htmlBody
    });
    Logger.log("Email sent to " + CONFIG.notifyEmail);
  } catch (err) {
    Logger.log("Email error: " + err.message);
  }
}

// ==================== TEST FUNCTIONS ====================

/**
 * Test disable relay (simulates speed > 100 km/h rule trigger)
 */
function testDisableRelay() {
  var mockEvent = {
    parameter: { action: "disable" },
    postData: {
      contents: JSON.stringify({
        DeviceId: "bB7",
        DeviceName: "Test Vehicle",
        Database: "alamui_indonesia",
        Date: new Date().toISOString(),
        Address: "Jakarta, Indonesia",
        DriverName: "Test Driver",
        RuleName: "Speed Over 100"
      }),
      type: "application/json"
    }
  };

  var result = doPost(mockEvent);
  Logger.log("Test disable result: " + result.getContent());
}

/**
 * Test enable relay (simulates speed back below 100 km/h)
 */
function testEnableRelay() {
  var mockEvent = {
    parameter: { action: "enable" },
    postData: {
      contents: JSON.stringify({
        DeviceId: "bB7",
        DeviceName: "Test Vehicle",
        Database: "alamui_indonesia",
        Date: new Date().toISOString(),
        Address: "Jakarta, Indonesia",
        DriverName: "Test Driver",
        RuleName: "Speed Under 100"
      }),
      type: "application/json"
    }
  };

  var result = doPost(mockEvent);
  Logger.log("Test enable result: " + result.getContent());
}
