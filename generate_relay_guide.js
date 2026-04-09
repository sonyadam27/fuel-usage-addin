const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak, TabStopType, TabStopPosition
} = require("docx");

// Read the source code file
const sourceCode = fs.readFileSync("C:/Users/sonyadam/Desktop/fuel-usage-addin/relay_speed_limiter.js", "utf-8");

// ==================== HELPERS ====================

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const thickBorder = { style: BorderStyle.SINGLE, size: 2, color: "333333" };
const thickBorders = { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder };

const TABLE_WIDTH = 9360;
const CELL_MARGINS = { top: 60, bottom: 60, left: 100, right: 100 };

function headerCell(text, width) {
  return new TableCell({
    borders: thickBorders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: "1F4E79", type: ShadingType.CLEAR },
    margins: CELL_MARGINS,
    verticalAlign: "center",
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, font: "Arial", size: 20, color: "FFFFFF" })] })]
  });
}

function cell(text, width, opts = {}) {
  const runs = [];
  if (opts.bold) {
    runs.push(new TextRun({ text, bold: true, font: "Arial", size: 20 }));
  } else {
    runs.push(new TextRun({ text, font: "Arial", size: 20 }));
  }
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    margins: CELL_MARGINS,
    children: [new Paragraph({ children: runs })]
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 32, color: "1F4E79" })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 26, color: "2E75B6" })]
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 22, color: "404040" })]
  });
}

function para(text, opts = {}) {
  const runOpts = { text, font: "Arial", size: 20 };
  if (opts.bold) runOpts.bold = true;
  if (opts.italic) runOpts.italics = true;
  if (opts.color) runOpts.color = opts.color;
  return new Paragraph({
    spacing: { after: opts.afterSpacing || 120 },
    children: [new TextRun(runOpts)]
  });
}

function codePara(text) {
  return new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text, font: "Consolas", size: 16, color: "333333" })]
  });
}

function codeBlock(code) {
  const lines = code.split("\n");
  return lines.map(line =>
    new Paragraph({
      spacing: { after: 0, line: 240 },
      indent: { left: 360 },
      children: [new TextRun({ text: line || " ", font: "Consolas", size: 15, color: "333333" })]
    })
  );
}

function bulletItem(text, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 20 })]
  });
}

function numberedItem(text, ref = "steps") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, font: "Arial", size: 20 })]
  });
}

function numberedItemBold(label, text, ref = "steps") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80 },
    children: [
      new TextRun({ text: label, bold: true, font: "Arial", size: 20 }),
      new TextRun({ text, font: "Arial", size: 20 })
    ]
  });
}

function importantNote(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    indent: { left: 360 },
    children: [
      new TextRun({ text: "IMPORTANT: ", bold: true, font: "Arial", size: 20, color: "CC0000" }),
      new TextRun({ text, font: "Arial", size: 20, color: "CC0000" })
    ]
  });
}

function spacer() {
  return new Paragraph({ spacing: { after: 80 }, children: [] });
}

// ==================== DOCUMENT CONTENT ====================

const children = [];

// Title page
children.push(new Paragraph({ spacing: { before: 3000 }, children: [] }));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({ text: "MyGeotab Relay Speed Limiter", font: "Arial", size: 48, bold: true, color: "1F4E79" })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "Complete Setup Guide", font: "Arial", size: 36, color: "2E75B6" })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 600 },
  children: [new TextRun({ text: "Database: alamui_indonesia", font: "Arial", size: 24, color: "666666", italics: true })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "Version 1.0  |  March 2026", font: "Arial", size: 22, color: "888888" })]
}));
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 120 },
  children: [new TextRun({ text: "Deployed by: sonyadam273@gmail.com", font: "Arial", size: 20, color: "888888" })]
}));

// Page break
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== TABLE OF CONTENTS ====================
children.push(heading1("Table of Contents"));
const tocItems = [
  "1. Overview",
  "2. System Architecture",
  "3. Configuration Reference",
  "4. MyGeotab Rules Reference",
  "5. Webhook Notifications Reference",
  "6. Step-by-Step Setup Guide",
  "7. Google Sheets Structure",
  "8. How It Works - Technical Details",
  "9. Complete Source Code",
  "10. MyGeotab Rule JSON Definitions",
  "11. Troubleshooting",
  "12. File Locations"
];
tocItems.forEach(item => {
  children.push(new Paragraph({
    spacing: { after: 100 },
    indent: { left: 360 },
    children: [new TextRun({ text: item, font: "Arial", size: 22, color: "2E75B6" })]
  }));
});
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 1. OVERVIEW ====================
children.push(heading1("1. Overview"));
children.push(para("The MyGeotab Relay Speed Limiter is a Google Apps Script-based system that automatically controls vehicle relays on the alamui_indonesia MyGeotab database based on vehicle speed thresholds."));
children.push(spacer());
children.push(heading3("What It Does"));
children.push(bulletItem("Monitors vehicle speed in real-time via MyGeotab exception rules"));
children.push(bulletItem("When a vehicle exceeds 100 km/h, the relay is DISABLED (AddInData is removed from MyGeotab)"));
children.push(bulletItem("When vehicle speed drops below 100 km/h, the relay is RE-ENABLED (AddInData is restored)"));
children.push(bulletItem("All events are logged to a Google Sheet for audit trail"));
children.push(bulletItem("Email notifications are sent for each relay state change"));
children.push(spacer());
children.push(heading3("System Components"));
children.push(bulletItem("MyGeotab Rules: Two exception rules monitor vehicle speed (over/under 100 km/h)"));
children.push(bulletItem("Web Request Notifications: Webhook POST requests sent to Google Apps Script"));
children.push(bulletItem("Google Apps Script: Middleware that processes webhooks and calls MyGeotab API"));
children.push(bulletItem("Google Sheets: Logging and state management (saves removed AddInData for restoration)"));
children.push(bulletItem("Email Notifications: Alerts sent to administrator on each relay change"));
children.push(spacer());
children.push(heading3("Architecture Flow"));
children.push(para("MyGeotab Rule Trigger  -->  Web Request Notification (POST)  -->  Google Apps Script  -->  MyGeotab API (AddInData Add/Remove)", { bold: true }));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 2. SYSTEM ARCHITECTURE ====================
children.push(heading1("2. System Architecture"));
children.push(heading3("Disable Flow (Vehicle Speed > 100 km/h)"));
children.push(numberedItem('Vehicle speed exceeds 100 km/h', "disableSteps"));
children.push(numberedItem('"Speed Over 100" exception rule triggers in MyGeotab', "disableSteps"));
children.push(numberedItem('MyGeotab sends webhook POST to Google Apps Script URL with ?action=disable', "disableSteps"));
children.push(numberedItem('Apps Script receives the POST request and parses device information', "disableSteps"));
children.push(numberedItem('Script authenticates to MyGeotab API using service account credentials', "disableSteps"));
children.push(numberedItem('Script queries AddInData for the specific device (by addInId and deviceId)', "disableSteps"));
children.push(numberedItem('Script saves the full AddInData JSON to the "Saved State" Google Sheet tab', "disableSteps"));
children.push(numberedItem('Script calls MyGeotab API Remove method to delete the AddInData entry', "disableSteps"));
children.push(numberedItem('Event is logged to the "Relay Control Log" sheet tab', "disableSteps"));
children.push(numberedItem('Email notification is sent to the administrator', "disableSteps"));
children.push(spacer());
children.push(heading3("Enable Flow (Vehicle Speed < 100 km/h)"));
children.push(numberedItem('Vehicle speed drops below 100 km/h', "enableSteps"));
children.push(numberedItem('"Speed Under 100" exception rule triggers in MyGeotab', "enableSteps"));
children.push(numberedItem('MyGeotab sends webhook POST to Google Apps Script URL with ?action=enable', "enableSteps"));
children.push(numberedItem('Apps Script receives the POST request and parses device information', "enableSteps"));
children.push(numberedItem('Script authenticates to MyGeotab API using service account credentials', "enableSteps"));
children.push(numberedItem('Script reads saved AddInData from the "Saved State" Google Sheet for that device', "enableSteps"));
children.push(numberedItem('Script calls MyGeotab API Add method with the saved AddInData details', "enableSteps"));
children.push(numberedItem('Saved state row is cleared from the Google Sheet', "enableSteps"));
children.push(numberedItem('Event is logged to the "Relay Control Log" sheet tab', "enableSteps"));
children.push(numberedItem('Email notification is sent to the administrator', "enableSteps"));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 3. CONFIGURATION REFERENCE ====================
children.push(heading1("3. Configuration Reference"));
children.push(para("All configuration values used by the Relay Speed Limiter system:"));
children.push(spacer());

const configCol1 = 3200;
const configCol2 = 6160;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [configCol1, configCol2],
  rows: [
    new TableRow({ children: [headerCell("Parameter", configCol1), headerCell("Value", configCol2)] }),
    new TableRow({ children: [cell("Server", configCol1, { bold: true }), cell("my.geotab.com", configCol2)] }),
    new TableRow({ children: [cell("Database", configCol1, { bold: true }), cell("alamui_indonesia", configCol2)] }),
    new TableRow({ children: [cell("Service Account Username", configCol1, { bold: true }), cell("api.adapter@alamui.service", configCol2)] }),
    new TableRow({ children: [cell("Service Account Password", configCol1, { bold: true }), cell("Fl33tD@t@!2026x", configCol2)] }),
    new TableRow({ children: [cell("AddIn ID", configCol1, { bold: true }), cell("aS_lt7cUYYEutQoXZGoQPZq", configCol2)] }),
    new TableRow({ children: [cell("Google Sheet ID", configCol1, { bold: true }), cell("1RRNt7mXQVGeOQQW1j6i9k6QpSoprZR2FoHGicMHZaNA", configCol2)] }),
    new TableRow({ children: [cell("Log Sheet Tab", configCol1, { bold: true }), cell("Relay Control Log", configCol2)] }),
    new TableRow({ children: [cell("State Sheet Tab", configCol1, { bold: true }), cell("Saved State", configCol2)] }),
    new TableRow({ children: [cell("Notification Email", configCol1, { bold: true }), cell("sonyadam273@gmail.com", configCol2)] }),
    new TableRow({ children: [cell("Email Enabled", configCol1, { bold: true }), cell("true", configCol2)] }),
    new TableRow({ children: [cell("Speed Limit", configCol1, { bold: true }), cell("100 km/h", configCol2)] }),
  ]
}));
children.push(spacer());
children.push(heading3("Deployed Web App URL"));
children.push(new Paragraph({
  spacing: { after: 120 },
  indent: { left: 360 },
  children: [new TextRun({
    text: "https://script.google.com/macros/s/AKfycbyi5tnp36BTpKOqKdjUPsxFsFXMmKWckp0ZuwaYnjdJEhXZ-xr1oysVYsPJ46MhhvoM/exec",
    font: "Consolas", size: 16, color: "2E75B6"
  })]
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 4. MYGEOTAB RULES REFERENCE ====================
children.push(heading1("4. MyGeotab Rules Reference"));
children.push(para("Two exception rules are configured on the alamui_indonesia database:"));
children.push(spacer());

// Rule 1
children.push(heading3('Rule 1: "Speed Over 100"'));
const r1c1 = 3200, r1c2 = 6160;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [r1c1, r1c2],
  rows: [
    new TableRow({ children: [headerCell("Property", r1c1), headerCell("Value", r1c2)] }),
    new TableRow({ children: [cell("Rule Name", r1c1, { bold: true }), cell("Speed Over 100", r1c2)] }),
    new TableRow({ children: [cell("Rule ID", r1c1, { bold: true }), cell("aLUa_Rs9bo0uI_vpB09TvPw", r1c2)] }),
    new TableRow({ children: [cell("Base Type", r1c1, { bold: true }), cell("Custom", r1c2)] }),
    new TableRow({ children: [cell("Condition Type", r1c1, { bold: true }), cell("IsValueMoreThan", r1c2)] }),
    new TableRow({ children: [cell("Condition", r1c1, { bold: true }), cell("Speed > 100 km/h", r1c2)] }),
    new TableRow({ children: [cell("Color", r1c1, { bold: true }), cell("Red (R:220, G:53, B:69)", r1c2, { shading: "FDE8EA" })] }),
    new TableRow({ children: [cell("State", r1c1, { bold: true }), cell("Active (ExceptionRuleStateActiveId)", r1c2)] }),
    new TableRow({ children: [cell("Groups", r1c1, { bold: true }), cell("GroupCompanyId (all vehicles)", r1c2)] }),
    new TableRow({ children: [cell("Active From", r1c1, { bold: true }), cell("2020-01-01", r1c2)] }),
    new TableRow({ children: [cell("Active To", r1c1, { bold: true }), cell("2050-01-01", r1c2)] }),
    new TableRow({ children: [cell("Comment", r1c1, { bold: true }), cell("Triggers when vehicle speed exceeds 100 km/h. Sends web request to disable relay.", r1c2)] }),
  ]
}));
children.push(spacer());

// Rule 2
children.push(heading3('Rule 2: "Speed Under 100"'));
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [r1c1, r1c2],
  rows: [
    new TableRow({ children: [headerCell("Property", r1c1), headerCell("Value", r1c2)] }),
    new TableRow({ children: [cell("Rule Name", r1c1, { bold: true }), cell("Speed Under 100", r1c2)] }),
    new TableRow({ children: [cell("Rule ID", r1c1, { bold: true }), cell("apEGV7MMTyUedc-e3nx0GNw", r1c2)] }),
    new TableRow({ children: [cell("Base Type", r1c1, { bold: true }), cell("Custom", r1c2)] }),
    new TableRow({ children: [cell("Condition Type", r1c1, { bold: true }), cell("IsValueLessThan", r1c2)] }),
    new TableRow({ children: [cell("Condition", r1c1, { bold: true }), cell("Speed < 100 km/h", r1c2)] }),
    new TableRow({ children: [cell("Color", r1c1, { bold: true }), cell("Green (R:40, G:167, B:69)", r1c2, { shading: "E8F5E9" })] }),
    new TableRow({ children: [cell("State", r1c1, { bold: true }), cell("Active (ExceptionRuleStateActiveId)", r1c2)] }),
    new TableRow({ children: [cell("Groups", r1c1, { bold: true }), cell("GroupCompanyId (all vehicles)", r1c2)] }),
    new TableRow({ children: [cell("Active From", r1c1, { bold: true }), cell("2020-01-01", r1c2)] }),
    new TableRow({ children: [cell("Active To", r1c1, { bold: true }), cell("2050-01-01", r1c2)] }),
    new TableRow({ children: [cell("Comment", r1c1, { bold: true }), cell("Triggers when vehicle speed drops below 100 km/h. Sends web request to enable relay.", r1c2)] }),
  ]
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 5. WEBHOOK NOTIFICATIONS ====================
children.push(heading1("5. Webhook Notifications Reference"));
children.push(para("Two web request notification distribution lists are configured to send POST requests to the Google Apps Script web app:"));
children.push(spacer());

children.push(heading3('Notification 1: "Relay Disable - Speed Over 100"'));
const n1c1 = 3200, n1c2 = 6160;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [n1c1, n1c2],
  rows: [
    new TableRow({ children: [headerCell("Property", n1c1), headerCell("Value", n1c2)] }),
    new TableRow({ children: [cell("Name", n1c1, { bold: true }), cell("Relay Disable - Speed Over 100", n1c2)] }),
    new TableRow({ children: [cell("Distribution List ID", n1c1, { bold: true }), cell("b1DD", n1c2)] }),
    new TableRow({ children: [cell("Recipient Type", n1c1, { bold: true }), cell("WebRequest", n1c2)] }),
    new TableRow({ children: [cell("HTTP Method", n1c1, { bold: true }), cell("POST", n1c2)] }),
    new TableRow({ children: [cell("Linked Rule", n1c1, { bold: true }), cell("Speed Over 100 (aLUa_Rs9bo0uI_vpB09TvPw)", n1c2)] }),
    new TableRow({ children: [cell("Notification Template ID", n1c1, { bold: true }), cell("b157", n1c2)] }),
  ]
}));
children.push(spacer());
children.push(para("Webhook URL:", { bold: true }));
children.push(new Paragraph({
  spacing: { after: 120 },
  indent: { left: 360 },
  children: [new TextRun({
    text: "https://script.google.com/macros/s/AKfycbyi5tnp36BTpKOqKdjUPsxFsFXMmKWckp0ZuwaYnjdJEhXZ-xr1oysVYsPJ46MhhvoM/exec?action=disable",
    font: "Consolas", size: 15, color: "CC0000"
  })]
}));
children.push(spacer());

children.push(heading3('Notification 2: "Relay Enable - Speed Under 100"'));
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [n1c1, n1c2],
  rows: [
    new TableRow({ children: [headerCell("Property", n1c1), headerCell("Value", n1c2)] }),
    new TableRow({ children: [cell("Name", n1c1, { bold: true }), cell("Relay Enable - Speed Under 100", n1c2)] }),
    new TableRow({ children: [cell("Distribution List ID", n1c1, { bold: true }), cell("b1DE", n1c2)] }),
    new TableRow({ children: [cell("Recipient Type", n1c1, { bold: true }), cell("WebRequest", n1c2)] }),
    new TableRow({ children: [cell("HTTP Method", n1c1, { bold: true }), cell("POST", n1c2)] }),
    new TableRow({ children: [cell("Linked Rule", n1c1, { bold: true }), cell("Speed Under 100 (apEGV7MMTyUedc-e3nx0GNw)", n1c2)] }),
    new TableRow({ children: [cell("Notification Template ID", n1c1, { bold: true }), cell("b158", n1c2)] }),
  ]
}));
children.push(spacer());
children.push(para("Webhook URL:", { bold: true }));
children.push(new Paragraph({
  spacing: { after: 120 },
  indent: { left: 360 },
  children: [new TextRun({
    text: "https://script.google.com/macros/s/AKfycbyi5tnp36BTpKOqKdjUPsxFsFXMmKWckp0ZuwaYnjdJEhXZ-xr1oysVYsPJ46MhhvoM/exec?action=enable",
    font: "Consolas", size: 15, color: "28A745"
  })]
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 6. STEP-BY-STEP SETUP GUIDE ====================
children.push(heading1("6. Step-by-Step Setup Guide"));
children.push(para("Follow these steps in order to set up the complete Relay Speed Limiter system from scratch."));
children.push(spacer());

// Step 1
children.push(heading2("Step 1: Create Google Sheet for Logging"));
children.push(para("The Google Sheet stores relay control event logs and saved AddInData state for restoration."));
children.push(numberedItem("Go to https://sheets.google.com and sign in with your Google account (sonyadam273@gmail.com)", "step1"));
children.push(numberedItem('Create a new blank spreadsheet by clicking the "+" button', "step1"));
children.push(numberedItem('Name the spreadsheet "Relay Speed Limiter Log" (or any name you prefer)', "step1"));
children.push(numberedItem("Copy the Sheet ID from the browser URL bar. The ID is the long string between /d/ and /edit in the URL.", "step1"));
children.push(new Paragraph({
  spacing: { after: 80 },
  indent: { left: 720 },
  children: [
    new TextRun({ text: "Example URL: ", font: "Arial", size: 18, color: "666666" }),
    new TextRun({ text: "https://docs.google.com/spreadsheets/d/", font: "Consolas", size: 16, color: "666666" }),
    new TextRun({ text: "THIS_IS_THE_SHEET_ID", font: "Consolas", size: 16, color: "CC0000", bold: true }),
    new TextRun({ text: "/edit", font: "Consolas", size: 16, color: "666666" }),
  ]
}));
children.push(numberedItem('You do NOT need to manually create tabs. The script will automatically create "Relay Control Log" and "Saved State" tabs on first run.', "step1"));
children.push(spacer());

// Step 2
children.push(heading2("Step 2: Create Google Apps Script Project"));
children.push(para("Google Apps Script hosts the middleware that processes MyGeotab webhooks."));
children.push(numberedItem("Go to https://script.google.com", "step2"));
children.push(numberedItem('Click "New Project" in the top-left corner', "step2"));
children.push(numberedItem('Click on "Untitled project" at the top and rename it to "Relay Speed Limiter"', "step2"));
children.push(numberedItem("In the editor, select all the default code in Code.gs and delete it", "step2"));
children.push(numberedItem("Copy the entire contents of relay_speed_limiter.js (see Section 9 for the full source code) and paste it into the editor", "step2"));
children.push(numberedItem("Update the CONFIG section at the top of the script with your values:", "step2"));
children.push(spacer());

const configUpdateCol1 = 3000, configUpdateCol2 = 6360;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [configUpdateCol1, configUpdateCol2],
  rows: [
    new TableRow({ children: [headerCell("CONFIG Field", configUpdateCol1), headerCell("What to Set", configUpdateCol2)] }),
    new TableRow({ children: [cell("sheetId", configUpdateCol1, { bold: true }), cell("Paste the Google Sheet ID you copied in Step 1", configUpdateCol2)] }),
    new TableRow({ children: [cell("database", configUpdateCol1, { bold: true }), cell("Your MyGeotab database name (e.g., alamui_indonesia)", configUpdateCol2)] }),
    new TableRow({ children: [cell("username", configUpdateCol1, { bold: true }), cell("Your MyGeotab service account username", configUpdateCol2)] }),
    new TableRow({ children: [cell("password", configUpdateCol1, { bold: true }), cell("Your MyGeotab service account password", configUpdateCol2)] }),
    new TableRow({ children: [cell("addInId", configUpdateCol1, { bold: true }), cell("The AddIn ID for relay control from your MyGeotab setup", configUpdateCol2)] }),
    new TableRow({ children: [cell("notifyEmail", configUpdateCol1, { bold: true }), cell("Your email address for notifications", configUpdateCol2)] }),
  ]
}));
children.push(spacer());
children.push(numberedItem("Click Save (Ctrl+S) to save the script", "step2"));
children.push(spacer());

// Step 3
children.push(heading2("Step 3: Authorize the Script"));
children.push(para("Google requires you to grant permissions before the script can access Google Sheets, make HTTP requests, and send emails."));
children.push(numberedItem('In the Apps Script editor, find the function dropdown (next to the Run button) and select "testDisableRelay"', "step3"));
children.push(numberedItem('Click the "Run" button (play icon)', "step3"));
children.push(numberedItem('A dialog will appear: "Authorization required" - Click "Review permissions"', "step3"));
children.push(numberedItem("Select your Google account (sonyadam273@gmail.com)", "step3"));
children.push(numberedItem('If you see "Google hasn\'t verified this app", click "Advanced" at the bottom-left', "step3"));
children.push(numberedItem('Click "Go to Relay Speed Limiter (unsafe)" - this is safe because you created the script yourself', "step3"));
children.push(numberedItem("Review the permissions requested:", "step3"));
children.push(bulletItem("Google Sheets - to read/write log data and saved state"));
children.push(bulletItem("External URL Fetch - to call the MyGeotab API"));
children.push(bulletItem("Send Email - to send notification emails"));
children.push(numberedItem('Click "Allow" to grant all permissions', "step3"));
children.push(numberedItem("The test function will now run. Check the Execution Log at the bottom for results.", "step3"));
children.push(spacer());
children.push(importantNote('The test function uses a mock DeviceId "bB7". If there are real AddInData entries for this device, they WILL be removed. Use a test device ID or be prepared to restore data.'));
children.push(spacer());

// Step 4
children.push(heading2("Step 4: Deploy as Web App"));
children.push(para("Deploy the script as a web app so MyGeotab can send webhook requests to it."));
children.push(numberedItem('In the Apps Script editor, click "Deploy" in the top menu bar', "step4"));
children.push(numberedItem('Select "New deployment" from the dropdown', "step4"));
children.push(numberedItem('Click the gear icon next to "Select type" and choose "Web app"', "step4"));
children.push(numberedItem('Set the Description to "Relay Speed Limiter v1" (or any descriptive name)', "step4"));
children.push(numberedItem('Set "Execute as" to "Me" (your Google account email)', "step4"));
children.push(numberedItem('Set "Who has access" to "Anyone" - this allows MyGeotab to send POST requests without Google authentication', "step4"));
children.push(numberedItem('Click "Deploy"', "step4"));
children.push(numberedItem("A deployment confirmation will appear with the Web app URL. Copy this URL - this is your webhook endpoint.", "step4"));
children.push(spacer());
children.push(para("The URL will look like:", { italic: true }));
children.push(new Paragraph({
  spacing: { after: 120 },
  indent: { left: 360 },
  children: [new TextRun({ text: "https://script.google.com/macros/s/[DEPLOYMENT_ID]/exec", font: "Consolas", size: 16 })]
}));
children.push(importantNote("Save this URL! You will need it for Steps 7 and 8 when configuring webhook notifications."));
children.push(new Paragraph({ children: [new PageBreak()] }));

// Step 5
children.push(heading2('Step 5: Create MyGeotab Rule - "Speed Over 100"'));
children.push(para("This rule triggers when any vehicle in the organization exceeds 100 km/h."));
children.push(spacer());
children.push(heading3("Option A: Using MyGeotab UI"));
children.push(numberedItem("Log in to MyGeotab (my.geotab.com) with the alamui_indonesia database", "step5a"));
children.push(numberedItem('Navigate to Rules & Groups > Rules', "step5a"));
children.push(numberedItem('Click "Add" to create a new exception rule', "step5a"));
children.push(numberedItem('Set the Name to "Speed Over 100"', "step5a"));
children.push(numberedItem('Set the Comment to "Triggers when vehicle speed exceeds 100 km/h. Sends web request to disable relay."', "step5a"));
children.push(numberedItem("Set the Condition: Speed is more than 100 km/h", "step5a"));
children.push(numberedItem("Apply to: Entire Organization (GroupCompanyId) - this applies to all vehicles", "step5a"));
children.push(numberedItem("Set Active period: 2020-01-01 to 2050-01-01", "step5a"));
children.push(numberedItem("Set Color to Red", "step5a"));
children.push(numberedItem("Save the rule and note the Rule ID from the URL or API response", "step5a"));
children.push(spacer());
children.push(heading3("Option B: Using MyGeotab API"));
children.push(para("Use the Add method with typeName \"Rule\" and the following JSON payload:"));
children.push(spacer());
const ruleOverJson = `{
  "name": "Speed Over 100",
  "comment": "Triggers when vehicle speed exceeds 100 km/h...",
  "baseType": "Custom",
  "activeFrom": "2020-01-01T00:00:00.000Z",
  "activeTo": "2050-01-01T00:00:00.000Z",
  "groups": [{"id": "GroupCompanyId"}],
  "condition": {
    "conditionType": "IsValueMoreThan",
    "children": [{"conditionType": "Speed", "value": 0}],
    "value": 100
  },
  "color": {"r": 220, "g": 53, "b": 69, "a": 255},
  "state": "ExceptionRuleStateActiveId"
}`;
children.push(...codeBlock(ruleOverJson));
children.push(spacer());

// Step 6
children.push(heading2('Step 6: Create MyGeotab Rule - "Speed Under 100"'));
children.push(para('Follow the same process as Step 5, but with these differences:'));
children.push(bulletItem('Name: "Speed Under 100"'));
children.push(bulletItem('Comment: "Triggers when vehicle speed drops below 100 km/h. Sends web request to enable relay."'));
children.push(bulletItem("Condition: Speed is less than 100 km/h (conditionType: IsValueLessThan)"));
children.push(bulletItem("Color: Green (R:40, G:167, B:69)"));
children.push(para("See Section 10 for the complete JSON definition."));
children.push(spacer());

// Step 7
children.push(heading2("Step 7: Create Web Request Notification - Disable Relay"));
children.push(para('Link the "Speed Over 100" rule to a web request that calls the Apps Script with ?action=disable.'));
children.push(spacer());
children.push(heading3("Option A: Using MyGeotab UI"));
children.push(numberedItem('Go to Rules & Groups > Rules', "step7a"));
children.push(numberedItem('Open the "Speed Over 100" rule', "step7a"));
children.push(numberedItem("Under Notifications section, add a new notification", "step7a"));
children.push(numberedItem('Select Type: "Web Request"', "step7a"));
children.push(numberedItem("Paste the URL: [Your Web App URL from Step 4]?action=disable", "step7a"));
children.push(numberedItem("Set Method: POST", "step7a"));
children.push(numberedItem("Save the rule", "step7a"));
children.push(spacer());
children.push(heading3("Option B: Using MyGeotab API"));
children.push(para('Add a DistributionList with recipientType "WebRequest":'));
children.push(spacer());
const distListJson = `{
  "recipientType": "WebRequest",
  "rule": {"id": "[SPEED_OVER_100_RULE_ID]"},
  "recipient": {
    "address": "[WEB_APP_URL]?action=disable",
    "method": "POST"
  }
}`;
children.push(...codeBlock(distListJson));
children.push(spacer());

// Step 8
children.push(heading2("Step 8: Create Web Request Notification - Enable Relay"));
children.push(para('Follow the same process as Step 7, but for the "Speed Under 100" rule:'));
children.push(bulletItem('Open the "Speed Under 100" rule'));
children.push(bulletItem("Add a Web Request notification"));
children.push(bulletItem("URL: [Your Web App URL from Step 4]?action=enable"));
children.push(bulletItem("Method: POST"));
children.push(spacer());

// Step 9
children.push(heading2("Step 9: Verify the Complete Setup"));
children.push(para("After completing all steps, verify the system is working:"));
children.push(spacer());
children.push(numberedItem('Open the Google Sheet - confirm "Relay Control Log" and "Saved State" tabs exist', "step9"));
children.push(numberedItem("In the Apps Script editor, run testDisableRelay - should log a Processing entry and then Success/Failure result", "step9"));
children.push(numberedItem("Run testEnableRelay - should restore saved state and log the event", "step9"));
children.push(numberedItem('Open MyGeotab > Rules & Groups > Rules - both "Speed Over 100" and "Speed Under 100" rules should be Active', "step9"));
children.push(numberedItem("Check your email inbox (and spam folder) - you should have received notification emails from the test runs", "step9"));
children.push(numberedItem("Open the Web App URL in a browser - should return a JSON health check response with status: ok", "step9"));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 7. GOOGLE SHEETS STRUCTURE ====================
children.push(heading1("7. Google Sheets Structure"));
children.push(para("The Google Sheet has two tabs that the script creates and manages automatically."));
children.push(spacer());

children.push(heading3('Tab 1: "Relay Control Log"'));
children.push(para("This tab is the audit trail for all relay control events."));
const logCol1 = 2340, logCol2 = 7020;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [logCol1, logCol2],
  rows: [
    new TableRow({ children: [headerCell("Column", logCol1), headerCell("Description", logCol2)] }),
    new TableRow({ children: [cell("Timestamp", logCol1, { bold: true }), cell("ISO 8601 timestamp of the event (e.g., 2026-03-27T10:30:00.000Z)", logCol2)] }),
    new TableRow({ children: [cell("Device ID", logCol1, { bold: true }), cell("MyGeotab device identifier (e.g., bB7)", logCol2)] }),
    new TableRow({ children: [cell("Device Name", logCol1, { bold: true }), cell("Human-readable vehicle name", logCol2)] }),
    new TableRow({ children: [cell("Driver", logCol1, { bold: true }), cell("Driver name (if available from the rule notification)", logCol2)] }),
    new TableRow({ children: [cell("Rule Name", logCol1, { bold: true }), cell('Name of the rule that triggered (e.g., "Speed Over 100")', logCol2)] }),
    new TableRow({ children: [cell("Action", logCol1, { bold: true }), cell('"disable" or "enable"', logCol2)] }),
    new TableRow({ children: [cell("Speed (km/h)", logCol1, { bold: true }), cell("Current vehicle speed at time of event", logCol2)] }),
    new TableRow({ children: [cell("Status", logCol1, { bold: true }), cell("Processing... / Success - disable / Success - enable / Failed / Error", logCol2)] }),
  ]
}));
children.push(spacer());

children.push(heading3('Tab 2: "Saved State"'));
children.push(para("This tab temporarily stores removed AddInData so it can be restored when the vehicle slows down. Rows are automatically deleted after successful restoration."));
const stateCol1 = 2340, stateCol2 = 7020;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [stateCol1, stateCol2],
  rows: [
    new TableRow({ children: [headerCell("Column", stateCol1), headerCell("Description", stateCol2)] }),
    new TableRow({ children: [cell("Device ID", stateCol1, { bold: true }), cell("MyGeotab device identifier", stateCol2)] }),
    new TableRow({ children: [cell("AddInData ID", stateCol1, { bold: true }), cell("ID of the removed AddInData entry", stateCol2)] }),
    new TableRow({ children: [cell("AddInData JSON", stateCol1, { bold: true }), cell("Full JSON of the AddInData object (used for restoration via API Add)", stateCol2)] }),
    new TableRow({ children: [cell("Saved At", stateCol1, { bold: true }), cell("ISO 8601 timestamp when the state was saved", stateCol2)] }),
  ]
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 8. TECHNICAL DETAILS ====================
children.push(heading1("8. How It Works - Technical Details"));

children.push(heading3("Authentication"));
children.push(para("The script uses username/password authentication via the MyGeotab BasicAuthentication service account. It authenticates fresh on each invocation by calling the Authenticate API method. There is no token management, session caching, or bearer token expiry to worry about. The credentials are stored in the CONFIG object at the top of the script."));
children.push(spacer());

children.push(heading3("Action Detection"));
children.push(para("The script supports two methods to determine the action (disable or enable):"));
children.push(spacer());
children.push(para("Method 1: Query Parameter (Primary)", { bold: true }));
children.push(para("The webhook URL includes ?action=disable or ?action=enable as a query parameter. The script reads this from e.parameter.action. This is the most reliable method and is used by the configured notifications."));
children.push(spacer());
children.push(para("Method 2: Rule Name Fallback", { bold: true }));
children.push(para('If no query parameter is provided, the script inspects the RuleName field from the POST data. It looks for keywords: "over", "exceed", "above", or "lebih" (Indonesian) to determine disable action. Keywords "under", "below", "normal", or "kurang" (Indonesian) determine enable action.'));
children.push(spacer());

children.push(heading3("Disable Flow (Remove AddInData)"));
children.push(numberedItem("Check if the device already has saved state (relay already disabled) - if yes, skip", "techDisable"));
children.push(numberedItem("Query MyGeotab API: Get AddInData filtered by addInId and deviceId", "techDisable"));
children.push(numberedItem("If filtered search fails, fall back to broader search and filter results manually by device", "techDisable"));
children.push(numberedItem('Save the full AddInData JSON to the "Saved State" Google Sheet tab', "techDisable"));
children.push(numberedItem("Call MyGeotab API Remove method with the AddInData entity ID", "techDisable"));
children.push(numberedItem("Log the event and send email notification", "techDisable"));
children.push(spacer());

children.push(heading3("Enable Flow (Restore AddInData)"));
children.push(numberedItem('Read all saved state rows from the "Saved State" sheet for the device', "techEnable"));
children.push(numberedItem("Parse the saved AddInData JSON", "techEnable"));
children.push(numberedItem("Build an Add entity with the original addInId and details, updating the date to current time", "techEnable"));
children.push(numberedItem("Call MyGeotab API Add method with typeName AddInData", "techEnable"));
children.push(numberedItem("Delete the saved state row from the Google Sheet", "techEnable"));
children.push(numberedItem("Log the event and send email notification", "techEnable"));
children.push(spacer());

children.push(heading3("Rate Limiting"));
children.push(para("The MyGeotab API enforces a rate limit of approximately 10 calls per minute per user. The script handles API rate limit errors (OverLimit) with try/catch blocks. If you need to process many devices in batch, add Utilities.sleep(7000) between API calls and implement retry logic with a 61-second wait on OverLimit errors."));
children.push(spacer());

children.push(heading3("AddInData and Relay Control"));
children.push(para('The addInId "aS_lt7cUYYEutQoXZGoQPZq" corresponds to the NFC Key Authorization add-in on the alamui_indonesia database. Each AddInData entry contains authorization details (vehicle, user, driverKey, etc.). Removing this data effectively disables the relay authorization for that vehicle, which activates the speed limiter. Restoring the data re-enables the relay.'));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 9. COMPLETE SOURCE CODE ====================
children.push(heading1("9. Complete Source Code - relay_speed_limiter.js"));
children.push(para("File location: C:\\Users\\sonyadam\\Desktop\\fuel-usage-addin\\relay_speed_limiter.js"));
children.push(para("Total lines: 620"));
children.push(spacer());

// Break source code into lines and add as code paragraphs
const sourceLines = sourceCode.split("\n");
sourceLines.forEach(line => {
  children.push(new Paragraph({
    spacing: { after: 0, line: 220 },
    indent: { left: 240 },
    children: [new TextRun({ text: line || " ", font: "Consolas", size: 14, color: "333333" })]
  }));
});
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 10. RULE JSON DEFINITIONS ====================
children.push(heading1("10. MyGeotab Rule JSON Definitions"));

children.push(heading3("rule_speed_over_100.json"));
children.push(para("File: C:\\Users\\sonyadam\\Desktop\\fuel-usage-addin\\rule_speed_over_100.json"));
children.push(spacer());
const ruleOver = `{
  "name": "Speed Over 100",
  "comment": "Triggers when vehicle speed exceeds 100 km/h. Sends web request to disable relay.",
  "baseType": "Custom",
  "activeFrom": "2020-01-01T00:00:00.000Z",
  "activeTo": "2050-01-01T00:00:00.000Z",
  "groups": [
    {
      "id": "GroupCompanyId"
    }
  ],
  "condition": {
    "conditionType": "IsValueMoreThan",
    "children": [
      {
        "conditionType": "Speed",
        "value": 0
      }
    ],
    "value": 100
  },
  "color": {
    "r": 220,
    "g": 53,
    "b": 69,
    "a": 255
  },
  "state": "ExceptionRuleStateActiveId"
}`;
children.push(...codeBlock(ruleOver));
children.push(spacer());

children.push(heading3("rule_speed_under_100.json"));
children.push(para("File: C:\\Users\\sonyadam\\Desktop\\fuel-usage-addin\\rule_speed_under_100.json"));
children.push(spacer());
const ruleUnder = `{
  "name": "Speed Under 100",
  "comment": "Triggers when vehicle speed drops below 100 km/h. Sends web request to enable relay.",
  "baseType": "Custom",
  "activeFrom": "2020-01-01T00:00:00.000Z",
  "activeTo": "2050-01-01T00:00:00.000Z",
  "groups": [
    {
      "id": "GroupCompanyId"
    }
  ],
  "condition": {
    "conditionType": "IsValueLessThan",
    "children": [
      {
        "conditionType": "Speed",
        "value": 0
      }
    ],
    "value": 100
  },
  "color": {
    "r": 40,
    "g": 167,
    "b": 69,
    "a": 255
  },
  "state": "ExceptionRuleStateActiveId"
}`;
children.push(...codeBlock(ruleUnder));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 11. TROUBLESHOOTING ====================
children.push(heading1("11. Troubleshooting"));
children.push(para("Common issues and their solutions:"));
children.push(spacer());

const tsCol1 = 3000, tsCol2 = 6360;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [tsCol1, tsCol2],
  rows: [
    new TableRow({ children: [headerCell("Issue", tsCol1), headerCell("Solution", tsCol2)] }),
    new TableRow({ children: [
      cell('"Auth failed" error', tsCol1, { bold: true }),
      cell("Check service account credentials in CONFIG. Verify the account exists on the MyGeotab database and has API access permissions.", tsCol2)
    ]}),
    new TableRow({ children: [
      cell('"No AddInData found for device"', tsCol1, { bold: true }),
      cell("The device may not have relay data configured, or the addInId in CONFIG may be wrong. Verify the addInId matches your add-in setup.", tsCol2)
    ]}),
    new TableRow({ children: [
      cell('"OverLimit" API error', tsCol1, { bold: true }),
      cell("API rate limit exceeded (10 calls/minute). Wait 60 seconds before retrying. If doing batch operations, add Utilities.sleep(7000) between calls.", tsCol2)
    ]}),
    new TableRow({ children: [
      cell("Web app returns 403", tsCol1, { bold: true }),
      cell('Redeploy the web app with "Who has access" set to "Anyone". Also verify the Google account permissions.', tsCol2)
    ]}),
    new TableRow({ children: [
      cell("Rule not triggering", tsCol1, { bold: true }),
      cell("Verify the rule is Active in MyGeotab. Check the condition threshold (100 km/h). Ensure the vehicle is in the correct group (GroupCompanyId).", tsCol2)
    ]}),
    new TableRow({ children: [
      cell("Email not received", tsCol1, { bold: true }),
      cell("Check CONFIG.sendEmail is true. Verify the email address in CONFIG.notifyEmail. Check your spam/junk folder. Google Apps Script has daily email quotas.", tsCol2)
    ]}),
    new TableRow({ children: [
      cell("AddInData accidentally removed", tsCol1, { bold: true }),
      cell('Check the "Saved State" sheet tab for saved entries. Run the enableRelay function to restore them. If saved states were also lost, you may need to manually re-add AddInData via the MyGeotab API.', tsCol2)
    ]}),
    new TableRow({ children: [
      cell("Webhook not reaching Apps Script", tsCol1, { bold: true }),
      cell("Verify the webhook URL is correct in the DistributionList. Test the URL by opening it in a browser (should return JSON health check). Check that the MyGeotab rule notification is properly linked.", tsCol2)
    ]}),
    new TableRow({ children: [
      cell("Google Sheet errors", tsCol1, { bold: true }),
      cell("Verify the sheetId in CONFIG matches your Google Sheet. Ensure the Google account has edit access to the sheet. Check that the sheet has not been deleted.", tsCol2)
    ]}),
  ]
}));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ==================== 12. FILE LOCATIONS ====================
children.push(heading1("12. File Locations"));
children.push(para("All project files and their locations:"));
children.push(spacer());

const fCol1 = 2600, fCol2 = 4160, fCol3 = 2600;
children.push(new Table({
  width: { size: TABLE_WIDTH, type: WidthType.DXA },
  columnWidths: [fCol1, fCol2, fCol3],
  rows: [
    new TableRow({ children: [
      headerCell("File", fCol1),
      headerCell("Path", fCol2),
      headerCell("Description", fCol3)
    ]}),
    new TableRow({ children: [
      cell("relay_speed_limiter.js", fCol1, { bold: true }),
      cell("C:\\Users\\sonyadam\\Desktop\\fuel-usage-addin\\", fCol2),
      cell("Main Google Apps Script source code (620 lines)", fCol3)
    ]}),
    new TableRow({ children: [
      cell("rule_speed_over_100.json", fCol1, { bold: true }),
      cell("C:\\Users\\sonyadam\\Desktop\\fuel-usage-addin\\", fCol2),
      cell("MyGeotab rule definition for speed > 100", fCol3)
    ]}),
    new TableRow({ children: [
      cell("rule_speed_under_100.json", fCol1, { bold: true }),
      cell("C:\\Users\\sonyadam\\Desktop\\fuel-usage-addin\\", fCol2),
      cell("MyGeotab rule definition for speed < 100", fCol3)
    ]}),
    new TableRow({ children: [
      cell("This guide (.docx)", fCol1, { bold: true }),
      cell("C:\\Users\\sonyadam\\Desktop\\fuel-usage-addin\\", fCol2),
      cell("Complete setup documentation", fCol3)
    ]}),
  ]
}));
children.push(spacer());
children.push(spacer());
children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { before: 400 },
  children: [new TextRun({ text: "--- End of Document ---", font: "Arial", size: 20, color: "888888", italics: true })]
}));

// ==================== BUILD DOCUMENT ====================

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 20 } }
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 }
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "404040" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 }
      },
    ]
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      { reference: "steps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "step1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "step2", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "step3", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "step4", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "step5a", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "step7a", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "step9", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "disableSteps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "enableSteps", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "techDisable", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "techEnable", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "MyGeotab Relay Speed Limiter - Setup Guide", font: "Arial", size: 16, color: "888888", italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", font: "Arial", size: 16, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "888888" }),
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("C:/Users/sonyadam/Desktop/fuel-usage-addin/Relay_Speed_Limiter_Setup_Guide.docx", buffer);
  console.log("Document created successfully!");
  console.log("Path: C:/Users/sonyadam/Desktop/fuel-usage-addin/Relay_Speed_Limiter_Setup_Guide.docx");
});
