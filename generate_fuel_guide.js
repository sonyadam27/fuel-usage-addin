const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require("docx");

// Read source files
const mainJs = fs.readFileSync("C:/Users/sonyadam/Desktop/fuel-usage-addin/js/main.js", "utf-8");
const indexHtml = fs.readFileSync("C:/Users/sonyadam/Desktop/fuel-usage-addin/index.html", "utf-8");
const stylesCss = fs.readFileSync("C:/Users/sonyadam/Desktop/fuel-usage-addin/css/styles.css", "utf-8");
const configJson = fs.readFileSync("C:/Users/sonyadam/Desktop/fuel-usage-addin/config.json", "utf-8");
const configInlineJson = fs.readFileSync("C:/Users/sonyadam/Desktop/fuel-usage-addin/config-inline.json", "utf-8");

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

function para(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, font: "Arial", size: 20 })]
  });
}

function boldPara(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, bold: true, font: "Arial", size: 20 })]
  });
}

function codePara(text) {
  return new Paragraph({
    spacing: { after: 40 },
    children: [new TextRun({ text, font: "Consolas", size: 16 })]
  });
}

function codeBlock(code) {
  const lines = code.split("\n");
  return lines.map(line => codePara(line));
}

function bulletItem(text, ref = "bullets", level = 0) {
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20 })]
  });
}

function numberedItem(text, ref = "numbers", level = 0) {
  return new Paragraph({
    numbering: { reference: ref, level },
    spacing: { after: 60 },
    children: [new TextRun({ text, font: "Arial", size: 20 })]
  });
}

function importantNote(text) {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    children: [
      new TextRun({ text: "IMPORTANT: ", bold: true, font: "Arial", size: 20, color: "C00000" }),
      new TextRun({ text, font: "Arial", size: 20 })
    ]
  });
}

// ==================== BUILD DOCUMENT ====================

const c1 = 3200;
const c2 = TABLE_WIDTH - c1;

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers2",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers3",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers4",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers5",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      },
      {
        reference: "numbers6",
        levels: [{
          level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
      },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 }
      },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial" },
        paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 2 }
      }
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
          children: [new TextRun({ text: "MyGeotab Fuel Usage Add-In \u2014 Setup Guide", font: "Arial", size: 16, color: "888888", italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", font: "Arial", size: 16, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "888888" })
          ]
        })]
      })
    },
    children: [

      // ==================== TITLE ====================
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [new TextRun({ text: "MyGeotab Fuel Usage Per Day Add-In", bold: true, font: "Arial", size: 44, color: "1F4E79" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Complete Setup Guide", font: "Arial", size: 28, color: "2E75B6" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "Works on any database (auto-detected)", font: "Arial", size: 22, color: "666666" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: "Generated: " + new Date().toISOString().split("T")[0], font: "Arial", size: 18, color: "999999" })]
      }),

      // ==================== 1. OVERVIEW ====================
      heading1("1. Overview"),
      para("The Fuel Usage Per Day add-in is a custom MyGeotab add-in that provides daily fuel consumption reporting for all vehicles in a fleet. It works on any MyGeotab database \u2014 the database name is automatically detected from the URL (e.g., sasa_inti, Travl, etc.) and displayed in the UI."),
      heading3("Key Features"),
      bulletItem("Daily fuel usage breakdown per vehicle"),
      bulletItem("Summary dashboard with total vehicles, total fuel used, average fuel per vehicle per day, and total idle fuel"),
      bulletItem("Date range picker (defaults to last 7 days)"),
      bulletItem("Distinguishes between driving fuel and idle fuel consumption"),
      bulletItem("Export to CSV functionality"),
      bulletItem("Auto-loads data on page focus"),
      bulletItem("Responsive design with professional styling"),

      heading3("MyGeotab API Objects Used"),
      bulletItem("Device \u2014 loads all vehicles and caches device names"),
      bulletItem("FuelUsed \u2014 retrieves fuel consumption records by date range"),

      // ==================== 2. ARCHITECTURE ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("2. System Architecture"),

      heading3("Deployment Model"),
      para("The add-in uses two deployment approaches:"),
      boldPara("Option A: External URL (GitHub Pages)"),
      bulletItem("HTML/CSS/JS hosted on GitHub Pages at https://sonyadam27.github.io/fuel-usage-addin/index.html"),
      bulletItem("config.json references the external URL"),
      bulletItem("Requires \"Allow unsigned Add-Ins\" to be enabled in MyGeotab System Settings"),

      boldPara("Option B: Inline (Self-contained)"),
      bulletItem("All HTML, JS, and CSS embedded directly in config-inline.json"),
      bulletItem("No external hosting needed"),
      bulletItem("The \"files\" object in the config contains customPage.html, customPage.js, and customPage.css"),

      heading3("Component Flow"),
      numberedItem("User opens the \"Fuel Usage Per Day\" page in MyGeotab sidebar (under Fuel and Energy or Add-Ins)", "numbers"),
      numberedItem("The add-in initializes: sets default date range (last 7 days), binds button events", "numbers"),
      numberedItem("On Load Data click (or auto on focus): calls MyGeotab API Get for Device to cache vehicle names", "numbers"),
      numberedItem("Then calls MyGeotab API Get for FuelUsed with the selected date range", "numbers"),
      numberedItem("Groups fuel records by device + date, calculates totals", "numbers"),
      numberedItem("Renders summary cards and data table", "numbers"),
      numberedItem("User can click Export CSV to download the table as a .csv file", "numbers"),

      // ==================== 3. CONFIGURATION ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("3. Configuration Reference"),

      heading2("3.1 config.json (External URL)"),
      para("Used when hosting the add-in on GitHub Pages or another external server."),
      new Table({
        width: { size: TABLE_WIDTH, type: WidthType.DXA },
        columnWidths: [c1, c2],
        rows: [
          new TableRow({ children: [headerCell("Parameter", c1), headerCell("Value", c2)] }),
          new TableRow({ children: [cell("Name", c1, { bold: true }), cell("Fuel Usage Per Day", c2)] }),
          new TableRow({ children: [cell("Version", c1, { bold: true }), cell("1.0.0", c2)] }),
          new TableRow({ children: [cell("Support Email", c1, { bold: true }), cell("sonyadam@geotab.com", c2)] }),
          new TableRow({ children: [cell("isSigned", c1, { bold: true }), cell("false", c2)] }),
          new TableRow({ children: [cell("URL", c1, { bold: true }), cell("https://sonyadam27.github.io/fuel-usage-addin/index.html", c2)] }),
          new TableRow({ children: [cell("Menu Path", c1, { bold: true }), cell("EngineMaintenanceLink/", c2)] }),
          new TableRow({ children: [cell("Menu Name (EN)", c1, { bold: true }), cell("Fuel Usage Per Day", c2)] }),
          new TableRow({ children: [cell("Icon", c1, { bold: true }), cell("https://cdn-icons-png.flaticon.com/512/2933/2933054.png", c2)] }),
        ]
      }),

      heading2("3.2 config-inline.json (Self-contained)"),
      para("Used for inline deployment where all code is embedded in the config JSON. The \"files\" object contains:"),
      new Table({
        width: { size: TABLE_WIDTH, type: WidthType.DXA },
        columnWidths: [c1, c2],
        rows: [
          new TableRow({ children: [headerCell("File", c1), headerCell("Description", c2)] }),
          new TableRow({ children: [cell("customPage.html", c1, { bold: true }), cell("Full HTML with inline CSS for the add-in UI", c2)] }),
          new TableRow({ children: [cell("customPage.js", c1, { bold: true }), cell("Minified JavaScript add-in logic (geotab.addin.fuelUsagePerDay)", c2)] }),
          new TableRow({ children: [cell("customPage.css", c1, { bold: true }), cell("Empty (CSS is inlined in the HTML)", c2)] }),
        ]
      }),

      // ==================== 4. UI COMPONENTS ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("4. User Interface"),

      heading2("4.1 Controls"),
      bulletItem("From Date picker \u2014 defaults to 7 days ago"),
      bulletItem("To Date picker \u2014 defaults to today"),
      bulletItem("Load Data button (blue) \u2014 fetches fuel data for the selected range"),
      bulletItem("Export CSV button (green) \u2014 downloads the current table as CSV"),

      heading2("4.2 Summary Cards"),
      para("Four summary cards appear after data loads:"),
      new Table({
        width: { size: TABLE_WIDTH, type: WidthType.DXA },
        columnWidths: [3120, 6240],
        rows: [
          new TableRow({ children: [headerCell("Card", 3120), headerCell("Description", 6240)] }),
          new TableRow({ children: [cell("Total Vehicles", 3120, { bold: true }), cell("Count of unique vehicles with fuel data in the range", 6240)] }),
          new TableRow({ children: [cell("Total Fuel Used (L)", 3120, { bold: true }), cell("Sum of all fuel consumed across all vehicles", 6240)] }),
          new TableRow({ children: [cell("Avg Fuel/Vehicle/Day (L)", 3120, { bold: true }), cell("Total fuel divided by number of vehicles divided by number of days", 6240)] }),
          new TableRow({ children: [cell("Total Idle Fuel (L)", 3120, { bold: true }), cell("Sum of fuel consumed while idling", 6240)] }),
        ]
      }),

      heading2("4.3 Data Table"),
      para("The results table shows one row per vehicle per day with these columns:"),
      new Table({
        width: { size: TABLE_WIDTH, type: WidthType.DXA },
        columnWidths: [3120, 6240],
        rows: [
          new TableRow({ children: [headerCell("Column", 3120), headerCell("Description", 6240)] }),
          new TableRow({ children: [cell("Vehicle Name", 3120, { bold: true }), cell("Device name from MyGeotab (or device ID if name unavailable)", 6240)] }),
          new TableRow({ children: [cell("Date", 3120, { bold: true }), cell("Date in YYYY-MM-DD format", 6240)] }),
          new TableRow({ children: [cell("Total Fuel Used (L)", 3120, { bold: true }), cell("Total fuel consumed that day (driving + idle)", 6240)] }),
          new TableRow({ children: [cell("Idle Fuel Used (L)", 3120, { bold: true }), cell("Fuel consumed while vehicle was idling", 6240)] }),
          new TableRow({ children: [cell("Driving Fuel Used (L)", 3120, { bold: true }), cell("Calculated: Total Fuel - Idle Fuel", 6240)] }),
        ]
      }),

      // ==================== 5. SETUP GUIDE ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("5. Step-by-Step Setup Guide"),

      heading2("Step 1: Enable Unsigned Add-Ins in MyGeotab"),
      numberedItem("Log in to MyGeotab (my.geotab.com) with your target database", "numbers2"),
      numberedItem("Go to Administration > System Settings > Add-Ins", "numbers2"),
      numberedItem("Toggle \"Allow unsigned Add-Ins\" to ON", "numbers2"),
      numberedItem("Click Save", "numbers2"),

      heading2("Step 2: Install the Add-In (Inline Method)"),
      numberedItem("In the same System Settings > Add-Ins page, click \"New Add-In\"", "numbers3"),
      numberedItem("A configuration text area will appear", "numbers3"),
      numberedItem("Copy the entire contents of config-inline.json", "numbers3"),
      numberedItem("Paste it into the configuration text area", "numbers3"),
      numberedItem("Click Save", "numbers3"),
      numberedItem("The add-in will appear under the sidebar menu (Fuel and Energy or the configured menu path)", "numbers3"),

      heading2("Step 2 (Alternative): Install via External URL"),
      numberedItem("Push the index.html, css/styles.css, and js/main.js to a GitHub repository", "numbers4"),
      numberedItem("Enable GitHub Pages for the repository (Settings > Pages > Source: main branch)", "numbers4"),
      numberedItem("Note the published URL (e.g., https://sonyadam27.github.io/fuel-usage-addin/index.html)", "numbers4"),
      numberedItem("In MyGeotab System Settings > Add-Ins, click \"New Add-In\"", "numbers4"),
      numberedItem("Paste the contents of config.json (which references the GitHub Pages URL)", "numbers4"),
      numberedItem("Click Save", "numbers4"),

      heading2("Step 3: Verify the Add-In"),
      numberedItem("Refresh MyGeotab", "numbers5"),
      numberedItem("Look for \"Fuel Usage Per Day\" in the sidebar menu", "numbers5"),
      numberedItem("Click on it to open the add-in page", "numbers5"),
      numberedItem("The default date range (last 7 days) should auto-load data", "numbers5"),
      numberedItem("Verify summary cards display correct totals", "numbers5"),
      numberedItem("Click Export CSV to confirm the download works", "numbers5"),

      // ==================== 6. HOW IT WORKS ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("6. How It Works \u2014 Technical Details"),

      heading2("6.1 Add-In Lifecycle"),
      para("The add-in follows the standard MyGeotab add-in lifecycle pattern:"),
      bulletItem("initialize(api, state, callback) \u2014 called once when the add-in first loads. Detects the database name from the URL, sets default dates, binds button events, calls callback() to signal ready."),
      bulletItem("focus(api, state) \u2014 called each time the user navigates to the add-in page. Auto-triggers loadFuelData() to refresh data."),
      bulletItem("blur() \u2014 called when the user navigates away. Currently a no-op."),

      heading2("6.2 Dynamic Database Detection"),
      para("The add-in automatically detects the current database name from the MyGeotab URL path using the pattern:"),
      codePara("window.location.pathname.match(/\\/([^\\/]+)\\//)"),
      para("This extracts the database name (e.g., \"Travl\", \"sasa_inti\", \"alamui_indonesia\") from URLs like https://my.geotab.com/Travl/#addin-... and displays it in the page subtitle. The same detection is used for the CSV export filename, producing files like fuel_usage_per_day_Travl.csv."),

      heading2("6.3 Data Flow"),
      numberedItem("loadFuelData() validates date inputs and shows loading spinner", "numbers6"),
      numberedItem("loadDevices() calls api.call(\"Get\", {typeName: \"Device\"}) to cache all vehicle names by ID", "numbers6"),
      numberedItem("api.call(\"Get\", {typeName: \"FuelUsed\", search: {fromDate, toDate}}) retrieves fuel records", "numbers6"),
      numberedItem("processFuelData() groups records by deviceId + date, summing totalFuelUsed and totalIdlingFuelUsedL", "numbers6"),
      numberedItem("Calculates driving fuel as totalFuel - idleFuel", "numbers6"),
      numberedItem("Computes summary statistics (unique vehicles, totals, averages)", "numbers6"),
      numberedItem("renderTable() builds HTML table rows and displays them", "numbers6"),

      heading2("6.4 CSV Export"),
      para("The exportCsv() function iterates through the rendered HTML table, extracts all cell text content, escapes double quotes, wraps each value in quotes, and downloads as fuel_usage_per_day_[database].csv (database name auto-detected from URL)."),

      // ==================== 7. TROUBLESHOOTING ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("7. Troubleshooting"),

      new Table({
        width: { size: TABLE_WIDTH, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({ children: [headerCell("Issue", 4680), headerCell("Solution", 4680)] }),
          new TableRow({ children: [cell("Add-in not appearing in sidebar", 4680, { bold: true }), cell("Verify \"Allow unsigned Add-Ins\" is enabled. Check the menu path in config (EngineMaintenanceLink/). Refresh the page.", 4680)] }),
          new TableRow({ children: [cell("\"No fuel data found\" error", 4680, { bold: true }), cell("Expand the date range. Verify vehicles have GO devices that report fuel data. Not all device types support FuelUsed.", 4680)] }),
          new TableRow({ children: [cell("\"Failed to load vehicles\" error", 4680, { bold: true }), cell("Check your MyGeotab permissions. The user must have access to view devices.", 4680)] }),
          new TableRow({ children: [cell("Blank page or script error", 4680, { bold: true }), cell("Open browser dev tools (F12). Check for JavaScript errors. Verify the add-in config JSON is valid.", 4680)] }),
          new TableRow({ children: [cell("External URL not loading", 4680, { bold: true }), cell("Verify GitHub Pages is enabled and the URL is accessible. Check for CORS issues. Try the inline config method instead.", 4680)] }),
          new TableRow({ children: [cell("CSV export not working", 4680, { bold: true }), cell("Ensure data is loaded first. Check browser popup blocker is not blocking the download.", 4680)] }),
          new TableRow({ children: [cell("Idle fuel showing 0", 4680, { bold: true }), cell("Not all devices report idle fuel separately. The totalIdlingFuelUsedL field may not be populated for some device types.", 4680)] }),
        ]
      }),

      // ==================== 8. FILE LOCATIONS ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("8. File Locations"),

      new Table({
        width: { size: TABLE_WIDTH, type: WidthType.DXA },
        columnWidths: [2800, 3760, 2800],
        rows: [
          new TableRow({ children: [headerCell("File", 2800), headerCell("Path", 3760), headerCell("Description", 2800)] }),
          new TableRow({ children: [cell("index.html", 2800, { bold: true }), cell("fuel-usage-addin/", 3760), cell("Main HTML page", 2800)] }),
          new TableRow({ children: [cell("main.js", 2800, { bold: true }), cell("fuel-usage-addin/js/", 3760), cell("Add-in JavaScript logic", 2800)] }),
          new TableRow({ children: [cell("styles.css", 2800, { bold: true }), cell("fuel-usage-addin/css/", 3760), cell("Stylesheet", 2800)] }),
          new TableRow({ children: [cell("config.json", 2800, { bold: true }), cell("fuel-usage-addin/", 3760), cell("External URL config", 2800)] }),
          new TableRow({ children: [cell("config-inline.json", 2800, { bold: true }), cell("fuel-usage-addin/", 3760), cell("Inline (self-contained) config", 2800)] }),
          new TableRow({ children: [cell("This guide", 2800, { bold: true }), cell("fuel-usage-addin/", 3760), cell("Setup documentation", 2800)] }),
        ]
      }),

      // ==================== 9. SOURCE CODE ====================
      new Paragraph({ children: [new PageBreak()] }),
      heading1("9. Complete Source Code"),

      heading2("9.1 config.json"),
      ...codeBlock(configJson.trim()),

      heading2("9.2 config-inline.json"),
      para("Note: The inline config is very large as it contains all HTML, CSS, and JS embedded. The key structure is shown below. See the actual file for the full embedded code."),
      ...codeBlock(JSON.stringify(JSON.parse(configInlineJson), (key, val) => {
        if (key === "customPage.html" || key === "customPage.js") return "[... embedded code ...]";
        return val;
      }, 2)),

      heading2("9.3 index.html"),
      ...codeBlock(indexHtml.trim()),

      heading2("9.4 js/main.js"),
      ...codeBlock(mainJs.trim()),

      heading2("9.5 css/styles.css"),
      ...codeBlock(stylesCss.trim()),
    ]
  }]
});

// ==================== GENERATE ====================
const outputPath = "C:/Users/sonyadam/Desktop/fuel-usage-addin/Fuel_Usage_AddIn_Setup_Guide.docx";

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log("Document created successfully!");
  console.log("Path: " + outputPath);
}).catch(err => {
  console.error("Error creating document:", err);
});
