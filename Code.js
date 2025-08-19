/**
 * (ALE) Data Collection Spreadsheet (Holmgreen)
 *
 * Objective: Create a data collection sheet that will be used as part of Holmgreen's MTSS. This project extends the original project named "(Master Copy) Data Collection Spreadsheet (Holmgreen)".
 *
 * Google Apps Script Development: Alvaro Gomez, Academic Technology Coach
 * Office: 210-397-9408
 */

/**
 * Creates a user menu in the Google Spreadsheet that this script is bounded to.
 * 
 * Note: This menu is commented out because the user does not need to create the weekly sheets. All of the sheets were created for them at the start of the school year. This is left here for reference.
 */
// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu("📈 Data Collection Information 📉")
//     .addItem("Create Weekly Sheets", "setupWeeklyBehaviorSheets")
//     .addItem("Help", "helpInformation")
//     .addToUi();
// }

/**
 * Duplicates the Week_Tempate sheet.
 * The user will be prompted to:
 *      1. enter the date when the data collection will start and
 *      2. enter the number of weekly sheets they want created.
 * The function will fill in the dates into each weekly sheet and create the
 * indicated number of weekly sheets.
 */
function setupWeeklyBehaviorSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const templateSheet = ss.getSheetByName("Week_Template");
  const summarySheet = ss.getSheetByName("Summary");

  const weekSheets = ss
    .getSheets()
    .filter((sheet) => /^Week \d+$/.test(sheet.getName()));
  let startingWeekNumber = 1;
  let colIndex = 2; // default staring column for summary formulas
  let clearSummary = true;

  if (weekSheets.length > 0) {
    const choice = ui.prompt(
      "Weekly Sheets Detected",
      "⚠️ This spreadsheet already has weekly sheets built in it.\n\nWhat do you want to do?\n\n" +
        "Type DELETE to remove all existing weekly sheets.\n" +
        "Type CONTINUE to keep them and start with the next available week number.\n" +
        "Type CANCEL to abort the setup.",
      ui.ButtonSet.OK_CANCEL,
    );

    if (choice.getSelectedButton() !== ui.Button.OK) return;

    const input = choice.getResponseText().trim().toUpperCase();

    if (input === "DELETE") {
      weekSheets.forEach((sheet) => ss.deleteSheet(sheet));
    } else if (input === "CONTINUE") {
      const weekNumbers = weekSheets
        .map((sheet) => parseInt(sheet.getName().match(/\d+/)[0]))
        .filter((num) => !isNaN(num));
      startingWeekNumber = Math.max(...weekNumbers) + 1;
      clearSummary = false;

      // Find the last used column in row 2, starting at B (col 2)
      const row2 = summarySheet
        .getRange(2, 2, 1, summarySheet.getMaxColumns() - 1)
        .getValues()[0];
      const lastColUsed = row2.reduceRight(
        (acc, val, idx) => (acc === null && val !== "" ? idx + 2 : acc),
        null,
      );
      colIndex = lastColUsed ? lastColUsed + 1 : 2;
    } else {
      ui.alert("Setup cancelled.");
      return;
    }
  }

  // Ask for start date
  const dateResponse = ui.prompt(
    "Start Date",
    "What day will you start collecting data? (MM/DD/YYYY)",
    ui.ButtonSet.OK_CANCEL,
  );
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;

  let startDate = new Date(dateResponse.getResponseText());
  if (isNaN(startDate)) {
    ui.alert("Invalid date. Please enter in MM/DD/YYYY format.");
    return;
  }

  // Backtrack to Monday if not already Monday
  const day = startDate.getDay(); // Sunday = 0, Monday = 1, ...
  if (day !== 1) {
    const diff = day === 0 ? 6 : day - 1; // go back to previous Monday
    startDate.setDate(startDate.getDate() - diff);
  }

  // Ask how many weeks
  const weekResponse = ui.prompt(
    "Number of Weeks",
    "How many weeks would you like to create?",
    ui.ButtonSet.OK_CANCEL,
  );
  if (weekResponse.getSelectedButton() !== ui.Button.OK) return;

  const numWeeks = parseInt(weekResponse.getResponseText());
  if (isNaN(numWeeks) || numWeeks <= 0) {
    ui.alert("Please enter a valid number of weeks.");
    return;
  }

  // Clear and setup summary sheet headers
  if (clearSummary) {
    summarySheet.getRange("A1:BD9").clearContent();
    const rowLabels = [
      "=Week_Template!N9",
      "=Week_Template!N10",
      "=Week_Template!N11",
      "=Week_Template!N12",
      "=Week_Template!N13",
      "=Week_Template!N14",
      "=Week_Template!N15",
    ];
    summarySheet.getRange("A3:A9").setValues(rowLabels.map((label) => [label]));
    colIndex = 2;
  }

  // Generate weekly sheets and collect formulas
  const sheetNames = [];
  let currentMonday = new Date(startDate);

  for (let i = 0; i < numWeeks; i++) {
    const sheetName = `Week ${startingWeekNumber + i}`;
    const newSheet = templateSheet.copyTo(ss).setName(sheetName);
    sheetNames.push(sheetName);

    // Set the weekly dates in the new sheet: C2, E2, G2, I2, K2
    const dateCells = ["C2", "E2", "G2", "I2", "K2"];
    const dateValues = getWeekDates(currentMonday);
    for (let j = 0; j < 5; j++) {
      newSheet.getRange(dateCells[j]).setValue(dateValues[j]);
    }

    currentMonday.setDate(currentMonday.getDate() + 7);
  }

  // Build the summary formulas in B2: across for all 5 days x N weeks
  for (const sheetName of sheetNames) {
    // Dates Row (row 2): C2, E2, G2, I2, K2
    const dateCells = ["C2", "E2", "G2", "I2", "K2"];
    for (let i = 0; i < 5; i++) {
      summarySheet
        .getRange(2, colIndex)
        .setFormula(`='${sheetName}'!${dateCells[i]}`);
      colIndex++;
    }
  }

  // Reset column pointer for behavior formulas
  let behaviorCol = clearSummary ? 2 : colIndex - 5 * sheetNames.length;

  // Behavior rows: 3–9 (Behavior 1–7)
  for (let behaviorRow = 9; behaviorRow <= 15; behaviorRow++) {
    let col = behaviorCol;
    for (const sheetName of sheetNames) {
      const behaviorCols = ["O", "Q", "S", "U", "W"];
      for (let i = 0; i < 5; i++) {
        const cell = `${behaviorCols[i]}${behaviorRow}`;
        summarySheet
          .getRange(behaviorRow - 6, col)
          .setFormula(`='${sheetName}'!${cell}`);
        col++;
      }
    }
  }

  // Clear emoji summary only if starting from scratch
  if (clearSummary) {
    summarySheet.getRange("E13:G19").clearContent();
  }

  // 🔁 Build emoji colors and their target summary columns
  const emojis = [
    { symbol: "🟦", col: "E" }, // Blue
    { symbol: "🟧", col: "F" }, // Orange
    { symbol: "🟥", col: "G" }, // Red
  ];

  // 🔁 Row mappings for behavior rows P9–P15
  const startRow = 13; // Summary sheet starting row
  const behaviorRowIndexes = [9, 10, 11, 12, 13, 14, 15]; // Source rows in week sheets
  const columnLetters = ["P", "R", "T", "V", "X"]; // Mon–Fri

  for (let i = 0; i < behaviorRowIndexes.length; i++) {
    const weekRow = behaviorRowIndexes[i];
    const summaryRow = startRow + i;

    for (const emoji of emojis) {
      const formulaParts = [];

      for (const sheetName of sheetNames) {
        const refs = columnLetters
          .map((col) => `'${sheetName}'!${col}${weekRow}`)
          .join(" & ");
        const part = `(LEN(${refs}) - LEN(SUBSTITUTE(${refs}, "${emoji.symbol}", ""))) / 2`;
        formulaParts.push(part);
      }

      const fullFormula = `=${formulaParts.join(" + ")}`;
      summarySheet
        .getRange(`${emoji.col}${summaryRow}`)
        .setFormula(fullFormula);
    }
  }

  ss.toast(
    `${numWeeks} weekly sheets created starting with Week ${startingWeekNumber} and the summary sheet was also populated.`,
    "Setup Completed",
    8,
  );
}

// Helper: Get dates Mon–Fri starting from Monday
function getWeekDates(monday) {
  const weekDates = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(monday);
    d.setDate(d.getDate() + i);
    weekDates.push(formatDate(d));
  }
  return weekDates;
}

// Helper: Format date as MM/DD/YYYY
function formatDate(date) {
  const mm = (date.getMonth() + 1).toString().padStart(2, "0");
  const dd = date.getDate().toString().padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
}

function insertBehaviorTrendChart() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Summary");

  // Remove existing charts
  summarySheet.getCharts().forEach((chart) => summarySheet.removeChart(chart));

  const lastCol = summarySheet.getLastColumn();
  const chartRange = summarySheet.getRange(2, 1, 8, lastCol); // A2:??9

  const chart = summarySheet
    .newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(chartRange)
    .setTransposeRowsAndColumns(true)
    .setOption("title", "Behavior Trends Over Time")
    .setOption("legend", { position: "right" })
    .setOption("hAxis", {
      title: "Date",
      baselineColor: "#000000", // black X axis line
      textPosition: "out",
      gridlines: { color: "none" }, // removes horizontal gridlines
      ticks: [], // leave empty to auto-generate ticks
    })
    .setOption("vAxis", {
      title: "Count",
      minValue: 0,
      baselineColor: "#000000", // black Y axis line
      textPosition: "out",
      gridlines: { color: "none" }, // removes vertical gridlines
      ticks: [], // leave empty to auto-generate ticks
    })
    .setPosition(11, 9, 0, 0);

  summarySheet.insertChart(chart.build());
}

/**
 * Displays help information about the spreadsheet. The function creates a modal dialog with help content.
 * 
 * The function is commented out since all of the sheets were created for the teachers at the start of the school year.
 */
// function helpInformation() {
//   const html = HtmlService.createHtmlOutputFromFile("Help")
//     .setWidth(1000)
//     .setHeight(600);
//   SpreadsheetApp.getUi().showModalDialog(
//     html,
//     "Radar Data Collection Spreadsheet",
//   );
// }
