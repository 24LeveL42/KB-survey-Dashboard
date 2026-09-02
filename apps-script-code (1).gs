/**
 * KOPI BOY RESEARCH — SURVEY BACKEND
 * ------------------------------------------------------------
 * doPost(e)  — receives submissions from rider-survey.html and
 *              merchant-survey.html, archives each as a row in
 *              a Google Sheet — one tab per survey type
 *              ("Riders" and "Merchants"), created automatically.
 *
 * doGet(e)   — read-only endpoint for the results dashboard.
 *              Returns both tabs' headers + rows as JSON so the
 *              dashboard can render live bar/pie charts.
 *
 * SETUP (unchanged from before):
 * 1. sheets.google.com > this spreadsheet > Extensions > Apps Script.
 * 2. Delete existing code, paste this whole file in its place.
 * 3. Deploy > Manage deployments > pencil icon > "New version" > Deploy.
 *    (Editing an existing deployment this way keeps the same
 *    /exec URL — you do NOT need to update the survey HTML files
 *    or the dashboard again.)
 */

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const surveyType = data.surveyType === "rider" ? "Riders" : "Merchants";

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(surveyType);

    if (!sheet) {
      sheet = ss.insertSheet(surveyType);
      const headers = surveyType === "Riders"
        ? ["Timestamp", "Q1 Currently do gig work", "Q2 Platforms used", "Q3 Hours per week",
           "Q4 Prefer immediate payment", "Q5 Comfortable paid by merchant directly",
           "Q6 Prefer choosing own jobs", "Q7 Tier system feels unfair",
           "Q8 Interested in casual no-schedule work", "Q9 Fair fee matters more than tier",
           "Q10 Biggest frustration (open text)"]
        : ["Timestamp", "Q1 Currently sell food", "Q2 Used delivery platform before",
           "Q3 Commission eats into profit", "Q4 Prefer flat subscription over commission",
           "Q5 $20-25/month feels fair", "Q6 Comfortable choosing own delivery helper",
           "Q7 Own refund policy gives more confidence", "Q8 Reach is harder than commission cost",
           "Q9 What would help sell more (open text)"];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    const row = surveyType === "Riders"
      ? [new Date(), data.q1 || "", data.q2 || "", data.q3 || "", data.q4 || "",
         data.q5 || "", data.q6 || "", data.q7 || "", data.q8 || "", data.q9 || "", data.q10 || ""]
      : [new Date(), data.q1 || "", data.q2 || "", data.q3 || "", data.q4 || "",
         data.q5 || "", data.q6 || "", data.q7 || "", data.q8 || "", data.q9 || ""];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = {};

    ["Riders", "Merchants"].forEach(function (name) {
      const sheet = ss.getSheetByName(name);
      if (!sheet || sheet.getLastRow() < 1) {
        result[name] = { headers: [], rows: [] };
        return;
      }
      const values = sheet.getDataRange().getValues();
      const headers = values.shift() || [];
      // Timestamps come through as Date objects — stringify them.
      const rows = values.map(function (row) {
        return row.map(function (cell) {
          return cell instanceof Date ? cell.toISOString() : cell;
        });
      });
      result[name] = { headers: headers, rows: rows };
    });

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
