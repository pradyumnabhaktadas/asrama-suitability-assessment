/**
 * Āśrama Suitability Assessment - Google Apps Script Backend
 * Attached to Spreadsheet ID: 15aBFnsXSWvwdvsSsjFYgAJq5Hk84B3szhEJhZRuMusY
 * 
 * Supports:
 *  1. doPost(e): Accepts JSON submissions from candidates on GitHub Pages and appends rows.
 *  2. doGet(e): Returns all candidate responses as JSON (?action=getCandidates) to populate Mentor Portal.
 */

function doGet(e) {
  e = e || { parameter: {} };
  var action = (e.parameter && e.parameter.action) || "";

  // Return all candidate records as JSON for Mentor Portal
  if (action === "getCandidates" || action === "getRoster") {
    return getCandidatesAsJson();
  }

  // Health check / status info
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    service: "Āśrama Suitability Assessment Webhook",
    spreadsheetId: "15aBFnsXSWvwdvsSsjFYgAJq5Hk84B3szhEJhZRuMusY",
    message: "Google Apps Script Backend is active and receiving submissions."
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);
    var result = recordToSheet(data);
    return ContentService.createTextOutput(JSON.stringify({ success: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Direct submission if called within Apps Script environment
function submitAssessment(data) {
  return recordToSheet(data);
}

function getCandidatesAsJson() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    var candidates = [];
    // Row 0 is header. Rows 1..N are candidate submissions
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || !row[1]) continue; // Skip empty row or missing name
      
      var rawPayload = row[26]; // Column AA (index 26) contains JSON.stringify(data)
      var candObj = null;
      if (rawPayload && typeof rawPayload === "string" && rawPayload.startsWith("{")) {
        try {
          candObj = JSON.parse(rawPayload);
        } catch(ex) {}
      }
      
      if (!candObj) {
        // Fallback reconstruction if raw payload string wasn't stored
        candObj = {
          id: "sheet_row_" + i,
          updatedAt: row[0] || new Date().toISOString(),
          a: {
            name: row[1] || "",
            age: row[2] || "",
            occupation: row[3] || "",
            living: row[4] || "",
            rounds: row[5] || "",
            mentor: row[6] || ""
          },
          calc: {
            totalScore: Number(row[7]) || 0,
            indication: row[8] || "",
            gatekeeperStatus: row[9] || ""
          }
        };
      }
      
      if (!candObj.id) {
        candObj.id = "sheet_row_" + i + "_" + ((candObj.a && candObj.a.name) ? candObj.a.name.replace(/\s+/g, '_') : 'devotee');
      }
      if (!candObj.updatedAt && row[0]) {
        candObj.updatedAt = row[0];
      }
      candidates.push(candObj);
    }

    // Return newest candidates first
    candidates.reverse();

    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      count: candidates.length,
      candidates: candidates
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function recordToSheet(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  // Scoring Key Mapping
  var key = {
    1:  { c: 2, b: 1, a: 0 },
    2:  { a: 2, b: 1, c: 0 },
    3:  { b: 2, c: 1, a: 0 },
    4:  { b: 2, c: 1, a: 0 },
    5:  { a: 2, b: 1, c: 0 },
    6:  { c: 2, b: 1, a: 0 },
    7:  { a: 2, c: 1, b: 0 },
    8:  { c: 2, b: 1, a: 0 },
    9:  { b: 2, c: 1, a: 0 },
    10: { b: 2, c: 1, a: 0 },
    11: { b: 2, c: 1, a: 0 },
    12: { a: 2, c: 1, b: 0 },
    13: { b: 2, c: 1, a: 0 },
    14: { a: 2, b: 1, c: 0 },
    15: { b: 2, c: 1, a: 0 },
    16: { a: 2, c: 1, b: 0 },
    17: { b: 2, c: 1, a: 0 },
    18: { a: 2, b: 1, c: 0 },
    19: { c: 2, b: 1, a: 0 },
    20: { c: 2, b: 1, a: 0 },
    21: { a: 2, b: 1, c: 0 },
    22: { b: 2, c: 1, a: 0 },
    23: { b: 2, c: 1, a: 0 },
    24: { b: 2, c: 1, a: 0 },
    25: { c: 2, a: 1, b: 0 },
    26: { a: 2, b: 1, c: 0 },
    27: { b: 2, a: 1, c: 0 },
    28: { b: 2, c: 1, a: 0 },
    29: { b: 2, c: 1, a: 0 },
    30: { a: 2, b: 1, c: 0 },
    31: { c: 2, b: 1, a: 0 },
    32: { b: 2, c: 1, a: 0 },
    33: { b: 2, c: 1, a: 0 },
    34: { a: 2, c: 1, b: 0 },
    35: { a: 2, c: 1, b: 0 },
    36: { c: 2, b: 1, a: 0 },
    37: { b: 2, c: 1, a: 0 },
    38: { a: 2, c: 1, b: 0 },
    39: { b: 2, c: 1, a: 0 },
    40: { a: 2, b: 1, c: 0 },
    41: { b: 2, c: 1, a: 0 },
    42: { c: 2, b: 1, a: 0 },
    43: { c: 2, b: 1, a: 0 },
    44: { c: 2, b: 1, a: 0 },
    45: { b: 2, c: 1, a: 0 },
    46: { c: 2, b: 1, a: 0 },
    47: { c: 2, a: 1, b: 0 },
    48: { a: 2, c: 1, b: 0 },
    49: { c: 2, b: 1, a: 0 },
    50: { a: 2, c: 1, b: 0 },
    51: { c: 2, b: 1, a: 0 },
    52: { a: 2, c: 1, b: 0 }
  };

  // 13 Areas ranges
  var areas = [
    [1, 4],   // Area 1
    [5, 8],   // Area 2
    [9, 12],  // Area 3
    [13, 16], // Area 4
    [17, 20], // Area 5
    [21, 24], // Area 6
    [25, 28], // Area 7
    [29, 32], // Area 8
    [33, 36], // Area 9
    [37, 40], // Area 10
    [41, 44], // Area 11
    [45, 48], // Area 12
    [49, 52]  // Area 13
  ];

  var areaScores = [];
  var totalScore = 0;
  var bAnswers = data.b || {};

  for (var a = 0; a < areas.length; a++) {
    var startQ = areas[a][0];
    var endQ = areas[a][1];
    var score = 0;
    for (var q = startQ; q <= endQ; q++) {
      var ans = bAnswers[q];
      if (ans && key[q] && key[q][ans] !== undefined) {
        score += key[q][ans];
      }
    }
    areaScores.push(score);
    totalScore += score;
  }

  // Gatekeeper check: Areas 1, 2, 4, 9 <= 3
  var gatekeeperFails = [];
  var gkIndices = [0, 1, 3, 8]; // correspond to areas 1, 2, 4, 9
  for (var g = 0; g < gkIndices.length; g++) {
    var idx = gkIndices[g];
    if (areaScores[idx] <= 3) {
      gatekeeperFails.push("Area " + (idx + 1) + " (" + areaScores[idx] + "/8)");
    }
  }

  var gatekeeperStatus = gatekeeperFails.length > 0 
    ? "ALERT: Low in " + gatekeeperFails.join(", ") 
    : "CLEARED (All ≥ 4)";

  // Primary Indication
  var indication = "";
  if (gatekeeperFails.length > 0) {
    indication = "Gṛhastha Nature Indicated (Gatekeeper low)";
  } else if (totalScore >= 81) {
    indication = "Brahmacārī Nature Clear (81–104)";
  } else if (totalScore >= 57) {
    indication = "Mixed / Still Forming (57–80)";
  } else {
    indication = "Gṛhastha Nature (0–56)";
  }

  // Dominant Varna
  var v = data.v || {};
  var counts = { a: 0, b: 0, c: 0, d: 0 };
  if (v.v1) counts[v.v1] = (counts[v.v1] || 0) + 1;
  if (v.v2) counts[v.v2] = (counts[v.v2] || 0) + 1;
  if (v.v3) counts[v.v3] = (counts[v.v3] || 0) + 1;

  var domVarna = "a";
  if (counts.b > counts[domVarna]) domVarna = "b";
  if (counts.c > counts[domVarna]) domVarna = "c";
  if (counts.d > counts[domVarna]) domVarna = "d";

  var varnaLabels = {
    a: "Brāhmaṇa (Scripture / Teaching)",
    b: "Kṣatriya (Leadership / Management)",
    c: "Vaiśya (Resources / Farming)",
    d: "Śūdra (Hands-on Seva / Assistance)"
  };
  var dominantVarnaText = varnaLabels[domVarna] || "Not Specified";

  // Summarize Practical Duties
  var duties = data.duties || {};
  var dutiesSummary = "Only Son: " + (duties.only_son || "N/A") + " | Financial: " + (duties.financial || "N/A")
    + " | Illness: " + (duties.illness || "N/A") + " | Loans: " + (duties.loans || "N/A")
    + " | Parents Attitude: " + (duties.attitude || "N/A") + " | Health: " + (duties.health || "N/A");

  // Summarize Part C answers
  var c = data.c || {};
  var cList = [];
  for (var i = 1; i <= 10; i++) {
    if (c[i]) cList.push("Q" + i + ": " + c[i]);
  }
  var cSummary = cList.join("\n---\n");

  var a = data.a || {};
  var timestamp = new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" });

  var newRow = [
    timestamp,
    a.name || "",
    a.age || "",
    a.occupation || "",
    a.living || "",
    a.rounds || "",
    a.mentor || "",
    totalScore,
    indication,
    gatekeeperStatus,
    dominantVarnaText,
    areaScores[0],
    areaScores[1],
    areaScores[2],
    areaScores[3],
    areaScores[4],
    areaScores[5],
    areaScores[6],
    areaScores[7],
    areaScores[8],
    areaScores[9],
    areaScores[10],
    areaScores[11],
    areaScores[12],
    dutiesSummary,
    cSummary,
    JSON.stringify(data)
  ];

  sheet.appendRow(newRow);

  return {
    success: true,
    candidate: a.name,
    totalScore: totalScore,
    indication: indication,
    gatekeeperStatus: gatekeeperStatus
  };
}
