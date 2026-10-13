/**
 * CURVA ULTRA DATABASE - BACKEND CONTROLLER & AI INGESTION ENGINE
 * Model: gemini-3.5-flash-lite
 */

const GEMINI_MODEL = "gemini-3.5-flash-lite";

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const data = JSON.parse(e.postData.contents);
    const adminPasscode = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSCODE");

    // 1. Verify Authentication
    if (!data.passcode || data.passcode !== adminPasscode) {
      return jsonResponse({ status: "error", message: "Unauthorized: Invalid passcode" }, 401);
    }

    const action = data.action;

    // 2. Route Actions
    switch (action) {
      case "stage_pdf":
        return jsonResponse(stagePdf(data.base64Data, data.fileName));

      case "ingest_chunk_1": // Group Metadata + Sections 1 & 2
        return jsonResponse(processChunk1(data.fileId));

      case "ingest_chunk_2": // Sections 3 & 4
        return jsonResponse(processChunk2(data.fileId, data.groupId, data.ultraGroup));

      case "ingest_chunk_3": // Sections 5 & 6
        return jsonResponse(processChunk3(data.fileId, data.groupId, data.ultraGroup));

      case "ingest_chunk_4": // Sources Appendix & Cleanup
        return jsonResponse(processChunk4(data.fileId, data.groupId, data.ultraGroup));

      case "update_field": // Manual Inline Cell Mutation
        return jsonResponse(updateSingleField(data.tabName, data.idColumn, data.idValue, data.field, data.value));

      case "create_media": // Add Media Item
        return jsonResponse(createMediaItem(data.payload));

      case "delete_media": // Delete Media Item
        return jsonResponse(deleteRowById("Media", "media_id", data.mediaId));

      case "delete_group": // Delete Ultra Group & Cascading Dependents
        return jsonResponse(deleteGroup(data.groupId));

      case "abort_ingestion": // Rollback staged PDF file & incomplete group
        return jsonResponse(abortIngestion(data.fileId, data.groupId));

      case "update_group_profile": // Update Ultra Group Basic Dimensions & Tags
        return jsonResponse(updateGroupProfile(data.groupId, data.profile));

      case "create_source": // Add Manual Source Item
        return jsonResponse(createSourceItem(data.payload));

      case "delete_source": // Delete Source Item
        return jsonResponse(deleteRowById("Sources", "source_id", data.sourceId));
      
      case "verify_passcode":
        return jsonResponse({ status: "success", verified: true });

      case "batch_update_responses":
        return jsonResponse(batchUpdateResponses(data.updates));

      case "batch_update_sources":
      return jsonResponse(batchUpdateSources(data.updates));

      case "batch_update_media":
        return jsonResponse(batchUpdateMedia(data.updates));

      default:
        return jsonResponse({ status: "error", message: "Unknown action: " + action }, 400);
    }
  } catch (err) {
    logError("GENERAL", "doPost", err.toString(), "");
    return jsonResponse({ status: "error", message: err.toString() }, 500);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 1. FILE STAGING (GOOGLE DRIVE)
// ==========================================

function stagePdf(base64Data, fileName) {
  let folderIterator = DriveApp.getFoldersByName("CURVA_Staging_Temp");
  let folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder("CURVA_Staging_Temp");

  const cleanBase64 = base64Data.replace(/^data:application\/pdf;base64,/, "");
  const decoded = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(decoded, "application/pdf", fileName || "Curva_Upload.pdf");
  const file = folder.createFile(blob);

  return { status: "success", fileId: file.getId() };
}

function cleanupStagedFile(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {}
}

function abortIngestion(fileId, groupId) {
  var driveCleaned = false;
  var groupDeleted = false;
  if (fileId) {
    try {
      cleanupStagedFile(fileId);
      driveCleaned = true;
    } catch (e) {
      logError("ROLLBACK", "cleanupStagedFile", e.toString(), fileId);
    }
  }
  if (groupId) {
    try {
      deleteGroup(groupId);
      groupDeleted = true;
    } catch (e) {
      logError("ROLLBACK", "deleteGroup", e.toString(), groupId);
    }
  }
  return { status: "success", aborted: true, driveCleaned: driveCleaned, groupDeleted: groupDeleted };
}

// ==========================================
// 2. CHUNKED EXTRACTION PIPELINE
// ==========================================

// Step 1: Metadata + Sections 1 & 2
function getNextGroupId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return "GRP_0001";
  
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let maxId = 0;
  values.forEach(val => {
    if (typeof val === "string" && val.startsWith("GRP_")) {
      const num = parseInt(val.replace("GRP_", ""), 10);
      if (!isNaN(num) && num > maxId) maxId = num;
    }
  });
  return "GRP_" + String(maxId + 1).padStart(4, "0");
}

function processChunk1(fileId) {
  const file = DriveApp.getFileById(fileId);
  const base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());

  const prompt = `You are parsing a CURVA Ultra research PDF.
Extract the PAGE 1 GROUP PROFILE and all questions from:
- SECTION 1: Group Identification (1.1 to 1.6)
- SECTION 2: Historical Dimension (2.1 to 2.5)

Do NOT extract source IDs or citations.
Return a strict JSON object with this exact schema:
{
  "profile": {
    "ultra_group": "Group Name",
    "associated_club": "Club Name",
    "city": "City",
    "country": "Country",
    "group_tags": "Tag1, Tag2, Tag3"
  },
  "responses": [
    {
      "section_number": 1,
      "section_name": "Group Identification (Baseline Metadata)",
      "question_number": "1.1",
      "question": "Full exact text of question",
      "answer": "Answer text",
      "confidence": "Well-supported | Partially supported | Not found with sufficient confidence | Not applicable",
      "confidence_note": "Note or empty string",
      "research_notes": "Notes or empty string"
    }
  ]
}`;

  const result = callGemini(base64Pdf, prompt);
  logRaw(result.groupId || "PENDING", "chunk_1", JSON.stringify(result));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupsSheet = ss.getSheetByName("Groups");
  
  const groupId = getNextGroupId(groupsSheet);
  const p = result.profile;

  groupsSheet.appendRow([
    groupId,
    p.ultra_group,
    p.associated_club,
    p.city,
    p.country,
    p.group_tags,
    new Date().toISOString(),
    ""
  ]);

  appendResponses(groupId, p.ultra_group, result.responses);

  return { status: "success", groupId: groupId, ultraGroup: p.ultra_group, count: result.responses.length };
}

// Step 2: Sections 3 & 4
function processChunk2(fileId, groupId, ultraGroup) {
  const file = DriveApp.getFileById(fileId);
  const base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());

  const prompt = `You are parsing a CURVA Ultra research PDF.
Extract all questions from:
- SECTION 3: Behavioral Dimension (3.1 to 3.7)
- SECTION 4: Ideological Dimension (4.1 to 4.15)

Do NOT extract source IDs or citations.
Return a strict JSON object with this exact schema:
{
  "responses": [
    {
      "section_number": 3,
      "section_name": "Behavioral Dimension",
      "question_number": "3.1",
      "question": "Full exact text of question",
      "answer": "Answer text",
      "confidence": "Well-supported | Partially supported | Not found with sufficient confidence | Not applicable",
      "confidence_note": "",
      "research_notes": ""
    }
  ]
}`;

  const result = callGemini(base64Pdf, prompt);
  logRaw(groupId, "chunk_2", JSON.stringify(result));
  appendResponses(groupId, ultraGroup, result.responses);

  return { status: "success", groupId: groupId, count: result.responses.length };
}

// Step 3: Sections 5 & 6
function processChunk3(fileId, groupId, ultraGroup) {
  const file = DriveApp.getFileById(fileId);
  const base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());

  const prompt = `You are parsing a CURVA Ultra research PDF.
Extract all questions from:
- SECTION 5: Organizational Dimension (5.1 to 5.20)
- SECTION 6: Representational Dimension (6.1 to 6.21)

Do NOT extract source IDs or citations.
Return a strict JSON object with this exact schema:
{
  "responses": [
    {
      "section_number": 5,
      "section_name": "Organizational Dimension",
      "question_number": "5.1",
      "question": "Full exact text of question",
      "answer": "Answer text",
      "confidence": "Well-supported | Partially supported | Not found with sufficient confidence | Not applicable",
      "confidence_note": "",
      "research_notes": ""
    }
  ]
}`;

  const result = callGemini(base64Pdf, prompt);
  logRaw(groupId, "chunk_3", JSON.stringify(result));
  appendResponses(groupId, ultraGroup, result.responses);

  return { status: "success", groupId: groupId, count: result.responses.length };
}

// Step 4: Sources Appendix & Cleanup
function processChunk4(fileId, groupId, ultraGroup) {
  const file = DriveApp.getFileById(fileId);
  const base64Pdf = Utilities.base64Encode(file.getBlob().getBytes());

  const prompt = `You are parsing a CURVA Ultra research PDF.
Locate the "Works cited" or "SOURCES" appendix at the end of the document.
Extract each entry into a clean list. Infer the source_type based on the URL and title.

Return a strict JSON object with this exact schema:
{
  "sources": [
    {
      "source_type": "Academic | Journalism | Official | Documentary / Video | Social Media / Community | Other",
      "title": "Title of article, video, or webpage",
      "url": "https://..."
    }
  ]
}`;

  const result = callGemini(base64Pdf, prompt);
  logRaw(groupId, "chunk_4", JSON.stringify(result));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourcesSheet = ss.getSheetByName("Sources");

  if (result.sources && result.sources.length > 0) {
    let currentMaxIndex = 0;
    const lastRow = sourcesSheet.getLastRow();
    if (lastRow > 1) {
      const ids = sourcesSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
      ids.forEach(id => {
        if (typeof id === "string" && id.startsWith(groupId + "_S")) {
          const num = parseInt(id.replace(groupId + "_S", ""), 10);
          if (!isNaN(num) && num > currentMaxIndex) currentMaxIndex = num;
        }
      });
    }

    const rows = result.sources.map(s => {
      currentMaxIndex++;
      const scopedSourceId = `${groupId}_S${String(currentMaxIndex).padStart(3, "0")}`;
      return [
        scopedSourceId,
        groupId,
        ultraGroup,
        s.source_type || "Other",
        s.title || "",
        s.url || ""
      ];
    });

    sourcesSheet.getRange(sourcesSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  // Delete staged PDF from Drive
  cleanupStagedFile(fileId);

  return { status: "success", groupId: groupId, sourcesCount: result.sources ? result.sources.length : 0 };
}

// ==========================================
// 3. PARSING & APPENDING HELPERS
// ==========================================

function appendResponses(groupId, ultraGroup, responses) {
  if (!responses || responses.length === 0) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Responses");

  let currentMaxIndex = 0;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    ids.forEach(id => {
      if (typeof id === "string" && id.startsWith(groupId + "_RES_")) {
        const num = parseInt(id.replace(groupId + "_RES_", ""), 10);
        if (!isNaN(num) && num > currentMaxIndex) currentMaxIndex = num;
      }
    });
  }

  const rows = responses.map(r => {
    currentMaxIndex++;
    const responseId = `${groupId}_RES_${String(currentMaxIndex).padStart(3, "0")}`;

    // Split leading binary/multi-choice token at the first period
    let binaryChoice = "";
    let answerText = r.answer ? r.answer.trim() : "";
    const periodIndex = answerText.indexOf(". ");
    if (periodIndex !== -1 && periodIndex <= 35) {
      binaryChoice = answerText.substring(0, periodIndex).trim();
      answerText = answerText.substring(periodIndex + 2).trim();
    }

    return [
      responseId,
      groupId,
      ultraGroup,
      r.section_number,
      r.section_name,
      r.question_number,
      r.question,
      binaryChoice,
      answerText,
      r.confidence,
      r.confidence_note || "",
      r.research_notes || ""
    ];
  });

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

// ==========================================
// 4. GEMINI API CALLER
// ==========================================

function callGemini(base64Data, promptText) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [
      {
        parts: [
          { text: promptText },
          {
            inlineData: {
              mimeType: "application/pdf",
              data: base64Data
            }
          }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const rawText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`Gemini API Error (${responseCode}): ${rawText}`);
  }

  const json = JSON.parse(rawText);
  const candidate = json.candidates && json.candidates[0];
  if (!candidate || !candidate.content || !candidate.content.parts) {
    throw new Error("Invalid output format returned by Gemini API");
  }

  const contentText = candidate.content.parts[0].text;
  return JSON.parse(contentText);
}

// ==========================================
// 5. INLINE MUTATION & CRUD HELPERS
// ==========================================

function updateSingleField(tabName, idColumn, idValue, field, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error("Tab not found: " + tabName);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.indexOf(field);
  const idIndex = headers.indexOf(idColumn);

  if (colIndex === -1) throw new Error(`Field '${field}' not found in ${tabName}`);
  if (idIndex === -1) throw new Error(`ID column '${idColumn}' not found in ${tabName}`);

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idIndex]) === String(idValue)) {
      sheet.getRange(r + 1, colIndex + 1).setValue(value);
      return { status: "success", updated: { row: r + 1, field: field, value: value } };
    }
  }

  throw new Error(`Record with ${idColumn} = ${idValue} not found`);
}

function createMediaItem(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Media");
  const lastRow = sheet.getLastRow();
  
  let maxIndex = 0;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    ids.forEach(id => {
      if (typeof id === "string" && id.startsWith(p.groupId + "_MED_")) {
        const num = parseInt(id.replace(p.groupId + "_MED_", ""), 10);
        if (!isNaN(num) && num > maxIndex) maxIndex = num;
      }
    });
  }

  const mediaId = `${p.groupId}_MED_${String(maxIndex + 1).padStart(3, "0")}`;
  sheet.appendRow([
    mediaId,
    p.groupId,
    p.ultraGroup,
    p.mediaType,
    p.mediaLink,
    p.tags || "",
    p.mediaTitle,
    p.mediaDescription || "",
    p.thumbnailUrl || ""
  ]);

  return { status: "success", mediaId: mediaId };
}

function createSourceItem(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sources");
  const lastRow = sheet.getLastRow();

  let maxIndex = 0;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    ids.forEach(id => {
      if (typeof id === "string" && id.startsWith(p.groupId + "_S")) {
        const num = parseInt(id.replace(p.groupId + "_S", ""), 10);
        if (!isNaN(num) && num > maxIndex) maxIndex = num;
      }
    });
  }

  const sourceId = `${p.groupId}_S${String(maxIndex + 1).padStart(3, "0")}`;
  sheet.appendRow([
    sourceId,
    p.groupId,
    p.ultraGroup,
    p.sourceType || "Other",
    p.title || "",
    p.url || ""
  ]);

  return { status: "success", sourceId: sourceId };
}

function deleteRowById(tabName, idColumn, idValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(tabName);
  const data = sheet.getDataRange().getValues();
  const idIndex = data[0].indexOf(idColumn);

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idIndex]) === String(idValue)) {
      sheet.deleteRow(r + 1);
      return { status: "success", deletedId: idValue };
    }
  }

  throw new Error(`Record with ${idColumn} = ${idValue} not found`);
}

function deleteRowsMatchingValue(sheet, colName, matchVal) {
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;

  const colIndex = data[0].indexOf(colName);
  if (colIndex === -1) return 0;

  const filtered = [data[0]]; // preserve header row
  let deletedCount = 0;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]) === String(matchVal)) {
      deletedCount++;
    } else {
      filtered.push(data[i]);
    }
  }

  if (deletedCount > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, filtered.length, filtered[0].length).setValues(filtered);
  }

  return deletedCount;
}

function deleteGroup(groupId) {
  if (!groupId) throw new Error("groupId is required for deletion");
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Fast batch deletion across Groups, Media, Responses, and Sources
  deleteRowsMatchingValue(ss.getSheetByName("Groups"), "group_id", groupId);
  deleteRowsMatchingValue(ss.getSheetByName("Media"), "group_id", groupId);
  deleteRowsMatchingValue(ss.getSheetByName("Responses"), "group_id", groupId);
  deleteRowsMatchingValue(ss.getSheetByName("Sources"), "group_id", groupId);

  return { status: "success", deletedGroupId: groupId };
}

function updateGroupProfile(groupId, profile) {
  if (!groupId) throw new Error("groupId is required for profile update");
  if (!profile) return { status: "success", updated: false };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const groupsSheet = ss.getSheetByName("Groups");
  if (!groupsSheet) throw new Error("Groups sheet not found");

  const gData = groupsSheet.getDataRange().getValues();
  const headers = gData[0];
  const gIdCol = headers.indexOf("group_id");
  const nameCol = headers.indexOf("ultra_group");
  const clubCol = headers.indexOf("associated_club");
  const cityCol = headers.indexOf("city");
  const countryCol = headers.indexOf("country");
  const tagsCol = headers.indexOf("group_tags");

  let targetRow = -1;
  let oldGroupName = "";

  for (let r = 1; r < gData.length; r++) {
    if (String(gData[r][gIdCol]) === String(groupId)) {
      targetRow = r + 1;
      oldGroupName = String(gData[r][nameCol] || "");
      break;
    }
  }

  if (targetRow === -1) {
    throw new Error(`Group with group_id = ${groupId} not found`);
  }

  if (profile.ultra_group !== undefined && nameCol !== -1) {
    groupsSheet.getRange(targetRow, nameCol + 1).setValue(profile.ultra_group);
  }
  if (profile.associated_club !== undefined && clubCol !== -1) {
    groupsSheet.getRange(targetRow, clubCol + 1).setValue(profile.associated_club);
  }
  if (profile.city !== undefined && cityCol !== -1) {
    groupsSheet.getRange(targetRow, cityCol + 1).setValue(profile.city);
  }
  if (profile.country !== undefined && countryCol !== -1) {
    groupsSheet.getRange(targetRow, countryCol + 1).setValue(profile.country);
  }
  if (profile.group_tags !== undefined && tagsCol !== -1) {
    groupsSheet.getRange(targetRow, tagsCol + 1).setValue(profile.group_tags);
  }

  // If group name changed, sync denormalized ultra_group in Media, Responses, and Sources
  if (profile.ultra_group && profile.ultra_group !== oldGroupName) {
    const newName = profile.ultra_group;

    // Sync Media
    const mediaSheet = ss.getSheetByName("Media");
    if (mediaSheet) {
      const mData = mediaSheet.getDataRange().getValues();
      const mIdCol = mData[0].indexOf("group_id");
      const mNameCol = mData[0].indexOf("ultra_group");
      if (mIdCol !== -1 && mNameCol !== -1) {
        for (let r = 1; r < mData.length; r++) {
          if (String(mData[r][mIdCol]) === String(groupId)) {
            mediaSheet.getRange(r + 1, mNameCol + 1).setValue(newName);
          }
        }
      }
    }

    // Sync Responses
    const respSheet = ss.getSheetByName("Responses");
    if (respSheet) {
      const rData = respSheet.getDataRange().getValues();
      const rIdCol = rData[0].indexOf("group_id");
      const rNameCol = rData[0].indexOf("ultra_group");
      if (rIdCol !== -1 && rNameCol !== -1) {
        for (let r = 1; r < rData.length; r++) {
          if (String(rData[r][rIdCol]) === String(groupId)) {
            respSheet.getRange(r + 1, rNameCol + 1).setValue(newName);
          }
        }
      }
    }

    // Sync Sources
    const srcSheet = ss.getSheetByName("Sources");
    if (srcSheet) {
      const sData = srcSheet.getDataRange().getValues();
      const sIdCol = sData[0].indexOf("group_id");
      const sNameCol = sData[0].indexOf("ultra_group");
      if (sIdCol !== -1 && sNameCol !== -1) {
        for (let r = 1; r < sData.length; r++) {
          if (String(sData[r][sIdCol]) === String(groupId)) {
            srcSheet.getRange(r + 1, sNameCol + 1).setValue(newName);
          }
        }
      }
    }
  }

  return { status: "success", groupId: groupId, updated: true };
}

// ==========================================
// 6. LOGGING & UTILITIES
// ==========================================

function logRaw(groupId, stepName, rawJson) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw_Logs");
    sheet.appendRow([new Date().toISOString(), groupId, stepName, rawJson]);
  } catch (e) {}
}

function logError(groupId, stepName, errorMessage, snippet) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Errors");
    sheet.appendRow([new Date().toISOString(), groupId, stepName, errorMessage, snippet]);
  } catch (e) {}
}

function jsonResponse(obj, statusCode) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// For updating pre-existing ultra group/source/media information
function batchUpdateResponses(updates) {
  if (!updates || updates.length === 0) return { status: "success", updated: 0 };
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Responses");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf("response_id");
  const tokenCol = headers.indexOf("binary_or_choice");
  const answerCol = headers.indexOf("answer_text");
  const confCol = headers.indexOf("confidence");
  const notesCol = headers.indexOf("research_notes");

  // Map response_id to row index
  const rowMap = {};
  for (let r = 1; r < data.length; r++) {
    rowMap[String(data[r][idCol])] = r + 1;
  }

  let updateCount = 0;
  updates.forEach(u => {
    const rowNum = rowMap[u.response_id];
    if (rowNum) {
      if (u.binary_or_choice !== undefined && tokenCol !== -1) {
        sheet.getRange(rowNum, tokenCol + 1).setValue(u.binary_or_choice);
      }
      if (u.answer_text !== undefined && answerCol !== -1) {
        sheet.getRange(rowNum, answerCol + 1).setValue(u.answer_text);
      }
      if (u.confidence !== undefined && confCol !== -1) {
        sheet.getRange(rowNum, confCol + 1).setValue(u.confidence);
      }
      if (u.research_notes !== undefined && notesCol !== -1) {
        sheet.getRange(rowNum, notesCol + 1).setValue(u.research_notes);
      }
      updateCount++;
    }
  });

  return { status: "success", updated: updateCount };
}

function batchUpdateSources(updates) {
  if (!updates || updates.length === 0) return { status: "success", updated: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Sources");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf("source_id");
  const typeCol = headers.indexOf("source_type");
  const titleCol = headers.indexOf("title");
  const urlCol = headers.indexOf("url");

  const rowMap = {};
  for (let r = 1; r < data.length; r++) {
    rowMap[String(data[r][idCol])] = r + 1;
  }

  let count = 0;
  updates.forEach(u => {
    const rowNum = rowMap[u.source_id];
    if (rowNum) {
      if (u.source_type !== undefined && typeCol !== -1) sheet.getRange(rowNum, typeCol + 1).setValue(u.source_type);
      if (u.title !== undefined && titleCol !== -1) sheet.getRange(rowNum, titleCol + 1).setValue(u.title);
      if (u.url !== undefined && urlCol !== -1) sheet.getRange(rowNum, urlCol + 1).setValue(u.url);
      count++;
    }
  });
  return { status: "success", updated: count };
}

function batchUpdateMedia(updates) {
  if (!updates || updates.length === 0) return { status: "success", updated: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Media");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idCol = headers.indexOf("media_id");
  const typeCol = headers.indexOf("media_type");
  const linkCol = headers.indexOf("media_link");
  const tagsCol = headers.indexOf("tags");
  const titleCol = headers.indexOf("media_title");
  const descCol = headers.indexOf("media_description");
  const thumbCol = headers.indexOf("thumbnail_url");

  const rowMap = {};
  for (let r = 1; r < data.length; r++) {
    rowMap[String(data[r][idCol])] = r + 1;
  }

  let count = 0;
  updates.forEach(u => {
    const rowNum = rowMap[u.media_id];
    if (rowNum) {
      if (u.media_type !== undefined && typeCol !== -1) sheet.getRange(rowNum, typeCol + 1).setValue(u.media_type);
      if (u.media_link !== undefined && linkCol !== -1) sheet.getRange(rowNum, linkCol + 1).setValue(u.media_link);
      if (u.tags !== undefined && tagsCol !== -1) sheet.getRange(rowNum, tagsCol + 1).setValue(u.tags);
      if (u.media_title !== undefined && titleCol !== -1) sheet.getRange(rowNum, titleCol + 1).setValue(u.media_title);
      if (u.media_description !== undefined && descCol !== -1) sheet.getRange(rowNum, descCol + 1).setValue(u.media_description);
      if (u.thumbnail_url !== undefined && thumbCol !== -1) sheet.getRange(rowNum, thumbCol + 1).setValue(u.thumbnail_url);
      count++;
    }
  });
  return { status: "success", updated: count };
}
