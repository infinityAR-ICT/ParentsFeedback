const COLLEGE_SPREADSHEET_ID = "164F488YllJ0KnqDIbYsfUtZ2kPgTtsoZZOdSMRkYMR0";
const COLLEGE_SHEET_NAME = "College Evaluation";
const COLLEGE_VALID_OPTION_VALUES = [1, 2, 3, 4, 5];
const COLLEGE_MAX_COMMENT_LENGTH = 45000;
const COLLEGE_VALID_CLASSES = ["VII", "VIII", "IX", "X", "XI", "XII"];
const COLLEGE_VALID_HOUSES = [
  "Shaheed Suhrawardy House",
  "Sher-e-Bangla House",
  "Shariatullah House"
];
const COLLEGE_VALID_FORMS = ["A", "B"];
const COLLEGE_VALID_MISSION_OBJECTIVES = ["হ্যাঁ", "না", "অন্যান্য"];
const COLLEGE_RATING_FIELDS = [
  {key: "environmentRating", header: "College Environment Rating"},
  {key: "educationDevelopmentRating", header: "Education Development Measures Rating"},
  {key: "diningQualityRating", header: "Cadet Dining and Food Quality Rating"},
  {key: "houseEnvironmentRating", header: "House Environment Rating"},
  {key: "formMasterSincerityRating", header: "Form Master Sincerity Rating"},
  {key: "houseMasterDedicationRating", header: "House Master Sincerity and Dedication Rating"},
  {key: "hospitalFacilitiesRating", header: "Hospital Facilities and Environment Rating"},
  {key: "armedForcesViewRating", header: "Your view about the Armed Forces"},
  {key: "collegeAdjutantSincerityRating", header: "Your view about The Adjutant"},
  {key: "vicePrincipalViewRating", header: "Your view about The Vice-Principal"},
  {key: "principalViewRating", header: "Your view about The Principal"}
];
const COLLEGE_REQUIRED_RATING_KEYS = [
  "environmentRating",
  "educationDevelopmentRating",
  "houseEnvironmentRating",
  "formMasterSincerityRating",
  "houseMasterDedicationRating",
  "hospitalFacilitiesRating",
  "collegeAdjutantSincerityRating",
  "vicePrincipalViewRating",
  "principalViewRating"
];

function setupCollegeEvaluationSheet() {
  const spreadsheet = SpreadsheetApp.openById(COLLEGE_SPREADSHEET_ID);
  const sheet = collegeEnsureSheet(spreadsheet);
  SpreadsheetApp.flush();
  const result = {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName(),
    sheetName: sheet.getName(),
    headers: collegeBuildHeaders()
  };
  console.log(JSON.stringify(result));
  return result;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const requestBody = e && e.parameter && typeof e.parameter.payload === "string"
      ? e.parameter.payload
      : e && e.postData && typeof e.postData.contents === "string"
        ? e.postData.contents
        : "";
    if (!requestBody) {
      throw new Error("Missing request body.");
    }
    const payload = JSON.parse(requestBody);
    if (payload.evaluationType && payload.evaluationType !== "college") {
      throw new Error("Invalid evaluation type.");
    }
    const submission = {
      className: collegeValidateChoice(payload.className, COLLEGE_VALID_CLASSES, "class"),
      houseName: collegeValidateChoice(payload.houseName, COLLEGE_VALID_HOUSES, "house"),
      formName: collegeValidateChoice(payload.formName, COLLEGE_VALID_FORMS, "form"),
      cadetName: collegeValidateCadetName(payload.cadetName),
      cadetNumber: collegeValidateCadetNumber(payload.cadetNumber),
      answers: collegeValidateAnswers(payload.collegeAnswers),
      submittedAt: new Date()
    };

    const spreadsheet = SpreadsheetApp.openById(COLLEGE_SPREADSHEET_ID);
    collegeSaveSubmission(spreadsheet, submission);
    SpreadsheetApp.flush();
    return collegeJsonResponse({ok: true, rowsAdded: 1});
  } catch (error) {
    console.error(error);
    return collegeJsonResponse({
      ok: false,
      error: "The response could not be saved. Check the College Apps Script execution log for details."
    });
  } finally {
    lock.releaseLock();
  }
}

function collegeSaveSubmission(spreadsheet, submission) {
  const sheet = collegeEnsureSheet(spreadsheet);
  const rowNumber = collegeGetNextResponseRow(sheet);
  const row = [
    rowNumber - 4,
    submission.submittedAt,
    submission.className,
    submission.houseName,
    submission.formName,
    collegeSafeCellText(submission.cadetName),
    collegeSafeCellText(submission.cadetNumber),
    ...COLLEGE_RATING_FIELDS.flatMap(field => [
      submission.answers.ratings[field.key],
      collegeSafeCellText(submission.answers.comments[field.key])
    ]),
    collegeSafeCellText(submission.answers.missionObjective),
    collegeSafeCellText(submission.answers.missionDescription),
    collegeSafeCellText(submission.answers.managementComment)
  ];
  sheet.getRange(rowNumber, 2, 1, row.length).setValues([row]);
}

function collegeEnsureSheet(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(COLLEGE_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(COLLEGE_SHEET_NAME);
  }

  const headers = collegeBuildHeaders();
  if (sheet.getLastRow() === 0) {
    sheet.getRange(2, 9, 1, 6).merge();
    sheet.getRange(2, 9).setValue(COLLEGE_SHEET_NAME);
    sheet.getRange(4, 2, 1, headers.length).setValues([headers]);
  } else {
    collegeMigrateOrValidateHeaders(sheet, headers);
    sheet.getRange(2, 9).setValue(COLLEGE_SHEET_NAME);
  }

  sheet.getRange(2, 9, 1, 6)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.getRange(4, 2, 1, headers.length)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);
  sheet.setFrozenRows(4);
  spreadsheet.setActiveSheet(sheet);
  spreadsheet.moveActiveSheet(spreadsheet.getSheets().length);
  return sheet;
}

function collegeBuildHeaders() {
  const headers = [
    "No",
    "Submit on",
    "Class",
    "House",
    "Form",
    "Cadet Name",
    "Cadet Number"
  ];
  COLLEGE_RATING_FIELDS.forEach(field => headers.push(field.header, `${field.header} Comment`));
  headers.push("College Mission Objective", "College Mission or Purpose", "College Management Comment");
  return headers;
}

function collegeMigrateOrValidateHeaders(sheet, headers) {
  if (sheet.getLastRow() < 4) {
    throw new Error("The existing College Evaluation sheet has an incomplete layout.");
  }

  const oldColumnCount = Math.max(sheet.getLastColumn() - 1, 1);
  const oldHeaders = sheet.getRange(4, 2, 1, oldColumnCount).getValues()[0];
  while (oldHeaders.length && oldHeaders[oldHeaders.length - 1] === "") {
    oldHeaders.pop();
  }
  if (headers.every((header, index) => oldHeaders[index] === header)) {
    return;
  }

  const basicHeaders = ["No", "Submit on", "Class", "House", "Form", "Cadet Name", "Cadet Number"];
  const hasBasicHeaders = basicHeaders.every((header, index) => oldHeaders[index] === header);
  const objectiveIndex = oldHeaders.indexOf("College Mission Objective");
  const missionIndex = oldHeaders.indexOf("College Mission or Purpose");
  const managementCommentIndex = oldHeaders.indexOf("College Management Comment");
  const knownRatingHeaders = new Set([
    ...COLLEGE_RATING_FIELDS.map(field => field.header),
    "College Environment Rating",
    "Education Development Measures Rating",
    "Cadet Dining and Food Quality Rating",
    "House Environment Rating",
    "Form Master Sincerity Rating",
    "House Master Sincerity and Dedication Rating",
    "Hospital Facilities and Environment Rating",
    "College Adjutant Sincerity Rating"
  ]);
  const ratingRangeEnd = missionIndex >= 0 ? missionIndex : oldHeaders.length;
  const ratingHeadersAreValid = oldHeaders
    .slice(basicHeaders.length, ratingRangeEnd)
    .every((header) => {
      if (header === "") {
        return true;
      }
      if (knownRatingHeaders.has(header)) {
        return true;
      }
      return header.endsWith(" Comment")
        && knownRatingHeaders.has(header.slice(0, -" Comment".length));
    });

  if (!hasBasicHeaders || missionIndex < basicHeaders.length || !ratingHeadersAreValid) {
    throw new Error("The existing College Evaluation sheet has incompatible headers.");
  }

  const lastRow = sheet.getLastRow();
  const oldRows = sheet.getRange(4, 2, lastRow - 3, oldColumnCount).getValues();
  const headerIndexes = new Map(oldHeaders.map((header, index) => [header, index]));
  const legacyHeaderByNewHeader = {
    "Your view about The Adjutant": "College Adjutant Sincerity Rating"
  };
  const migratedRows = oldRows.slice(1).map(oldRow => {
    const row = oldRow.slice(0, basicHeaders.length);
    COLLEGE_RATING_FIELDS.forEach(field => {
      const legacyHeader = legacyHeaderByNewHeader[field.header] || field.header;
      const ratingIndex = headerIndexes.has(field.header)
        ? headerIndexes.get(field.header)
        : headerIndexes.get(legacyHeader);
      const commentIndex = ratingIndex === undefined
        ? undefined
        : headerIndexes.get(`${oldHeaders[ratingIndex]} Comment`);
      row.push(
        ratingIndex === undefined ? "" : oldRow[ratingIndex],
        commentIndex === undefined ? "" : oldRow[commentIndex]
      );
    });
    row.push(
      objectiveIndex >= 0 ? oldRow[objectiveIndex] : "",
      missionIndex >= 0 ? oldRow[missionIndex] : "",
      managementCommentIndex >= 0 ? oldRow[managementCommentIndex] : ""
    );
    return row;
  });
  sheet.getRange(4, 2, 1, headers.length).setValues([headers]);
  if (migratedRows.length) {
    sheet.getRange(5, 2, migratedRows.length, headers.length).setValues(migratedRows);
  }
}

function collegeValidateAnswers(answers) {
  if (!answers
    || !answers.ratings
    || !answers.comments
    || typeof answers.missionObjective !== "string"
    || !COLLEGE_VALID_MISSION_OBJECTIVES.includes(answers.missionObjective)
    || typeof answers.missionDescription !== "string"
    || answers.missionDescription.length > 5000) {
    throw new Error("Invalid or incomplete College Evaluation answers.");
  }

  const ratings = {};
  const comments = {};
  const requiredRatingKeys = new Set(COLLEGE_REQUIRED_RATING_KEYS);
  COLLEGE_RATING_FIELDS.forEach(field => {
    const rating = answers.ratings[field.key];
    const comment = answers.comments[field.key] === undefined
      && !requiredRatingKeys.has(field.key)
        ? ""
        : answers.comments[field.key];
    if (requiredRatingKeys.has(field.key)
      && !COLLEGE_VALID_OPTION_VALUES.includes(rating)) {
      throw new Error(`Invalid or missing ${field.header}.`);
    }
    if (rating !== undefined
      && rating !== ""
      && !COLLEGE_VALID_OPTION_VALUES.includes(rating)) {
      throw new Error(`Invalid ${field.header}.`);
    }
    if (typeof comment !== "string"
      || comment.length > COLLEGE_MAX_COMMENT_LENGTH
      || (requiredRatingKeys.has(field.key) && rating <= 3 && !comment.trim())) {
      throw new Error(`Invalid comment for ${field.header}.`);
    }
    ratings[field.key] = rating === undefined ? "" : rating;
    comments[field.key] = comment.trim();
  });

  const missionDescription = answers.missionDescription.trim();
  if ((answers.missionObjective === "হ্যাঁ" || answers.missionObjective === "অন্যান্য") && !missionDescription) {
    throw new Error("Please describe the mission or purpose when the response is yes or other.");
  }

  return {
    ratings: ratings,
    comments: comments,
    missionObjective: answers.missionObjective.trim(),
    missionDescription: missionDescription,
    managementComment: collegeValidateOptionalText(
      answers.managementComment,
      "College Management Comment",
      COLLEGE_MAX_COMMENT_LENGTH
    )
  };
}

function collegeValidateChoice(value, allowedValues, fieldName) {
  if (!allowedValues.includes(value)) {
    throw new Error(`Invalid or missing ${fieldName}.`);
  }
  return value;
}

function collegeValidateOptionalText(value, fieldName, maxLength) {
  if (value === undefined || value === null || value === "") {
    return "";
  }
  if (typeof value !== "string" || value.length > maxLength) {
    throw new Error(`Invalid ${fieldName}.`);
  }
  return value.trim();
}

function collegeValidateCadetName(value) {
  const name = collegeValidateOptionalText(value, "Cadet Name", 200);
  if (name && !/^[\p{L}\p{M}]+(?:\s+[\p{L}\p{M}]+)*$/u.test(name)) {
    throw new Error("Cadet Name may contain letters and spaces only.");
  }
  return name;
}

function collegeValidateCadetNumber(value) {
  const number = collegeValidateOptionalText(value, "Cadet Number", 100);
  if (number && !/^\d+$/.test(number)) {
    throw new Error("Cadet Number may contain digits only.");
  }
  return number;
}

function collegeGetNextResponseRow(sheet) {
  const firstDataRow = 5;
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) {
    return firstDataRow;
  }

  const numbers = sheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, 1).getValues();
  for (let index = numbers.length - 1; index >= 0; index--) {
    if (numbers[index][0] !== "" && numbers[index][0] !== null) {
      return firstDataRow + index + 1;
    }
  }
  return firstDataRow;
}

function collegeSafeCellText(value) {
  return /^[\s]*[=+\-@]/.test(value) ? `'${value}` : value;
}

function collegeJsonResponse(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
