// ==================== BOOKINGS -> MEDIA DELIVERY SYNC ====================
// Paste this whole file into Apps Script as a new script file, then add this
// line inside your existing onOpen() menu (wherever you'd like it to appear):
//
//   .addItem('BOOKINGS - Send Ready Projects to Media Delivery', 'sendReadyBookingsToMediaDelivery')
//
// Example:
//   ui.createMenu('Admin Menu')
//     .addItem('BOOKINGS - Email Projects to Contractors', 'sendProjectsToContractors')
//     .addItem('BOOKINGS - Send Ready Projects to Media Delivery', 'sendReadyBookingsToMediaDelivery')
//     .addItem('MD - Send Projects to Editors', 'sendProjectsToDistribution')
//     ...
//
// One-time setup: run setupBookingsStatusDropdown() once (from the Apps
// Script editor, no menu needed) to add the Ready/Sent dropdown to column T.
//
// REQUIRED SETUP: column N links are Google Sheets "smart chips" (linked
// Drive file chips), not plain hyperlinks — Apps Script's normal
// RichTextValue API cannot read a chip's underlying URL. To read it
// reliably, this script uses the Sheets Advanced Service instead:
//   In the Apps Script editor -> Services (+ icon) -> add "Google Sheets
//   API" -> Add. That's it, no extra config needed.
// ===========================================================================

var MD_SYNC_CONFIG = {
  BOOKINGS_SHEET_NAME: 'Bookings',
  MEDIA_DELIVERY_SHEET_NAME: 'Media Delivery',

  // Bookings columns (1-based)
  COL_ZONE: 4,               // D
  COL_PROJECT_NAME: 5,       // E
  COL_EVENT_DATE: 8,         // H
  COL_PROJECT_TYPE: 10,      // J
  COL_PACKAGE_ADS: 11,       // K
  COL_LINK: 14,              // N
  COL_LEAD_PHOTOGRAPHER: 15, // O
  COL_SECOND_PHOTOGRAPHER: 16, // P
  COL_LEAD_VIDEOGRAPHER: 17, // Q
  COL_SECOND_VIDEOGRAPHER: 18, // R
  COL_STATUS: 20,            // T - free column used to track send status

  BOOKINGS_LAST_COL: 20,     // read columns A:T

  // Media Delivery columns (1-based)
  MD_COL_TIMESTAMP: 1,   // A
  MD_COL_FROM_EMAIL: 2,  // B
  MD_COL_ZONE: 3,        // C
  MD_COL_COMPANY: 4,     // D
  MD_COL_TYPE: 5,        // E
  MD_COL_PROJECT_NAME: 6,// F
  MD_COL_EVENT_DATE: 7,  // G
  MD_COL_PACKAGE_ADS: 8, // H
  MD_COL_LINK: 9,        // I
  MD_COL_LEAD_CREATIVE: 12,  // L
  MD_COL_SECOND_CREATIVE: 13,// M

  MD_LAST_COL: 13, // A:M

  FROM_EMAIL: 'contact@bluebelleweddings.com',
  COMPANY_NAME: 'Blue Belle Weddings',

  STATUS_READY: 'Ready to Send',
  STATUS_SENT: 'Sent to Media Delivery',

  COLOR_PHOTO_BG: '#FFE599',  // yellow
  COLOR_VIDEO_BG: '#9FC5E8',  // blue
  COLOR_READY_BG: '#F4CCCC',  // soft pink
  COLOR_SENT_BG: '#B6D7A8',   // soft green

  // Internal type keys stay 'Photo'/'Video' (used for lead/second creative
  // selection and background color); these are just the text written into
  // Media Delivery column E.
  TYPE_LABELS: { Photo: 'photos', Video: 'video' }
};

/**
 * Menu entry point. Copies every "Ready to Send" row from Bookings into
 * Media Delivery, then marks the source row as sent.
 */
function sendReadyBookingsToMediaDelivery() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookingsSheet = ss.getSheetByName(MD_SYNC_CONFIG.BOOKINGS_SHEET_NAME);
  var mdSheet = ss.getSheetByName(MD_SYNC_CONFIG.MEDIA_DELIVERY_SHEET_NAME);

  if (!bookingsSheet || !mdSheet) {
    SpreadsheetApp.getUi().alert('Could not find "' + MD_SYNC_CONFIG.BOOKINGS_SHEET_NAME +
      '" and/or "' + MD_SYNC_CONFIG.MEDIA_DELIVERY_SHEET_NAME + '" sheets.');
    return;
  }

  if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets) {
    SpreadsheetApp.getUi().alert('The "Google Sheets API" advanced service is not enabled.\n\n' +
      'In the Apps Script editor: Services (+ icon) -> add "Google Sheets API" -> Add. ' +
      'This is required to read the real link out of the smart-chip cells in column N.');
    return;
  }

  var lastRow = bookingsSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No booking rows found.');
    return;
  }

  var numRows = lastRow - 1;
  var range = bookingsSheet.getRange(2, 1, numRows, MD_SYNC_CONFIG.BOOKINGS_LAST_COL);
  var values = range.getValues();
  var linkUrls = getBookingsLinkUrls_(ss, bookingsSheet.getSheetName(), numRows);

  var timestamp = new Date();
  var newRowsData = [];   // array of {values: [...13], type: 'Photo'|'Video', linkUrl: string}
  var bookingsRowsToMark = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var sheetRow = i + 2;
    var status = String(row[MD_SYNC_CONFIG.COL_STATUS - 1] || '').trim().toLowerCase();

    if (status !== MD_SYNC_CONFIG.STATUS_READY.toLowerCase()) {
      continue;
    }

    var projectType = String(row[MD_SYNC_CONFIG.COL_PROJECT_TYPE - 1] || '').trim();
    var typesToCreate = [];
    if (projectType === 'Photography') {
      typesToCreate = ['Photo'];
    } else if (projectType === 'Videography') {
      typesToCreate = ['Video'];
    } else if (projectType === 'Photography + Videography') {
      typesToCreate = ['Photo', 'Video'];
    } else {
      Logger.log('Skipping Bookings row ' + sheetRow + ': unrecognized project type "' + projectType + '"');
      continue;
    }

    var linkUrl = linkUrls[i] || '';

    for (var t = 0; t < typesToCreate.length; t++) {
      var type = typesToCreate[t];
      var leadCreative = type === 'Photo'
        ? row[MD_SYNC_CONFIG.COL_LEAD_PHOTOGRAPHER - 1]
        : row[MD_SYNC_CONFIG.COL_LEAD_VIDEOGRAPHER - 1];
      var secondCreative = type === 'Photo'
        ? row[MD_SYNC_CONFIG.COL_SECOND_PHOTOGRAPHER - 1]
        : row[MD_SYNC_CONFIG.COL_SECOND_VIDEOGRAPHER - 1];

      var mdRow = new Array(MD_SYNC_CONFIG.MD_LAST_COL).fill('');
      mdRow[MD_SYNC_CONFIG.MD_COL_TIMESTAMP - 1] = timestamp;
      mdRow[MD_SYNC_CONFIG.MD_COL_FROM_EMAIL - 1] = MD_SYNC_CONFIG.FROM_EMAIL;
      mdRow[MD_SYNC_CONFIG.MD_COL_ZONE - 1] = row[MD_SYNC_CONFIG.COL_ZONE - 1];
      mdRow[MD_SYNC_CONFIG.MD_COL_COMPANY - 1] = MD_SYNC_CONFIG.COMPANY_NAME;
      mdRow[MD_SYNC_CONFIG.MD_COL_TYPE - 1] = MD_SYNC_CONFIG.TYPE_LABELS[type];
      mdRow[MD_SYNC_CONFIG.MD_COL_PROJECT_NAME - 1] = row[MD_SYNC_CONFIG.COL_PROJECT_NAME - 1];
      mdRow[MD_SYNC_CONFIG.MD_COL_EVENT_DATE - 1] = row[MD_SYNC_CONFIG.COL_EVENT_DATE - 1];
      mdRow[MD_SYNC_CONFIG.MD_COL_PACKAGE_ADS - 1] = row[MD_SYNC_CONFIG.COL_PACKAGE_ADS - 1];
      mdRow[MD_SYNC_CONFIG.MD_COL_LINK - 1] = linkUrl;
      mdRow[MD_SYNC_CONFIG.MD_COL_LEAD_CREATIVE - 1] = leadCreative;
      mdRow[MD_SYNC_CONFIG.MD_COL_SECOND_CREATIVE - 1] = secondCreative;

      newRowsData.push({ values: mdRow, type: type, linkUrl: linkUrl });
    }

    bookingsRowsToMark.push(sheetRow);
  }

  if (newRowsData.length === 0) {
    SpreadsheetApp.getUi().alert('No rows marked "' + MD_SYNC_CONFIG.STATUS_READY + '" were found.');
    return;
  }

  var startRow = getFirstEmptyRow_(mdSheet, MD_SYNC_CONFIG.MD_LAST_COL);
  var outputValues = newRowsData.map(function (r) { return r.values; });
  var targetRange = mdSheet.getRange(startRow, 1, outputValues.length, MD_SYNC_CONFIG.MD_LAST_COL);
  targetRange.setValues(outputValues);

  var typeBackgrounds = newRowsData.map(function (r) {
    return [r.type === 'Photo' ? MD_SYNC_CONFIG.COLOR_PHOTO_BG : MD_SYNC_CONFIG.COLOR_VIDEO_BG];
  });
  mdSheet.getRange(startRow, MD_SYNC_CONFIG.MD_COL_TYPE, newRowsData.length, 1).setBackgrounds(typeBackgrounds);

  for (var r = 0; r < newRowsData.length; r++) {
    if (newRowsData[r].linkUrl) {
      var linkCell = mdSheet.getRange(startRow + r, MD_SYNC_CONFIG.MD_COL_LINK);
      var rtv = SpreadsheetApp.newRichTextValue()
        .setText(newRowsData[r].linkUrl)
        .setLinkUrl(newRowsData[r].linkUrl)
        .build();
      linkCell.setRichTextValue(rtv);
    }
  }

  for (var b = 0; b < bookingsRowsToMark.length; b++) {
    var markRow = bookingsRowsToMark[b];
    var statusCell = bookingsSheet.getRange(markRow, MD_SYNC_CONFIG.COL_STATUS);
    statusCell.setValue(MD_SYNC_CONFIG.STATUS_SENT);
    statusCell.setBackground(MD_SYNC_CONFIG.COLOR_SENT_BG);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('Sent ' + bookingsRowsToMark.length + ' booking(s) to Media Delivery (' +
    newRowsData.length + ' row(s) created).');
}

/**
 * Finds the first empty row directly after the last row with data in the
 * given sheet, checking only the first `numCols` columns. Avoids appending
 * at the sheet's absolute bottom when there are trailing blank rows.
 */
function getFirstEmptyRow_(sheet, numCols) {
  var lastRow = sheet.getLastRow();
  if (lastRow === 0) return 1;

  var values = sheet.getRange(1, 1, lastRow, numCols).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var hasData = values[i].some(function (cell) {
      return cell !== '' && cell !== null;
    });
    if (hasData) return i + 2;
  }
  return 1;
}

/**
 * Batch-reads the real URL behind every cell in the Bookings link column
 * (column N), using the Sheets Advanced Service's `hyperlink` field. This
 * correctly resolves plain auto-linked URLs, manually-applied rich text
 * hyperlinks, =HYPERLINK(url, text) formulas, AND smart chips (linked Drive
 * file chips) — none of which Apps Script's built-in RichTextValue API can
 * reliably read the URL from, especially chips.
 *
 * Returns an array of length numRows, aligned to Bookings rows 2..lastRow
 * (index 0 = row 2), with '' for cells that have no link.
 */
function getBookingsLinkUrls_(ss, sheetName, numRows) {
  var colLetter = columnToLetter_(MD_SYNC_CONFIG.COL_LINK);
  var a1Range = "'" + sheetName + "'!" + colLetter + '2:' + colLetter + (numRows + 1);

  // includeGridData must be explicit — Spreadsheets.get returns ONLY
  // spreadsheet-level metadata (no cell data at all) unless this is true,
  // even when `ranges` is set. The `fields` mask keeps the response small
  // across ~thousands of rows by skipping per-cell formatting (confirmed via
  // debugInspectBookingsLinkCell that the real URL lives under
  // chipRuns[].chip.richLinkProperties.uri for this sheet's smart chips).
  var response = Sheets.Spreadsheets.get(ss.getId(), {
    ranges: [a1Range],
    includeGridData: true,
    fields: 'sheets.data.rowData.values(hyperlink,formattedValue,' +
      'chipRuns.chip.richLinkProperties.uri,textFormatRuns.format.link.uri)'
  });

  var urls = [];
  var rowData = response.sheets && response.sheets[0].data && response.sheets[0].data[0].rowData;

  for (var i = 0; i < numRows; i++) {
    var cell = rowData && rowData[i] && rowData[i].values && rowData[i].values[0];
    urls.push(extractCellUrl_(cell));
  }
  return urls;
}

/**
 * Pulls the real URL out of a Sheets API CellData object, trying every
 * place a link can live: the classic `hyperlink` field (manual links,
 * auto-detected URLs, =HYPERLINK() formulas), a smart chip's
 * richLinkProperties.uri, or (last resort) a formattedValue that is itself
 * a bare URL.
 */
function extractCellUrl_(cell) {
  if (!cell) return '';

  if (cell.hyperlink) return cell.hyperlink;

  if (cell.chipRuns && cell.chipRuns.length) {
    for (var i = 0; i < cell.chipRuns.length; i++) {
      var chip = cell.chipRuns[i].chip;
      var uri = chip && chip.richLinkProperties && chip.richLinkProperties.uri;
      if (uri) return uri;
    }
  }

  if (cell.textFormatRuns && cell.textFormatRuns.length) {
    for (var j = 0; j < cell.textFormatRuns.length; j++) {
      var link = cell.textFormatRuns[j].format && cell.textFormatRuns[j].format.link;
      if (link && link.uri) return link.uri;
    }
  }

  if (cell.formattedValue && /^https?:\/\//i.test(cell.formattedValue)) {
    return cell.formattedValue;
  }

  return '';
}

/**
 * ONE-TIME SETUP: adds a dropdown (data validation) to Bookings column T so
 * status can only be "Ready to Send" or "Sent to Media Delivery", and adds
 * conditional formatting so the cell auto-colors soft pink / soft green
 * based on whichever value is selected. Run this once from the Apps Script
 * editor (select this function, click Run) — it doesn't need to be wired to
 * any menu.
 */
function setupBookingsStatusDropdown() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookingsSheet = ss.getSheetByName(MD_SYNC_CONFIG.BOOKINGS_SHEET_NAME);
  if (!bookingsSheet) {
    SpreadsheetApp.getUi().alert('Could not find "' + MD_SYNC_CONFIG.BOOKINGS_SHEET_NAME + '" sheet.');
    return;
  }

  var numRows = bookingsSheet.getMaxRows() - 1; // exclude header row
  var range = bookingsSheet.getRange(2, MD_SYNC_CONFIG.COL_STATUS, numRows, 1);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([MD_SYNC_CONFIG.STATUS_READY, MD_SYNC_CONFIG.STATUS_SENT], true)
    .setAllowInvalid(false)
    .build();

  range.setDataValidation(rule);

  // Replace any conditional format rules already covering this exact range
  // (keeps this function safe to re-run without piling up duplicate rules).
  var existingRules = bookingsSheet.getConditionalFormatRules();
  var otherRules = existingRules.filter(function (r) {
    return !r.getRanges().some(function (rr) {
      return rr.getColumn() === range.getColumn() && rr.getRow() === range.getRow() &&
        rr.getNumRows() === range.getNumRows() && rr.getNumColumns() === range.getNumColumns();
    });
  });

  var readyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(MD_SYNC_CONFIG.STATUS_READY)
    .setBackground(MD_SYNC_CONFIG.COLOR_READY_BG)
    .setRanges([range])
    .build();

  var sentRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(MD_SYNC_CONFIG.STATUS_SENT)
    .setBackground(MD_SYNC_CONFIG.COLOR_SENT_BG)
    .setRanges([range])
    .build();

  otherRules.push(readyRule, sentRule);
  bookingsSheet.setConditionalFormatRules(otherRules);

  SpreadsheetApp.getUi().alert('Dropdown added to column ' + columnToLetter_(MD_SYNC_CONFIG.COL_STATUS) +
    ' (rows 2-' + bookingsSheet.getMaxRows() + '): "' + MD_SYNC_CONFIG.STATUS_READY + '" (soft pink) and "' +
    MD_SYNC_CONFIG.STATUS_SENT + '" (soft green).');
}

/**
 * Converts a 1-based column index to its A1 letter (e.g. 20 -> "T").
 */
function columnToLetter_(column) {
  var letter = '';
  while (column > 0) {
    var remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}
