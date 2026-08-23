/**
 * ============================================================================
 * PDF_ENGINE.GS
 * Drive folder automation + HTML-to-PDF generation for Visit reports and
 * signed-checklist image-to-PDF conversion.
 *
 * ⚠️ REQUIRED ONE-TIME SETUP — Advanced Drive Service
 * Apps Script has no built-in "HTML string to PDF" function. The reliable,
 * native way to get styled (tables, colors, layout) PDFs out of Apps Script
 * is: convert an HTML string into a temporary Google Doc (which Google's
 * import pipeline renders faithfully, including inline base64 images), then
 * export that Doc as a PDF, then delete the temp Doc. This requires the
 * "Drive API" ADVANCED SERVICE (Drive API v2), which is not on by default:
 *   1. Apps Script editor → Services (+) → find "Drive API" → Add.
 *   2. It will link to a Google Cloud project — click through and make sure
 *      the Cloud project also has the "Google Drive API" enabled under
 *      "APIs & Services" (the Advanced Service add step usually does this
 *      automatically, but verify if you see "Drive is not defined" errors).
 * Without this step, `convertHtmlToPdfBlob()` below will throw a clear error
 * rather than silently failing.
 * ============================================================================
 */

/**
 * Get (or create) a folder by name inside a given parent folder.
 */
function getOrCreateFolder(parentFolder, name) {
  const existing = parentFolder.getFoldersByName(name);
  if (existing.hasNext()) {
    return existing.next();
  }
  return parentFolder.createFolder(name);
}

/**
 * Walks/creates: Tour Visit / {Year} / {MonthName} / {District} / {BusStation} / {subfolder}
 * subfolder is either 'Generated PDF' or 'Signed Checklist'.
 * Returns the final Folder object.
 */
function getVisitFolder(district, busStation, subfolder, visitDate) {
  const date = visitDate ? new Date(visitDate) : new Date();
  const year = String(date.getFullYear());
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  const monthName = monthNames[date.getMonth()];

  const root = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
  const yearFolder = getOrCreateFolder(root, year);
  const monthFolder = getOrCreateFolder(yearFolder, monthName);
  const districtFolder = getOrCreateFolder(monthFolder, district || 'Unknown District');
  const stationFolder = getOrCreateFolder(districtFolder, busStation || 'Unknown Station');
  return getOrCreateFolder(stationFolder, subfolder);
}

/**
 * Fetches a QR code PNG for the given text and returns it as a base64 string
 * (for embedding inline in the PDF's HTML). Returns null on failure so a
 * transient QR-service outage never blocks the whole PDF/visit submission.
 */
function fetchQrCodeBase64(data) {
  try {
    const url = CONFIG.QR_ENDPOINT + '?text=' + encodeURIComponent(data) + '&size=180&margin=1';
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('QR fetch non-200: ' + response.getResponseCode());
      return null;
    }
    return Utilities.base64Encode(response.getBlob().getBytes());
  } catch (error) {
    Logger.log('QR generation failed (continuing without QR): ' + error.toString());
    return null;
  }
}

/**
 * Converts an HTML string into a PDF Blob via a temporary Google Doc.
 * Cleans up the temporary Doc/Sheet-import file afterward regardless of outcome.
 */
function convertHtmlToPdfBlob(htmlString, tempFileName) {
  const htmlBlob = Utilities.newBlob(htmlString, MimeType.HTML, tempFileName + '.html');
  let tempFileId = null;

  try {
    if (typeof Drive === 'undefined') {
      throw new Error(
        'The Drive API advanced service is not enabled for this project. ' +
        'Open the Apps Script editor → Services (+) → add "Drive API", then try again.'
      );
    }

    // Convert HTML -> Google Doc using the Advanced Drive Service (v2)
    const resource = {
      title: tempFileName,
      mimeType: MimeType.GOOGLE_DOCS
    };
    const converted = Drive.Files.insert(resource, htmlBlob, { convert: true });
    tempFileId = converted.id;

    // Export the resulting Doc as PDF
    const pdfBlob = DriveApp.getFileById(tempFileId).getAs(MimeType.PDF);
    pdfBlob.setName(tempFileName + '.pdf');
    return pdfBlob;
  } finally {
    // Always clean up the temporary Doc, even if conversion/export failed
    if (tempFileId) {
      try {
        Drive.Files.remove(tempFileId);
      } catch (cleanupError) {
        Logger.log('Temp file cleanup failed for ' + tempFileId + ': ' + cleanupError.toString());
      }
    }
  }
}

/**
 * Builds the full HTML for a Visit report PDF.
 * @param {Object} visit - plain object with visit fields
 * @param {Array<Object>} responses - [{category, question, answer}]
 * @param {string|null} qrBase64
 */
function buildVisitPdfHtml(visit, responses, qrBase64) {
  const NAVY = '#1E315C';
  const YELLOW = '#F4B400';
  const BG = '#F8FAFC';

  const escapeHtml = function (str) {
    return String(str == null ? '' : str)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  };

  // Group responses by category for a readable table layout
  const byCategory = {};
  const order = [];
  responses.forEach(function (r) {
    const cat = r.category || 'सर्वसाधारण';
    if (!byCategory[cat]) { byCategory[cat] = []; order.push(cat); }
    byCategory[cat].push(r);
  });

  let checklistRowsHtml = '';
  order.forEach(function (cat) {
    checklistRowsHtml +=
      '<tr><td colspan="2" style="background:' + NAVY + ';color:#fff;font-weight:bold;padding:8px 10px;">' +
      escapeHtml(cat) + '</td></tr>';
    byCategory[cat].forEach(function (r) {
      const isYes = r.answer === 'होय';
      const badgeColor = isYes ? '#1a7f37' : '#c0392b';

      // Star rating line (only present for तपासावी/तपासाव्यात questions
      // answered होय) — printed as filled/empty stars plus the numeric score.
      let extraHtml = '';
      const ratingNum = parseInt(r.rating, 10);
      if (ratingNum >= 1 && ratingNum <= 5) {
        const stars = '★★★★★'.slice(0, ratingNum) + '☆☆☆☆☆'.slice(0, 5 - ratingNum);
        extraHtml += '<div style="margin-top:3px;color:' + YELLOW + ';font-size:13px;letter-spacing:1px;">' +
          stars + ' <span style="color:#64748b;font-size:10px;">(' + ratingNum + '/5)</span></div>';
      }
      // Ticked difficulty options (only present for the अडचणी question)
      if (r.difficultyOptions) {
        extraHtml += '<div style="margin-top:3px;color:#475569;font-size:10.5px;"><b>अडचणी:</b> ' + escapeHtml(r.difficultyOptions) + '</div>';
      }
      // Remark (only present when answer is नाही)
      if (r.remark) {
        extraHtml += '<div style="margin-top:3px;color:#475569;font-size:10.5px;"><b>शेरा:</b> ' + escapeHtml(r.remark) + '</div>';
      }

      checklistRowsHtml +=
        '<tr>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;width:80%;">' + escapeHtml(r.question) + extraHtml + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e2e8f0;text-align:center;vertical-align:top;">' +
        '<span style="color:' + badgeColor + ';font-weight:bold;">' + escapeHtml(r.answer) + '</span>' +
        '</td></tr>';
    });
  });

  const qrHtml = qrBase64
    ? '<img src="data:image/png;base64,' + qrBase64 + '" width="90" height="90" />'
    : '';

  return '' +
    '<html><head><meta charset="UTF-8"></head>' +
    '<body style="font-family:Arial, sans-serif; color:#1a1a1a; background:' + BG + ';">' +

    // Header
    '<table width="100%" style="border-bottom:3px solid ' + NAVY + ';padding-bottom:10px;margin-bottom:16px;">' +
    '<tr>' +
    '<td style="width:70px;"><img src="data:' + ASSETS.LOGO_MIME_TYPE + ';base64,' + ASSETS.LOGO_ICON_BASE64 + '" width="60" height="60" /></td>' +
    '<td style="vertical-align:middle;">' +
    '<div style="font-size:20px;font-weight:bold;color:' + NAVY + ';">SMART SERVICES</div>' +
    '<div style="font-size:11px;color:#64748b;">Tour Visit Checklist Report</div>' +
    '</td>' +
    '<td style="text-align:right;vertical-align:middle;">' +
    '<div style="font-size:11px;color:#64748b;">Visit ID</div>' +
    '<div style="font-size:14px;font-weight:bold;color:' + YELLOW + ';background:' + NAVY + ';padding:4px 10px;border-radius:4px;display:inline-block;">' +
    escapeHtml(visit.visitId) + '</div>' +
    '</td>' +
    '</tr></table>' +

    // Officer / Visit details
    '<table width="100%" style="border-collapse:collapse;margin-bottom:14px;font-size:12px;">' +
    '<tr>' +
    '<td style="width:50%;padding:4px 0;"><b>अधिकारी:</b> ' + escapeHtml(visit.employeeName) + ' (' + escapeHtml(visit.employeeId) + ')</td>' +
    '<td style="width:50%;padding:4px 0;"><b>पदनाम:</b> ' + escapeHtml(visit.designation) + '</td>' +
    '</tr><tr>' +
    '<td style="padding:4px 0;"><b>जिल्हा:</b> ' + escapeHtml(visit.district) + '</td>' +
    '<td style="padding:4px 0;"><b>बस स्टेशन:</b> ' + escapeHtml(visit.busStation) + '</td>' +
    '</tr><tr>' +
    '<td style="padding:4px 0;"><b>भेट दिनांक:</b> ' + escapeHtml(visit.visitDate) + '</td>' +
    '<td style="padding:4px 0;"><b>भेट वेळ:</b> ' + escapeHtml(visit.visitTime) + '</td>' +
    '</tr><tr>' +
    '<td style="padding:4px 0;" colspan="2"><b>GPS स्थान:</b> ' + escapeHtml(visit.latitude) + ', ' + escapeHtml(visit.longitude) +
    (visit.gpsAddress ? ' — ' + escapeHtml(visit.gpsAddress) : '') + '</td>' +
    '</tr></table>' +

    // Checklist table
    '<table width="100%" style="border-collapse:collapse;border:1px solid #e2e8f0;margin-bottom:14px;">' +
    checklistRowsHtml +
    '</table>' +

    // Overall remark
    '<table width="100%" style="margin-bottom:20px;">' +
    '<tr><td style="background:#fff;border:1px solid #e2e8f0;border-radius:6px;padding:10px;font-size:12px;">' +
    '<b>एकूण शेरा:</b><br/>' + (escapeHtml(visit.overallRemark) || '—') +
    '</td></tr></table>' +

    // Footer with QR
    '<table width="100%" style="border-top:1px solid #e2e8f0;padding-top:10px;">' +
    '<tr>' +
    '<td style="font-size:10px;color:#94a3b8;vertical-align:middle;">' +
    'Generated by Smart Services Tour Visit System<br/>' +
    escapeHtml(visit.createdTimestamp) +
    '</td>' +
    '<td style="text-align:right;">' + qrHtml + '</td>' +
    '</tr></table>' +

    '</body></html>';
}

/**
 * Orchestrates: build HTML -> convert to PDF -> save into the correct
 * Drive folder -> return the file's shareable URL.
 */
function generateAndStoreVisitPdf(visit, responses) {
  const qrBase64 = fetchQrCodeBase64(visit.visitId);
  const html = buildVisitPdfHtml(visit, responses, qrBase64);
  const pdfBlob = convertHtmlToPdfBlob(html, 'Visit_' + visit.visitId);

  const folder = getVisitFolder(visit.district, visit.busStation, 'Generated PDF', visit.visitDate);
  const file = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

/**
 * Converts an uploaded signed-checklist file to PDF (if it's an image) and
 * stores it in the correct Drive folder. Returns the shareable URL.
 * @param {string} base64Data
 * @param {string} mimeType - one of CONFIG.ALLOWED_SIGNED_CHECKLIST_TYPES
 * @param {Object} visit - needs district, busStation, visitDate, visitId
 */
function storeSignedChecklist(base64Data, mimeType, visit) {
  const rawBytes = Utilities.base64Decode(base64Data);
  let pdfBlob;

  if (mimeType === 'application/pdf') {
    pdfBlob = Utilities.newBlob(rawBytes, MimeType.PDF, 'SignedChecklist_' + visit.visitId + '.pdf');
  } else {
    // Image (jpg/png) -> convert to a single full-page PDF
    const imageBlob = Utilities.newBlob(rawBytes, mimeType, 'signed.' + (mimeType.indexOf('png') > -1 ? 'png' : 'jpg'));
    const imageBase64 = Utilities.base64Encode(imageBlob.getBytes());
    const html =
      '<html><body style="margin:0;padding:0;text-align:center;">' +
      '<img src="data:' + mimeType + ';base64,' + imageBase64 + '" style="width:100%;max-width:750px;" />' +
      '</body></html>';
    pdfBlob = convertHtmlToPdfBlob(html, 'SignedChecklist_' + visit.visitId);
  }

  const folder = getVisitFolder(visit.district, visit.busStation, 'Signed Checklist', visit.visitDate);
  const file = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}
