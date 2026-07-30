/**
 * ============================================================================
 * TOURVISIT_BACKEND.GS
 * Smart Services — Tour Visit Checklist System
 * Production REST API backend (Google Apps Script)
 * Version 2.0.0
 *
 * Depends on: Setup.gs (CONFIG/SHEETS/COLS), PDF_Engine.gs (PDF generation),
 * Assets.gs (embedded logo). All four files must live in the same Apps
 * Script project.
 * ============================================================================
 */

// ============================================================================
// MAIN REQUEST HANDLER
// ============================================================================

function doGet(e) {
  return handleRequest(e, 'GET');
}

function doPost(e) {
  return handleRequest(e, 'POST');
}

function handleRequest(e, method) {
  const action = (e && e.parameter && e.parameter.action) || '';
  let response;

  try {
    switch (action) {
      // AUTH
      case 'login': response = handleLogin(e); break;
      case 'logout': response = handleLogout(e); break;
      case 'changePassword': response = handleChangePassword(e); break;
      case 'validateSession': response = handleValidateSession(e); break;

      // MASTER DATA
      case 'districts': response = handleGetDistricts(); break;
      case 'busstations': response = handleGetBusStations(e.parameter.district); break;
      case 'checklist': response = handleGetChecklist(); break;
      case 'user': response = handleGetUser(e.parameter.token); break;

      // VISITS
      case 'submitVisit': response = requirePost(method, function () { return handleSubmitVisit(e); }); break;
      case 'getVisit': response = handleGetVisit(e.parameter.visitId, e.parameter.token); break;
      case 'getMyVisits': response = handleGetMyVisits(e.parameter.token); break;
      case 'getAllVisits': response = handleGetAllVisits(e); break;
      case 'regeneratePDF': response = requirePost(method, function () { return handleRegeneratePdf(e); }); break;

      // UPLOADS
      case 'uploadSignedChecklist': response = requirePost(method, function () { return handleUploadSignedChecklist(e); }); break;

      // DASHBOARD / REPORTS
      case 'dashboard': response = handleDashboard(e.parameter.token); break;
      case 'reports': response = handleGetAllVisits(e); break; // alias, same filterable data source

      default:
        response = { success: false, message: 'Invalid endpoint: ' + action, version: CONFIG.API_VERSION };
    }
  } catch (error) {
    Logger.log('Unhandled error [' + action + ']: ' + error.toString() + '\n' + (error.stack || ''));
    response = { success: false, message: 'Server error', error: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function requirePost(method, fn) {
  if (method !== 'POST') {
    return { success: false, message: 'POST required for this action' };
  }
  return fn();
}

// ============================================================================
// AUTHENTICATION
// ============================================================================

function handleLogin(e) {
  const employeeId = (e.parameter.employeeId || '').trim();
  const password = e.parameter.password || '';

  if (!employeeId || !password) {
    return { success: false, message: 'कर्मचारी आयडी आणि पासवर्ड आवश्यक आहे' };
  }

  const lockKey = 'lockout_' + employeeId;
  const cache = CacheService.getScriptCache();
  const attempts = parseInt(cache.get(lockKey) || '0', 10);
  if (attempts >= CONFIG.MAX_LOGIN_ATTEMPTS) {
    return { success: false, message: 'खूप प्रयत्न अयशस्वी. 15 मिनिटांनी पुन्हा प्रयत्न करा.' };
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  const data = usersSheet.getDataRange().getValues();

  let user = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COLS.USERS.EMPLOYEE_ID]).trim() === employeeId) {
      const storedPassword = String(data[i][COLS.USERS.PASSWORD]);
      if (storedPassword !== password) {
        cache.put(lockKey, String(attempts + 1), CONFIG.LOGIN_LOCKOUT_SECONDS);
        logAudit(employeeId, 'LOGIN_FAILED', e);
        return { success: false, message: 'अवैध कर्मचारी आयडी किंवा पासवर्ड' };
      }
      user = {
        employeeId: data[i][COLS.USERS.EMPLOYEE_ID],
        name: data[i][COLS.USERS.NAME],
        designation: data[i][COLS.USERS.DESIGNATION],
        role: data[i][COLS.USERS.ROLE],
        district: data[i][COLS.USERS.DISTRICT],
        status: data[i][COLS.USERS.STATUS]
      };
      break;
    }
  }

  if (!user) {
    cache.put(lockKey, String(attempts + 1), CONFIG.LOGIN_LOCKOUT_SECONDS);
    logAudit(employeeId, 'LOGIN_FAILED', e);
    // Same generic message as a wrong-password failure — don't reveal whether the ID exists
    return { success: false, message: 'अवैध कर्मचारी आयडी किंवा पासवर्ड' };
  }

  if (String(user.status).trim() !== 'Active') {
    return { success: false, message: 'खाते निष्क्रिय आहे. कृपया प्रशासकाशी संपर्क साधा.' };
  }

  cache.remove(lockKey);
  const token = generateToken(employeeId);
  saveSession(token, user);
  logAudit(employeeId, 'LOGIN_SUCCESS', e);

  return { success: true, message: 'लॉगिन यशस्वी', token: token, user: user };
}

function handleLogout(e) {
  const token = e.parameter.token;
  const session = getSession(token);
  if (session) {
    logAudit(session.user.employeeId, 'LOGOUT', e);
    deleteSession(token);
  }
  return { success: true, message: 'Logout successful' };
}

function handleChangePassword(e) {
  const token = e.parameter.token;
  const oldPassword = e.parameter.oldPassword || '';
  const newPassword = e.parameter.newPassword || '';

  const session = validateSession(token);
  if (!session) return { success: false, message: 'Unauthorized' };
  if (!newPassword || newPassword.length < 4) {
    return { success: false, message: 'नवीन पासवर्ड किमान 4 अक्षरांचा असावा' };
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const usersSheet = ss.getSheetByName(SHEETS.USERS);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = usersSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][COLS.USERS.EMPLOYEE_ID] === session.user.employeeId) {
        if (String(data[i][COLS.USERS.PASSWORD]) !== oldPassword) {
          return { success: false, message: 'जुना पासवर्ड चुकीचा आहे' };
        }
        usersSheet.getRange(i + 1, COLS.USERS.PASSWORD + 1).setValue(newPassword);
        logAudit(session.user.employeeId, 'PASSWORD_CHANGED', e);
        return { success: true, message: 'पासवर्ड यशस्वीरित्या बदलला' };
      }
    }
    return { success: false, message: 'User not found' };
  } finally {
    lock.releaseLock();
  }
}

function handleValidateSession(e) {
  const session = validateSession(e.parameter.token);
  if (!session) return { success: false, authenticated: false, message: 'Session expired' };
  return { success: true, authenticated: true, user: session.user };
}

// ============================================================================
// MASTER DATA (cached — Base/Checklist_Master change rarely)
// ============================================================================

function handleGetDistricts() {
  const cacheKey = 'districts';
  const cached = getFromCache(cacheKey);
  if (cached) return { success: true, data: cached };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.BASE).getDataRange().getValues();

  const districts = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLS.BASE.DISTRICT]) districts.add(String(data[i][COLS.BASE.DISTRICT]).trim());
  }
  const result = Array.from(districts).sort();
  putInCache(cacheKey, result);
  return { success: true, data: result };
}

function handleGetBusStations(district) {
  if (!district) return { success: false, message: 'District required' };

  const cacheKey = 'busstations_' + district;
  const cached = getFromCache(cacheKey);
  if (cached) return { success: true, data: cached };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.BASE).getDataRange().getValues();

  const stations = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COLS.BASE.DISTRICT]).trim() === district && data[i][COLS.BASE.BUS_STATION]) {
      stations.push(String(data[i][COLS.BASE.BUS_STATION]).trim());
    }
  }
  stations.sort();
  putInCache(cacheKey, stations);
  return { success: true, data: stations };
}

function handleGetChecklist() {
  const cacheKey = 'checklist';
  const cached = getFromCache(cacheKey);
  if (cached) return { success: true, data: cached };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.CHECKLIST_MASTER).getDataRange().getValues();

  const questions = [];
  for (let i = 1; i < data.length; i++) {
    const active = data[i][COLS.CHECKLIST_MASTER.ACTIVE];
    if (active === true || String(active).toUpperCase() === 'TRUE') {
      questions.push({
        questionId: data[i][COLS.CHECKLIST_MASTER.QUESTION_ID],
        category: data[i][COLS.CHECKLIST_MASTER.CATEGORY],
        question: data[i][COLS.CHECKLIST_MASTER.QUESTION],
        displayOrder: Number(data[i][COLS.CHECKLIST_MASTER.DISPLAY_ORDER]) || 0
      });
    }
  }
  questions.sort(function (a, b) { return a.displayOrder - b.displayOrder; });

  const categorized = {};
  const categoryOrder = [];
  questions.forEach(function (q) {
    if (!categorized[q.category]) { categorized[q.category] = []; categoryOrder.push(q.category); }
    categorized[q.category].push(q);
  });

  const result = { categories: categoryOrder, questions: categorized };
  putInCache(cacheKey, result);
  return { success: true, data: result };
}

function handleGetUser(token) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Unauthorized' };
  return { success: true, data: session.user };
}

// ============================================================================
// VISITS
// ============================================================================

/**
 * Single combined action: saves the visit + checklist responses, generates
 * the PDF, stores it in Drive, and writes the PDF URL back — matching the
 * spec's "click Submit → everything happens immediately" requirement.
 *
 * If PDF generation fails (e.g. transient QR/Drive issue), the visit data is
 * NOT lost — it's already saved before PDF generation is attempted. The
 * response reports pdfGenerated:false so the frontend can offer a retry via
 * the `regeneratePDF` action instead of forcing the officer to redo the form.
 */
function handleSubmitVisit(e) {
  const session = validateSession(e.parameter.token);
  if (!session) return { success: false, message: 'Unauthorized' };

  const district = e.parameter.district;
  const busStation = e.parameter.busStation;
  if (!district || !busStation) {
    return { success: false, message: 'जिल्हा आणि बस स्टेशन आवश्यक आहे' };
  }

  let responses;
  try {
    responses = JSON.parse(e.parameter.responses || '[]');
  } catch (parseError) {
    return { success: false, message: 'Invalid responses payload' };
  }

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const visitId = 'VIS-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss') + '-' + randomSuffix(4);

  const visit = {
    visitId: visitId,
    employeeId: session.user.employeeId,
    employeeName: session.user.name,
    designation: session.user.designation,
    district: district,
    busStation: busStation,
    visitDate: e.parameter.visitDate || Utilities.formatDate(now, tz, 'yyyy-MM-dd'),
    visitTime: e.parameter.visitTime || Utilities.formatDate(now, tz, 'HH:mm:ss'),
    latitude: e.parameter.latitude || '',
    longitude: e.parameter.longitude || '',
    gpsAddress: e.parameter.gpsAddress || '',
    overallRemark: e.parameter.overallRemark || '',
    createdTimestamp: Utilities.formatDate(now, tz, 'dd-MM-yyyy HH:mm:ss')
  };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const visitSheet = ss.getSheetByName(SHEETS.TOUR_VISIT);
  const responseSheet = ss.getSheetByName(SHEETS.CHECKLIST_RESPONSE);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    visitSheet.appendRow([
      visit.visitId, visit.employeeId, visit.employeeName, visit.designation,
      visit.district, visit.busStation, visit.visitDate, visit.visitTime,
      visit.latitude, visit.longitude, visit.gpsAddress, visit.overallRemark,
      '', '', now
    ]);

    // Force Visit Date / Visit Time to plain-text format immediately.
    // Sheets otherwise silently auto-converts strings that look like
    // dates/times into locale-formatted Date values on write — lock the
    // cell format so what's displayed always matches what was submitted.
    const newRow = visitSheet.getLastRow();
    visitSheet.getRange(newRow, COLS.TOUR_VISIT.VISIT_DATE + 1, 1, 2).setNumberFormat('@');

    const responseRows = responses.map(function (r) {
      return [visitId, r.questionId, r.question, r.answer, now];
    });
    if (responseRows.length > 0) {
      responseSheet.getRange(responseSheet.getLastRow() + 1, 1, responseRows.length, 5).setValues(responseRows);
    }
  } finally {
    lock.releaseLock();
  }

  logAudit(session.user.employeeId, 'VISIT_CREATED: ' + visitId, e);

  // PDF generation is attempted AFTER the data is safely saved above.
  let pdfUrl = '';
  let pdfError = '';
  try {
    pdfUrl = generateAndStoreVisitPdf(visit, responses);
    updateVisitCell(visitId, COLS.TOUR_VISIT.GENERATED_PDF_URL, pdfUrl);
    logAudit(session.user.employeeId, 'PDF_GENERATED: ' + visitId, e);
  } catch (pdfGenError) {
    pdfError = pdfGenError.toString();
    Logger.log('PDF generation failed for ' + visitId + ': ' + pdfError);
  }

  return {
    success: true,
    message: pdfUrl ? '✅ भेट यशस्वीरित्या जतन झाली. PDF तयार झाला आहे.' : '✅ भेट जतन झाली, पण PDF तयार करताना अडचण आली.',
    visitId: visitId,
    pdfGenerated: !!pdfUrl,
    pdfUrl: pdfUrl,
    pdfError: pdfError
  };
}

function handleRegeneratePdf(e) {
  const session = validateSession(e.parameter.token);
  if (!session) return { success: false, message: 'Unauthorized' };

  const visitId = e.parameter.visitId;
  const visitRow = findVisitRow(visitId);
  if (!visitRow) return { success: false, message: 'Visit not found' };

  const visit = rowToVisitObject(visitRow.values);
  if (session.user.role !== ROLES.ADMIN && session.user.employeeId !== visit.employeeId) {
    return { success: false, message: 'Unauthorized' };
  }

  const responses = getResponsesForVisit(visitId);
  try {
    const pdfUrl = generateAndStoreVisitPdf(visit, responses);
    updateVisitCell(visitId, COLS.TOUR_VISIT.GENERATED_PDF_URL, pdfUrl);
    logAudit(session.user.employeeId, 'PDF_REGENERATED: ' + visitId, e);
    return { success: true, pdfUrl: pdfUrl };
  } catch (error) {
    return { success: false, message: 'PDF generation failed: ' + error.toString() };
  }
}

function handleGetVisit(visitId, token) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Unauthorized' };

  const visitRow = findVisitRow(visitId);
  if (!visitRow) return { success: false, message: 'Visit not found' };

  const visit = rowToVisitObject(visitRow.values);
  if (session.user.role !== ROLES.ADMIN && session.user.employeeId !== visit.employeeId) {
    return { success: false, message: 'Unauthorized' };
  }

  return { success: true, visit: visit, responses: getResponsesForVisit(visitId) };
}

function handleGetMyVisits(token) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Unauthorized' };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.TOUR_VISIT).getDataRange().getValues();
  const visits = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][COLS.TOUR_VISIT.EMPLOYEE_ID] === session.user.employeeId) {
      visits.push(rowToVisitObject(data[i]));
    }
  }
  visits.reverse();
  return { success: true, data: visits };
}

/**
 * Org-wide visit search/listing for Admin dashboards & reports, with
 * search + filters. Non-admin roles are automatically scoped to their own
 * district regardless of the `district` filter they pass — this is a
 * server-side permission boundary, not just a UI convenience.
 */
function handleGetAllVisits(e) {
  const session = validateSession(e.parameter.token);
  if (!session) return { success: false, message: 'Unauthorized' };

  const search = (e.parameter.search || '').toLowerCase().trim();
  const dateFrom = e.parameter.dateFrom ? new Date(e.parameter.dateFrom) : null;
  const dateTo = e.parameter.dateTo ? new Date(e.parameter.dateTo) : null;
  const districtFilter = e.parameter.district || '';
  const busStationFilter = e.parameter.busStation || '';
  const roleFilter = e.parameter.role || '';

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const visitData = ss.getSheetByName(SHEETS.TOUR_VISIT).getDataRange().getValues();
  const userRoleByEmployeeId = getUserRoleMap();

  const results = [];
  for (let i = 1; i < visitData.length; i++) {
    const row = visitData[i];
    const visit = rowToVisitObject(row);

    // Permission boundary: non-admins only ever see their own district
    if (session.user.role !== ROLES.ADMIN && visit.district !== session.user.district) continue;

    if (districtFilter && visit.district !== districtFilter) continue;
    if (busStationFilter && visit.busStation !== busStationFilter) continue;

    const visitorRole = userRoleByEmployeeId[visit.employeeId] || '';
    if (roleFilter && visitorRole !== roleFilter) continue;

    if (dateFrom || dateTo) {
      const vDate = new Date(visit.visitDate);
      if (dateFrom && vDate < dateFrom) continue;
      if (dateTo && vDate > dateTo) continue;
    }

    if (search) {
      const haystack = (visit.visitId + ' ' + visit.employeeId + ' ' + visit.employeeName + ' ' + visit.busStation).toLowerCase();
      if (haystack.indexOf(search) === -1) continue;
    }

    visit.employeeRole = visitorRole;
    results.push(visit);
  }

  results.reverse();
  return { success: true, data: results, count: results.length };
}

// ============================================================================
// SIGNED CHECKLIST UPLOAD
// ============================================================================

function handleUploadSignedChecklist(e) {
  const session = validateSession(e.parameter.token);
  if (!session) return { success: false, message: 'Unauthorized' };

  const visitId = e.parameter.visitId;
  const fileData = e.parameter.fileData;
  const mimeType = (e.parameter.mimeType || '').toLowerCase();

  if (!visitId || !fileData) {
    return { success: false, message: 'visitId आणि file आवश्यक आहे' };
  }
  if (CONFIG.ALLOWED_SIGNED_CHECKLIST_TYPES.indexOf(mimeType) === -1) {
    return { success: false, message: 'फक्त PDF, JPG किंवा PNG फाईल अनुमत आहे' };
  }
  // Base64 inflates data by ~4/3 — validate before decoding to avoid
  // wasting compute on an oversized payload.
  const approxBytes = fileData.length * 0.75;
  if (approxBytes > CONFIG.MAX_UPLOAD_BYTES) {
    return { success: false, message: 'फाईल 10 MB पेक्षा मोठी आहे' };
  }

  const visitRow = findVisitRow(visitId);
  if (!visitRow) return { success: false, message: 'Visit not found' };
  const visit = rowToVisitObject(visitRow.values);

  if (session.user.role !== ROLES.ADMIN && session.user.employeeId !== visit.employeeId) {
    return { success: false, message: 'Unauthorized' };
  }

  try {
    const url = storeSignedChecklist(fileData, mimeType, visit);
    updateVisitCell(visitId, COLS.TOUR_VISIT.SIGNED_CHECKLIST_URL, url);
    logAudit(session.user.employeeId, 'SIGNED_CHECKLIST_UPLOADED: ' + visitId, e);
    return { success: true, message: 'सही केलेली चेकलिस्ट अपलोड झाली', url: url };
  } catch (error) {
    return { success: false, message: 'Upload failed: ' + error.toString() };
  }
}

// ============================================================================
// DASHBOARD
// ============================================================================

function handleDashboard(token) {
  const session = validateSession(token);
  if (!session) return { success: false, message: 'Unauthorized' };

  const scopeKey = session.user.role === ROLES.ADMIN ? 'ALL' : session.user.district;
  const cacheKey = 'dashboard_' + scopeKey;
  const cached = getFromCache(cacheKey); // short TTL: aggregate stats, refreshed every 2 minutes
  if (cached) return { success: true, dashboard: cached };

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const visitData = ss.getSheetByName(SHEETS.TOUR_VISIT).getDataRange().getValues();
  const userRoleByEmployeeId = getUserRoleMap();
  const userNameByEmployeeId = getUserNameMap();

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');

  let totalVisits = 0, todayVisits = 0, pendingSigned = 0, completedSigned = 0;
  const districtWise = {};
  const roleWise = { PM: {}, DM: {}, ADM: {} };
  const monthlyTrend = {};

  for (let i = 1; i < visitData.length; i++) {
    const row = visitData[i];
    const district = row[COLS.TOUR_VISIT.DISTRICT];

    if (session.user.role !== ROLES.ADMIN && district !== session.user.district) continue;

    totalVisits++;
    if (String(row[COLS.TOUR_VISIT.VISIT_DATE]) === todayStr) todayVisits++;

    const signedUrl = row[COLS.TOUR_VISIT.SIGNED_CHECKLIST_URL];
    if (signedUrl) completedSigned++; else pendingSigned++;

    districtWise[district] = (districtWise[district] || 0) + 1;

    const empId = row[COLS.TOUR_VISIT.EMPLOYEE_ID];
    const role = userRoleByEmployeeId[empId];
    const name = userNameByEmployeeId[empId] || empId;
    if (roleWise[role]) {
      roleWise[role][name] = (roleWise[role][name] || 0) + 1;
    }

    const createdTs = row[COLS.TOUR_VISIT.CREATED_TIMESTAMP];
    if (createdTs) {
      const d = new Date(createdTs);
      const monthKey = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      monthlyTrend[monthKey] = (monthlyTrend[monthKey] || 0) + 1;
    }
  }

  const dashboard = {
    cards: {
      totalVisits: totalVisits,
      todayVisits: todayVisits,
      monthlyVisits: monthlyTrend[Utilities.formatDate(today, tz, 'yyyy-MM')] || 0,
      pendingSignedChecklist: pendingSigned,
      completedVisits: completedSigned
    },
    charts: {
      districtWise: districtWise,
      pmWise: roleWise.PM,
      dmWise: roleWise.DM,
      admWise: roleWise.ADM,
      monthlyTrend: monthlyTrend
    },
    scope: session.user.role === ROLES.ADMIN ? 'all' : session.user.district
  };

  putInCache(cacheKey, dashboard, 120);
  return { success: true, dashboard: dashboard };
}

// ============================================================================
// SHEET HELPERS
// ============================================================================

function findVisitRow(visitId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.TOUR_VISIT);
  if (sheet.getLastRow() < 2) return null; // header row only, no visits yet

  const finder = sheet.getRange('A2:A' + sheet.getLastRow())
    .createTextFinder(visitId).matchEntireCell(true).findNext();
  if (!finder) return null;

  const rowIndex = finder.getRow();
  const values = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  return { rowIndex: rowIndex, values: values };
}

function updateVisitCell(visitId, colIndex, value) {
  const visitRow = findVisitRow(visitId);
  if (!visitRow) return false;
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ss.getSheetByName(SHEETS.TOUR_VISIT).getRange(visitRow.rowIndex, colIndex + 1).setValue(value);
  return true;
}

function rowToVisitObject(row) {
  const c = COLS.TOUR_VISIT;
  return {
    visitId: row[c.VISIT_ID],
    employeeId: row[c.EMPLOYEE_ID],
    employeeName: row[c.EMPLOYEE_NAME],
    designation: row[c.DESIGNATION],
    district: row[c.DISTRICT],
    busStation: row[c.BUS_STATION],
    visitDate: row[c.VISIT_DATE] instanceof Date
      ? Utilities.formatDate(row[c.VISIT_DATE], Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : row[c.VISIT_DATE],
    visitTime: row[c.VISIT_TIME],
    latitude: row[c.LATITUDE],
    longitude: row[c.LONGITUDE],
    gpsAddress: row[c.GPS_ADDRESS],
    overallRemark: row[c.OVERALL_REMARK],
    generatedPdfUrl: row[c.GENERATED_PDF_URL],
    signedChecklistUrl: row[c.SIGNED_CHECKLIST_URL],
    createdTimestamp: row[c.CREATED_TIMESTAMP] instanceof Date
      ? Utilities.formatDate(row[c.CREATED_TIMESTAMP], Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss')
      : row[c.CREATED_TIMESTAMP]
  };
}

function getResponsesForVisit(visitId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.CHECKLIST_RESPONSE).getDataRange().getValues();
  const categoryByQuestionId = getQuestionCategoryMap();

  const responses = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][COLS.CHECKLIST_RESPONSE.VISIT_ID] === visitId) {
      const questionId = data[i][COLS.CHECKLIST_RESPONSE.QUESTION_ID];
      responses.push({
        questionId: questionId,
        question: data[i][COLS.CHECKLIST_RESPONSE.QUESTION],
        answer: data[i][COLS.CHECKLIST_RESPONSE.ANSWER],
        // Checklist_Response doesn't store category (matches spec's schema) —
        // joined back from Checklist_Master here so regenerated PDFs group
        // questions identically to the original submission.
        category: categoryByQuestionId[questionId] || ''
      });
    }
  }
  return responses;
}

function getQuestionCategoryMap() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.CHECKLIST_MASTER).getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    map[data[i][COLS.CHECKLIST_MASTER.QUESTION_ID]] = data[i][COLS.CHECKLIST_MASTER.CATEGORY];
  }
  return map;
}

function getUserRoleMap() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.USERS).getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    map[data[i][COLS.USERS.EMPLOYEE_ID]] = data[i][COLS.USERS.ROLE];
  }
  return map;
}

function getUserNameMap() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const data = ss.getSheetByName(SHEETS.USERS).getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    map[data[i][COLS.USERS.EMPLOYEE_ID]] = data[i][COLS.USERS.NAME];
  }
  return map;
}

// ============================================================================
// SESSION / TOKEN HELPERS (ScriptCache — deterministic regardless of caller
// identity, unlike UserCache in an anonymously-accessed web app)
// ============================================================================

function generateToken(employeeId) {
  const raw = employeeId + ':' + new Date().getTime() + ':' + Utilities.getUuid();
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw));
}

function saveSession(token, user) {
  CacheService.getScriptCache().put(
    'session_' + token,
    JSON.stringify({ user: user, timestamp: new Date().getTime() }),
    CONFIG.SESSION_TIMEOUT_SECONDS
  );
}

function getSession(token) {
  if (!token) return null;
  const data = CacheService.getScriptCache().get('session_' + token);
  return data ? JSON.parse(data) : null;
}

function deleteSession(token) {
  CacheService.getScriptCache().remove('session_' + token);
}

function validateSession(token) {
  return getSession(token); // ScriptCache already enforces the TTL as absolute expiry
}

function randomSuffix(length) {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // no ambiguous 0/O/1/I
  let out = '';
  for (let i = 0; i < length; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

// ============================================================================
// GENERIC CACHE HELPERS (for rarely-changing master data)
// ============================================================================

function getFromCache(key) {
  const data = CacheService.getScriptCache().get('data_' + key);
  return data ? JSON.parse(data) : null;
}

function putInCache(key, value, ttlSeconds) {
  CacheService.getScriptCache().put('data_' + key, JSON.stringify(value), ttlSeconds || 21600);
}

// ============================================================================
// AUDIT LOG
// Note: Apps Script's `e` event object does not expose caller IP or
// User-Agent server-side — those fields are captured client-side (see
// api.js) and passed as `device`/`browser` params where available.
// ============================================================================

function logAudit(employeeId, action, e) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEETS.AUDIT_LOG);
    const device = (e && e.parameter && e.parameter.device) || 'Unknown';
    const browser = (e && e.parameter && e.parameter.browser) || 'Unknown';
    sheet.appendRow([new Date(), employeeId, action, device, browser]);
  } catch (error) {
    Logger.log('Audit logging error: ' + error.toString());
  }
}
