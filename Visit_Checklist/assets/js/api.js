/**
 * ============================================================================
 * API.JS
 * Centralized communication with the Google Apps Script backend.
 *
 * All calls use POST (form-encoded body), including read-only ones. This is
 * a deliberate choice: session tokens are bearer credentials, and sending
 * them as GET query-string parameters would leak them into Apps Script
 * execution logs and browser history. POST keeps them in the request body.
 * ============================================================================
 */
import { API_CONFIG } from './config.js';
import { getToken } from './auth.js';

function deviceInfo() {
  return {
    device: (navigator.platform || 'Unknown') + (/Mobi/i.test(navigator.userAgent) ? ' (Mobile)' : ''),
    browser: navigator.userAgent || 'Unknown'
  };
}

export async function apiCall(action, params = {}) {
  const token = getToken();
  const body = { ...params, action };
  if (token && action !== 'login') {
    body.token = token;
  }
  if (['login', 'logout', 'submitVisit', 'uploadSignedChecklist'].includes(action)) {
    Object.assign(body, deviceInfo());
  }

  const bodyParams = new URLSearchParams();
  Object.keys(body).forEach((key) => {
    if (body[key] !== undefined && body[key] !== null) {
      bodyParams.append(key, body[key]);
    }
  });

  const options = {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded;charset=UTF-8' },
    body: bodyParams.toString()
  };

  try {
    return await requestWithRetry(options, 0);
  } catch (error) {
    console.error('API call error:', action, error);
    return { success: false, message: 'नेटवर्क त्रुटी: ' + error.message, error: error.message };
  }
}

async function requestWithRetry(options, attempt) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), API_CONFIG.TIMEOUT_MS);

  try {
    const response = await fetch(API_CONFIG.BASE_URL, { ...options, signal: controller.signal });
    clearTimeout(timeoutId);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return await response.json();
  } catch (error) {
    clearTimeout(timeoutId);
    if (attempt < API_CONFIG.RETRY_ATTEMPTS - 1) {
      await new Promise((r) => setTimeout(r, API_CONFIG.RETRY_DELAY_MS));
      return requestWithRetry(options, attempt + 1);
    }
    throw error;
  }
}

// ---- Auth ----
export const login = (employeeId, password) => apiCall('login', { employeeId, password });
export const logout = () => apiCall('logout', {});
export const validateSession = () => apiCall('validateSession', {});
export const changePassword = (oldPassword, newPassword) => apiCall('changePassword', { oldPassword, newPassword });

// ---- Master data ----
export const getDistricts = () => apiCall('districts', {});
export const getBusStations = (district) => apiCall('busstations', { district });
export const getChecklist = () => apiCall('checklist', {});
export const getUser = () => apiCall('user', {});

// ---- Visits ----
export const submitVisit = (visitData, responses) =>
  apiCall('submitVisit', { ...visitData, responses: JSON.stringify(responses) });
export const getVisit = (visitId) => apiCall('getVisit', { visitId });
export const getMyVisits = () => apiCall('getMyVisits', {});
export const getAllVisits = (filters = {}) => apiCall('getAllVisits', filters);
export const regeneratePDF = (visitId) => apiCall('regeneratePDF', { visitId });

// ---- Uploads ----
export const uploadSignedChecklist = (visitId, fileData, fileName, mimeType) =>
  apiCall('uploadSignedChecklist', { visitId, fileData, fileName, mimeType });

// ---- Dashboard ----
export const getDashboard = () => apiCall('dashboard', {});
