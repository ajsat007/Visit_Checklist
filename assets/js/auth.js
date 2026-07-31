/**
 * ============================================================================
 * AUTH.JS
 * Token and session management.
 * ============================================================================
 */
import { SESSION_CONFIG } from './config.js';

const KEYS = {
  TOKEN: 'ss_token',
  USER: 'ss_user',
  EXPIRY: 'ss_session_expiry'
};

export function setSession(token, user) {
  localStorage.setItem(KEYS.TOKEN, token);
  localStorage.setItem(KEYS.USER, JSON.stringify(user));
  localStorage.setItem(KEYS.EXPIRY, String(Date.now() + SESSION_CONFIG.DURATION_MS));
}

export function getToken() {
  return localStorage.getItem(KEYS.TOKEN);
}

export function getUser() {
  const raw = localStorage.getItem(KEYS.USER);
  return raw ? JSON.parse(raw) : null;
}

export function isSessionExpired() {
  const expiry = localStorage.getItem(KEYS.EXPIRY);
  if (!expiry) return true;
  return Date.now() > parseInt(expiry, 10);
}

export function isAuthenticated() {
  return !!getToken() && !!getUser() && !isSessionExpired();
}

export function clearSession() {
  localStorage.removeItem(KEYS.TOKEN);
  localStorage.removeItem(KEYS.USER);
  localStorage.removeItem(KEYS.EXPIRY);
}

export function getRemainingSessionTime() {
  const expiry = localStorage.getItem(KEYS.EXPIRY);
  if (!expiry) return 0;
  return Math.max(0, parseInt(expiry, 10) - Date.now());
}

export function getFormattedRemainingTime() {
  const ms = getRemainingSessionTime();
  if (ms <= 0) return 'संपले';
  const h = Math.floor(ms / 3600000);
  const m = Math.floor((ms % 3600000) / 60000);
  return h > 0 ? `${h}ता ${m}मि` : `${m}मि`;
}

export function hasRole(role) {
  const user = getUser();
  return !!user && user.role === role;
}

export function isAdmin() {
  return hasRole('Admin');
}

/**
 * Redirects to login if not authenticated. Call at the top of every
 * protected page. Returns true if the guard passed (page can continue).
 */
export function requireAuth() {
  if (!isAuthenticated()) {
    clearSession();
    window.location.href = 'index.html';
    return false;
  }
  return true;
}

/**
 * Additionally requires a specific role (used by dashboard.html for Admin).
 * Non-admin roles are redirected to the visit form rather than shown an error.
 */
export function requireRole(role) {
  if (!requireAuth()) return false;
  if (!hasRole(role)) {
    window.location.href = 'visit.html';
    return false;
  }
  return true;
}

export function logoutAndRedirect() {
  clearSession();
  window.location.href = 'index.html';
}
