/**
 * ============================================================================
 * CONFIG.JS
 * Central configuration for the Smart Services Tour Visit frontend.
 * ============================================================================
 */

export const API_CONFIG = {
  // ⚠️ REQUIRED: Replace with your deployed Apps Script Web App /exec URL
  // (Deploy > New deployment > Web app, in the backend project).
  BASE_URL: 'https://script.google.com/macros/s/REPLACE_WITH_YOUR_DEPLOYMENT_ID/exec',
  TIMEOUT_MS: 30000,
  RETRY_ATTEMPTS: 3,
  RETRY_DELAY_MS: 1000
};

export const BRAND = {
  NAME: 'Smart Services',
  TAGLINE: 'Leave it to us',
  COLOR_PRIMARY: '#1E315C', // Navy
  COLOR_SECONDARY: '#F4B400', // Yellow
  COLOR_BG: '#F8FAFC',
  COLOR_ACCENT: '#FFFFFF'
};

export const SESSION_CONFIG = {
  // Must match backend CONFIG.SESSION_TIMEOUT_SECONDS (Setup.gs) — Apps
  // Script's CacheService caps sessions at 6 hours server-side regardless
  // of what's set here, so this stays in sync rather than promising longer.
  DURATION_MS: 6 * 60 * 60 * 1000
};

export const UPLOAD_CONFIG = {
  MAX_BYTES: 10 * 1024 * 1024,
  ALLOWED_SIGNED_CHECKLIST_TYPES: ['application/pdf', 'image/jpeg', 'image/jpg', 'image/png']
};

// Free, no-API-key reverse geocoding for the "GPS Address" field. This is a
// best-effort convenience (spec says "GPS Address, if available") — OSM's
// Nominatim public instance rate-limits heavily-used deployments, so if
// this becomes unreliable at your traffic volume, swap in the Google Maps
// Geocoding API (requires a billed API key) by replacing reverseGeocode()
// in utils.js.
export const GEOCODING_CONFIG = {
  REVERSE_GEOCODE_URL: 'https://nominatim.openstreetmap.org/reverse'
};
