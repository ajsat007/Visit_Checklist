/**
 * ============================================================================
 * UTILS.JS
 * Shared helpers: HTML escaping, GPS capture, reverse geocoding, file
 * handling, formatting.
 * ============================================================================
 */
import { GEOCODING_CONFIG, UPLOAD_CONFIG } from './config.js';

/**
 * Escapes text before it's ever placed into innerHTML. Used everywhere
 * dynamic sheet-sourced text (questions, names, remarks, districts) is
 * rendered, so a stray "<" or "&" in someone's data can never be
 * interpreted as markup.
 */
export function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Wraps navigator.geolocation in a Promise with sane error messages.
 */
export function getCurrentPosition() {
  return new Promise((resolve, reject) => {
    if (!navigator.geolocation) {
      reject(new Error('या ब्राउझरमध्ये GPS उपलब्ध नाही'));
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (position) => resolve({
        latitude: position.coords.latitude,
        longitude: position.coords.longitude,
        accuracy: position.coords.accuracy
      }),
      (error) => {
        const messages = {
          1: 'GPS परवानगी नाकारली गेली',
          2: 'GPS स्थान उपलब्ध नाही',
          3: 'GPS स्थान घेण्यास वेळ लागला'
        };
        reject(new Error(messages[error.code] || 'GPS त्रुटी'));
      },
      { enableHighAccuracy: true, timeout: 15000, maximumAge: 60000 }
    );
  });
}

/**
 * Best-effort reverse geocoding via OSM Nominatim (no API key). Returns ''
 * (never throws) on any failure so a slow/unavailable geocoder never blocks
 * visit submission — GPS Address is explicitly "if available" in the spec.
 */
export async function reverseGeocode(latitude, longitude) {
  try {
    const url = `${GEOCODING_CONFIG.REVERSE_GEOCODE_URL}?format=json&lat=${latitude}&lon=${longitude}&zoom=16&addressdetails=0`;
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 6000);
    const response = await fetch(url, {
      signal: controller.signal,
      headers: { Accept: 'application/json' }
    });
    clearTimeout(timeoutId);
    if (!response.ok) return '';
    const data = await response.json();
    return data && data.display_name ? data.display_name : '';
  } catch (error) {
    console.warn('Reverse geocoding unavailable:', error.message);
    return '';
  }
}

export function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/**
 * Resizes + re-compresses an image before upload. Mobile camera photos are
 * routinely 3-8MB at full resolution — for a reference photo attached to a
 * checklist, that's pure wasted upload time and Drive storage. Resizes to
 * fit within maxDimension (long edge) and re-encodes as JPEG at the given
 * quality. Returns { base64, mimeType } ready for the upload API.
 * Falls back to the original file (still base64-encoded) if compression
 * fails for any reason — never blocks the actual upload over this.
 */
export async function compressImage(file, maxDimension = 1600, quality = 0.75) {
  try {
    const bitmap = await createImageBitmap(file);
    let { width, height } = bitmap;

    if (width > maxDimension || height > maxDimension) {
      const scale = maxDimension / Math.max(width, height);
      width = Math.round(width * scale);
      height = Math.round(height * scale);
    }

    const canvas = document.createElement('canvas');
    canvas.width = width;
    canvas.height = height;
    const ctx = canvas.getContext('2d');
    ctx.drawImage(bitmap, 0, 0, width, height);

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', quality));
    if (!blob) throw new Error('canvas.toBlob returned null');

    const base64 = await fileToBase64(blob);
    return { base64, mimeType: 'image/jpeg' };
  } catch (error) {
    console.warn('Image compression skipped, uploading original:', error.message);
    const base64 = await fileToBase64(file);
    return { base64, mimeType: file.type || 'image/jpeg' };
  }
}

/**
 * Validates a signed-checklist file against type/size before it's ever
 * base64-encoded and sent — mirrors the server-side check in
 * TourVisit_Backend.gs so the user gets instant feedback instead of a
 * round trip.
 */
export function validateSignedChecklistFile(file) {
  if (!UPLOAD_CONFIG.ALLOWED_SIGNED_CHECKLIST_TYPES.includes(file.type)) {
    return 'फक्त PDF, JPG किंवा PNG फाईल अनुमत आहे';
  }
  if (file.size > UPLOAD_CONFIG.MAX_BYTES) {
    return 'फाईल 10 MB पेक्षा मोठी आहे';
  }
  return null;
}

export function formatDateTime(date = new Date()) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

export function formatTime(date = new Date()) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

export function isOnline() {
  return navigator.onLine;
}
