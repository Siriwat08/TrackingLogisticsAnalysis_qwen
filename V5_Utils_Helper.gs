/**
 * V5_Utils_Helper.gs
 * Version: 5.0
 * Description: Core utility functions used across the entire system.
 *              Handles UUID generation, text normalization, distance calculation, and data cleaning.
 */

// --- 1. UUID Generation ---
/**
 * Generates a standard UUID v4 string.
 * @returns {string} Unique ID
 */
function V5_GenerateUUID() {
  return Utilities.getUuid();
}

// --- 2. Text Normalization ---
/**
 * Cleans and normalizes text for comparison.
 * - Trims whitespace
 * - Converts to lowercase
 * - Removes special characters (keeping Thai/Eng letters and numbers)
 * - Collapses multiple spaces into one
 * @param {string} text - Raw text input
 * @returns {string} Normalized text
 */
function V5_NormalizeText(text) {
  if (!text || text === "") return "";
  
  let t = text.toString().trim();
  
  // Convert to lowercase (works for English, Thai remains same case-wise but we clean noise)
  t = t.toLowerCase();
  
  // Remove specific noise words common in Thai logistics (Optional customization)
  const noiseWords = ["บจก.", "บริษัท", "จำกัด", "ห้างหุ้นส่วน", "ฮูท", "ร้าน", "เลขที่", "หมู่", "ม.", "ซอย", "ซ.", "ถ.", "ถนน", "ชั้น", "floor"];
  // Note: We might not want to remove these entirely for display, but for matching logic, sometimes helpful. 
  // For V5, we will keep them simple: just remove special chars.
  
  // Remove special characters except Thai/Eng letters, numbers, and spaces
  // This regex keeps: a-z, 0-9, ก-ฮ, space
  t = t.replace(/[^a-z0-9\u0E00-\u0E7F\s]/g, " ");
  
  // Collapse multiple spaces
  t = t.replace(/\s+/g, " ").trim();
  
  return t;
}

// --- 3. Distance Calculation (Haversine Formula) ---
/**
 * Calculates distance between two lat/lng points in meters.
 * @param {number} lat1
 * @param {number} lon1
 * @param {number} lat2
 * @param {number} lon2
 * @returns {number} Distance in meters
 */
function V5_CalculateDistanceMeters(lat1, lon1, lat2, lon2) {
  if (lat1 === "" || lon1 === "" || lat2 === "" || lon2 === "") return 999999;
  
  const R = 6371e3; // Earth radius in metres
  const φ1 = lat1 * Math.PI / 180;
  const φ2 = lat2 * Math.PI / 180;
  const Δφ = (lat2 - lat1) * Math.PI / 180;
  const Δλ = (lon2 - lon1) * Math.PI / 180;

  const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
            Math.cos(φ1) * Math.cos(φ2) *
            Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
            
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));

  return R * c;
}

// --- 4. LatLong Key Generation ---
/**
 * Creates a unique key from Latitude and Longitude (rounded to 6 decimal places).
 * Useful for quick duplicate checking of locations.
 * Format: "13.756300,100.501800"
 * @param {number} lat
 * @param {number} lng
 * @returns {string} LatLong Key
 */
function V5_CreateLatLongKey(lat, lng) {
  if (lat === "" || lng === "" || lat === null || lng === null) return "";
  const lats = parseFloat(lat);
  const lngs = parseFloat(lng);
  if (isNaN(lats) || isNaN(lngs)) return "";
  
  return lats.toFixed(6) + "," + lngs.toFixed(6);
}

// --- 5. Safe Data Parsing ---
/**
 * Safely parses a string "lat,lng" into an object {lat, lng}.
 * @param {string} str - e.g., "13.75,100.50"
 * @returns {object|null} {lat: number, lng: number} or null if invalid
 */
function V5_ParseLatLongString(str) {
  if (!str || typeof str !== 'string') return null;
  const parts = str.split(',');
  if (parts.length < 2) return null;
  
  const lat = parseFloat(parts[0].trim());
  const lng = parseFloat(parts[1].trim());
  
  if (isNaN(lat) || isNaN(lng)) return null;
  
  return { lat, lng };
}

// --- 6. Date Formatting ---
/**
 * Returns current timestamp in ISO format or Spreadsheet friendly format.
 * @returns {string} Date string
 */
function V5_GetTimestamp() {
  return new Date();
}
