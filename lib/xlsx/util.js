/**
 * Utility functions for the SigmaLifting app
 */

/**
 * Generates a standardized filename for exports
 *
 * @param {string} type - The type of export ("process" or "program")
 * @param {string} extension - The file extension (e.g., "json", "xlsx")
 * @param {object} options - Optional parameters
 * @param {string} [options.name] - Optional name to use (process or program name) for detailed filenames
 * @param {boolean} [options.simpleFormat] - If true, use process_export_TIMESTAMP format
 * @returns {string} The generated filename
 */
export const generateExportFilename = (type, extension, options = {}) => {
  const { name = "", simpleFormat = false } = options;

  if (simpleFormat || (type === "process" && extension === "xlsx")) {
    // Simple format used by useViewProcesses hook for Excel exports
    const timestamp = new Date()
      .toISOString()
      .replace(/[-T:]/g, "")
      .substring(0, 14);
    return `process_export_${timestamp}.${extension}`;
  } else {
    // Standard format for other export types
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

    if (!name) {
      return `${type}_export_${timestamp}.${extension}`;
    }

    // Sanitize the name
    const safeName = name.replace(/[^a-z0-9]/gi, "_").toLowerCase();

    // Create filename based on type
    if (type === "process") {
      return `${safeName}_process_${timestamp}.${extension}`;
    } else if (type === "program") {
      return `${safeName}_${timestamp}.${extension}`;
    } else {
      // Default format if type is not recognized
      return `${safeName}_export_${timestamp}.${extension}`;
    }
  }
};

const INVALID_WORKSHEET_NAME_CHARACTERS = new Set([
  "*",
  "?",
  ":",
  "\\",
  "/",
  "[",
  "]",
]);

const sanitizeWorksheetName = (worksheetName) => {
  const safeName = String(worksheetName || "")
    .split("")
    .map((character) => {
      const isControlCharacter = character.charCodeAt(0) < 32;
      return INVALID_WORKSHEET_NAME_CHARACTERS.has(character) ||
        isControlCharacter
        ? "-"
        : character;
    })
    .join("")
    .replace(/\s+/g, " ")
    .trim();

  return safeName || "Sheet";
};

/**
 * Sanitizes and truncates an Excel worksheet name to fit Excel's constraints.
 *
 * @param {string} worksheetName - The desired worksheet name
 * @param {number} maxLength - Maximum allowed length (default: 31 for Excel. 30 for safety)
 * @returns {string} The safe worksheet name
 */
export const truncateWorksheetName = (worksheetName, maxLength = 30) => {
  const safeWorksheetName = sanitizeWorksheetName(worksheetName);

  if (safeWorksheetName.length <= maxLength) {
    return safeWorksheetName;
  }

  // For block names with format "Block X - Name", preserve the prefix
  const blockMatch = safeWorksheetName.match(/^(Block \d+ - )/);
  if (blockMatch) {
    const prefix = blockMatch[1];
    const remainingLength = maxLength - prefix.length;
    if (remainingLength > 0) {
      const blockName = safeWorksheetName.substring(prefix.length);
      return prefix + blockName.substring(0, remainingLength);
    }
  }

  // For other names, just truncate to maxLength
  return safeWorksheetName.substring(0, maxLength);
};

// Configuration constants for buffer operations
const BUFFER_CONFIG = {
  BASE64_CHUNK_SIZE: 1024 * 64, // 64KB chunks for base64 encoding
};

/**
 * Safely converts a buffer to base64 string without stack overflow.
 * Processes the buffer in chunks to avoid hitting string length limits.
 * @param {ArrayBuffer} buffer - The buffer to convert
 * @returns {string} Base64 encoded string
 */
export const bufferToBase64 = (buffer) => {
  const bytes = new Uint8Array(buffer);
  const chunks = [];

  // Process in chunks to avoid memory issues
  for (let i = 0; i < bytes.length; i += BUFFER_CONFIG.BASE64_CHUNK_SIZE) {
    const chunk = bytes.slice(
      i,
      Math.min(i + BUFFER_CONFIG.BASE64_CHUNK_SIZE, bytes.length)
    );

    // Convert chunk to binary string using apply for better performance
    // Note: fromCharCode.apply has a limit of ~65k arguments, our 64KB chunks are safe
    const binaryChunk = String.fromCharCode.apply(null, chunk);
    chunks.push(binaryChunk);
  }

  // Join all chunks and convert to base64
  return btoa(chunks.join(""));
};
