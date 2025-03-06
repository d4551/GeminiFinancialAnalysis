/**
 * Gets the active sheet's data.
 * @returns {Array<Array<any>>} The 2D array of sheet data.
 */
function getSheetData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    if (!sheet) {
        throw new Error('No active sheet found.');
    }
    return sheet.getDataRange().getValues();
}


/**
 * Logs an informational message.
 * @param {string} message
 */
function logMessage(message) {
    Logger.log(`[INFO] ${new Date().toISOString()} - ${message}`);
}

/**
 * Logs an error message and optionally shows it to the user.
 * @param {string} message
 */
function logError(message) {
    Logger.log(`[ERROR] ${new Date().toISOString()} - ${message}`);
}

/**
 * Show Error
 * @param {*} message
 */
function showError(message) {
    SpreadsheetApp.getUi().alert("An error occurred: " + message);
}

/**
 *
 * @returns Returns spreadsheet file
 */
function getOrCreateSpreadsheet() {
    try {
        let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        if (!spreadsheet) {
        const fileName = 'FiscalEye_Spreadsheet_' + new Date().getTime();
        spreadsheet = SpreadsheetApp.create(fileName);
        logMessage(`Created a new spreadsheet in Google Drive with name: ${fileName}`);
        } else {
        logMessage(`Using existing active spreadsheet with ID: ${spreadsheet.getId()}`);
        }
        return spreadsheet;
    } catch (error) {
        logError(`Error creating or retrieving spreadsheet: ${error.message}`);
        throw new Error('Failed to create or retrieve the spreadsheet.');
    }
}