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
 * Logs an error message to the Apps Script logger.
 * @param {string} message The error message.
 */
function logError(message) {
  console.error(`[ERROR] ${new Date().toISOString()} - ${message}`);
}

/**
 * Logs an informational message to the Apps Script logger.
 * @param {string} message The message to log.
 */
function logMessage(message) {
  console.log(`[INFO] ${new Date().toISOString()} - ${message}`);
}

/**
 * Shows an error message to the user.
 * @param {string} message The error message to display.
 */
function showError(message) {
  SpreadsheetApp.getUi().alert('Error', message, SpreadsheetApp.getUi().ButtonSet.OK);
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

/**
 * Returns the user-selected Gemini model or default if not set.
 * @returns {string} The model name
 */
function getSelectedGeminiModel() {
    const userModel = PropertiesService.getUserProperties().getProperty('USER_SELECTED_MODEL');
    if (userModel) {
        return userModel;
    }
    
    const scriptModel = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL');
    if (scriptModel) {
        return scriptModel;
    }
    
    return "gemini-1.5-pro-latest";
}

/**
 * Formats data for chart display.
 * @param {Array<Array<any>>} data The data to format
 * @returns {object} Formatted chart data
 */
function formatChartData(data) {
    if (!data || data.length < 2) {
        return null;
    }
    
    const headers = data[0];
    const values = data.slice(1);
    
    // Find date and numeric columns
    let dateColIndex = -1;
    const numericColIndexes = [];
    
    headers.forEach((header, index) => {
        if (/date/i.test(header)) {
            dateColIndex = index;
        } else if (values.some(row => typeof row[index] === 'number')) {
            numericColIndexes.push(index);
        }
    });
    
    // Format data for charts
    const chartData = {
        labels: [],
        datasets: numericColIndexes.map(idx => ({
            label: headers[idx],
            data: []
        }))
    };
    
    // If we have a date column, use it for labels
    if (dateColIndex >= 0) {
        values.forEach(row => {
            const dateLabel = row[dateColIndex] instanceof Date
                ? Utilities.formatDate(row[dateColIndex], Session.getScriptTimeZone(), 'yyyy-MM-dd')
                : row[dateColIndex].toString();
            
            chartData.labels.push(dateLabel);
            
            numericColIndexes.forEach((colIdx, datasetIdx) => {
                chartData.datasets[datasetIdx].data.push(row[colIdx]);
            });
        });
    } else {
        // Otherwise use row indices as labels
        values.forEach((row, rowIdx) => {
            chartData.labels.push(`Row ${rowIdx + 2}`);
            
            numericColIndexes.forEach((colIdx, datasetIdx) => {
                chartData.datasets[datasetIdx].data.push(row[colIdx]);
            });
        });
    }
    
    return chartData;
}

/**
 * Formats date values consistently across the application.
 * @param {Date} date The date to format
 * @param {string} locale The locale to use (defaults to user locale)
 * @returns {string} The formatted date string
 */
function formatDate(date, locale = null) {
    if (!(date instanceof Date)) {
        return String(date);
    }
    
    const userLocale = locale || getDefaultLocale();
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Formats currency values consistently across the application.
 * @param {number} amount The amount to format
 * @param {string} locale The locale to use (defaults to user locale)
 * @returns {string} The formatted currency string
 */
function formatCurrency(amount, locale = null) {
    if (typeof amount !== 'number') {
        return String(amount);
    }
    
    const userLocale = locale || getDefaultLocale();
    return amount.toLocaleString(userLocale, {
        style: 'currency',
        currency: 'USD', // You might want to make this configurable
        minimumFractionDigits: 2
    });
}

/**
 * Converts a value to a certain type based on expected format.
 * @param {any} value The value to convert
 * @param {string} type The target type ('date', 'number', 'string', etc.)
 * @returns {any} The converted value
 */
function convertValueToType(value, type) {
    if (value === null || value === undefined) {
        return null;
    }
    
    switch (type.toLowerCase()) {
        case 'date':
            if (value instanceof Date) return value;
            
            // Try to parse as date
            const dateVal = new Date(value);
            if (!isNaN(dateVal.getTime())) {
                return dateVal;
            }
            return null;
        
        case 'number':
            if (typeof value === 'number') return value;
            
            // Try to parse as number
            const num = parseFloat(value);
            return isNaN(num) ? null : num;
        
        case 'boolean':
            if (typeof value === 'boolean') return value;
            
            // Convert various string representations to boolean
            if (typeof value === 'string') {
                const lowerValue = value.toLowerCase().trim();
                if (['true', 'yes', '1'].includes(lowerValue)) return true;
                if (['false', 'no', '0'].includes(lowerValue)) return false;
            }
            return null;
        
        case 'string':
        default:
            return String(value);
    }
}

/**
 * Returns a URL with proper escaping for use in HTML.
 * @param {string} url The URL to sanitize
 * @returns {string} The sanitized URL
 */
function sanitizeUrl(url) {
    if (!url) return '';
    
    // Only allow http://, https:// and mailto: protocols
    if (!url.startsWith('http://') && !url.startsWith('https://') && !url.startsWith('mailto:')) {
        return '';
    }
    
    return url.replace(/[<>"]/g, '');
}

/**
 * Analyzes selected data in the current sheet.
 * Implementation of the missing function referenced in the menu.
 */
function analyzeSelectedData() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    ui.alert('Selection Required', 'Please select a range of data to analyze.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    ui.alert('Analyzing Selection', 'Analyzing the selected data...', ui.ButtonSet.OK);
    
    // Get the selected data
    const data = selection.getValues();
    
    // Perform basic analysis
    const summary = analyzeDataArray(data);
    
    // Display the results
    displaySelectionAnalysis(summary);
  } catch (error) {
    logError(`Error analyzing selected data: ${error.message}`);
    showError(`Failed to analyze selected data: ${error.message}`);
  }
}

/**
 * Analyzes a 2D array of data and provides a summary.
 * @param {Array<Array<any>>} data The data to analyze.
 * @returns {Object} A summary of the data.
 */
function analyzeDataArray(data) {
  try {
    // Count rows and columns
    const rowCount = data.length;
    const columnCount = data[0].length;
    
    // Count numeric cells and calculate basic statistics
    let numericCount = 0;
    let sum = 0;
    let min = Number.MAX_VALUE;
    let max = Number.MIN_VALUE;
    
    // Count empty cells
    let emptyCellCount = 0;
    
    // Analyze each cell
    for (let i = 0; i < rowCount; i++) {
      for (let j = 0; j < columnCount; j++) {
        const value = data[i][j];
        
        if (value === '' || value === null || value === undefined) {
          emptyCellCount++;
        } else if (typeof value === 'number' || !isNaN(Number(value))) {
          const numValue = typeof value === 'number' ? value : Number(value);
          numericCount++;
          sum += numValue;
          min = Math.min(min, numValue);
          max = Math.max(max, numValue);
        }
      }
    }
    
    // Calculate average if there are numeric cells
    const average = numericCount > 0 ? sum / numericCount : 0;
    
    return {
      rowCount,
      columnCount,
      cellCount: rowCount * columnCount,
      numericCount,
      emptyCellCount,
      sum,
      average,
      min: numericCount > 0 ? min : 0,
      max: numericCount > 0 ? max : 0
    };
  } catch (error) {
    logError(`Error in analyzeDataArray: ${error.message}`);
    return {
      error: error.message
    };
  }
}

/**
 * Displays the analysis results for selected data.
 * @param {Object} summary The data analysis summary.
 */
function displaySelectionAnalysis(summary) {
  if (summary.error) {
    showError(`Analysis error: ${summary.error}`);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  const message = `Analysis Results:
  
  • Rows: ${summary.rowCount}
  • Columns: ${summary.columnCount}
  • Total cells: ${summary.cellCount}
  • Numeric cells: ${summary.numericCount}
  • Empty cells: ${summary.emptyCellCount}
  
  Statistics (numeric values only):
  • Sum: ${summary.sum}
  • Average: ${summary.average.toFixed(2)}
  • Minimum: ${summary.min}
  • Maximum: ${summary.max}
  
  Would you like to generate a detailed report?`;
  
  const response = ui.alert('Selection Analysis', message, ui.ButtonSet.YES_NO);
  
  if (response === ui.Button.YES) {
    showReportDialog();
  }
}

/**
 * Shows the Gemini configuration dialog
 */
function showGeminiConfigDialog() {
  const content = `
    <div class="card">
      <h3>Gemini AI Settings</h3>
      
      <div class="form-group">
        <label for="apiKey">Gemini API Key:</label>
        <input type="text" id="apiKey" value="${getGeminiAPIKey()}">
      </div>
      
      <div class="form-group">
        <label>Response Settings:</label>
        <div class="form-group">
          <label>
            <input type="checkbox" id="enableAI" checked>
            Enable AI Analysis
          </label>
        </div>
      </div>
      
      <div class="form-group">
        <label>Model Configuration:</label>
        <button onclick="google.script.run.showGeminiModelDialog()">Configure Models</button>
      </div>
      
      <button onclick="saveGeminiConfig()">Save Settings</button>
      <div id="status"></div>
    </div>
    
    <script>
      function saveGeminiConfig() {
        const apiKey = document.getElementById('apiKey').value;
        const enableAI = document.getElementById('enableAI').checked;
        
        document.getElementById('status').innerHTML = 'Saving...';
        
        google.script.run
          .withSuccessHandler(message => {
            document.getElementById('status').className = 'success';
            document.getElementById('status').innerHTML = message;
          })
          .withFailureHandler(error => {
            document.getElementById('status').className = 'error';
            document.getElementById('status').innerHTML = 'Error: ' + error.message;
          })
          .saveGeminiConfig(apiKey, enableAI);
      }
    </script>
  `;
  
  const htmlOutput = createStandardDialog('Gemini AI Settings', content, 450, 350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gemini AI Settings');
}

/**
 * Saves Gemini configuration settings
 * @param {string} apiKey Gemini API key
 * @param {boolean} enableAI Whether to enable AI features
 * @returns {string} Status message
 */
function saveGeminiConfig(apiKey, enableAI) {
  try {
    PropertiesService.getScriptProperties().setProperties({
      'GEMINI_API_KEY': apiKey,
      'ENABLE_AI': enableAI.toString()
    });
    
    return 'Gemini AI settings saved successfully.';
  } catch (error) {
    logError(`Error saving Gemini config: ${error.message}`);
    throw new Error('Failed to save Gemini configuration.');
  }
}

/**
 * Calculates quartiles for a sorted array of numbers
 * @param {number[]} sortedArr The sorted array
 * @returns {Object} Object containing q1 and q3
 */
function calculateQuartiles(sortedArr) {
  const len = sortedArr.length;
  const q1Index = Math.floor(len / 4);
  const q3Index = Math.floor(3 * len / 4);
  return {
    q1: sortedArr[q1Index],
    q3: sortedArr[q3Index]
  };
}

/**
 * Appends content to a Google Doc report
 * @param {string} reportUrl URL of the report document
 * @param {string} content Content to append
 * @returns {Promise<void>}
 */
async function appendToReport(reportUrl, content) {
  try {
    // Extract the document ID from the URL
    const regex = /\/d\/([a-zA-Z0-9-_]+)/;
    const match = reportUrl.match(regex);
    
    if (!match || !match[1]) {
      throw new Error('Invalid document URL format');
    }
    
    const docId = match[1];
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    
    // Add a section header for the appended content
    body.appendParagraph('Additional Analysis')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      
    // Add the content
    body.appendParagraph(content);
    
    // Save the document
    doc.saveAndClose();
    
    return;
  } catch (error) {
    logError(`Error appending to report: ${error.message}`);
    throw new Error(`Failed to append to report: ${error.message}`);
  }
}

/**
 * Merges two arrays of anomalies, avoiding duplicates
 * @param {Array<Object>} anomalies1 First array of anomalies
 * @param {Array<Object>} anomalies2 Second array of anomalies
 * @returns {Array<Object>} Merged array
 */
function mergeAnomalies(anomalies1, anomalies2) {
  // Use a Map to identify duplicates by row number
  const merged = new Map();
  
  // Add all anomalies from the first array
  anomalies1.forEach(anomaly => {
    const key = anomaly.row || JSON.stringify(anomaly);
    merged.set(key, anomaly);
  });
  
  // Add or update with anomalies from the second array
  anomalies2.forEach(anomaly => {
    const key = anomaly.row || JSON.stringify(anomaly);
    
    if (merged.has(key)) {
      // Combine the errors arrays if they exist
      const existing = merged.get(key);
      if (Array.isArray(existing.errors) && Array.isArray(anomaly.errors)) {
        existing.errors = [...new Set([...existing.errors, ...anomaly.errors])];
      }
      
      // Keep the highest confidence score
      if (anomaly.confidence && (!existing.confidence || anomaly.confidence > existing.confidence)) {
        existing.confidence = anomaly.confidence;
      }
      
      merged.set(key, existing);
    } else {
      merged.set(key, anomaly);
    }
  });
  
  return Array.from(merged.values());
}

/**
 * Mocks the detection of anomalies until actual implementation is created
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to analyze
 * @param {Object} config Optional configuration
 * @returns {Array<Object>} Detected anomalies
 */
async function detectAnomalies(sheet, config = null) {
  try {
    // Sample implementation - this should be replaced with actual anomaly detection logic
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    // Simple anomaly detection for demonstration
    const anomalies = [];
    
    // Check for missing data in important fields
    rows.forEach((row, i) => {
      const rowNum = i + 2; // +2 because of 1-indexing and header row
      const anomaly = { row: rowNum };
      const errors = [];
      
      // Check for empty cells in the row
      headers.forEach((header, j) => {
        const value = row[j];
        anomaly[header.toLowerCase()] = value;
        
        // If cell should contain data but is empty
        if (value === "" && j < 3) {
          errors.push(`Missing ${header}`);
        }
      });
      
      // If numeric values are too large
      const amount = parseFloat(row[1]);
      if (!isNaN(amount) && amount > 10000) {
        errors.push("Unusually large amount");
        anomaly.confidence = 0.9;
      }
      
      // Add anomaly if errors were found
      if (errors.length > 0) {
        anomaly.errors = errors;
        if (!anomaly.confidence) {
          anomaly.confidence = 0.7;
        }
        anomalies.push(anomaly);
      }
    });
    
    return anomalies;
  } catch (error) {
    logError(`Error detecting anomalies: ${error.message}`);
    throw new Error(`Anomaly detection failed: ${error.message}`);
  }
}

/**
 * Highlights and annotates anomalies in the sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to annotate
 * @param {Array<Object>} anomalies Detected anomalies
 */
async function highlightAndAnnotateAnomalies(sheet, anomalies) {
  if (!sheet || !anomalies || anomalies.length === 0) return;
  
  try {
    // Create a consistent color scheme for different confidence levels
    const highConfidenceColor = "#f4cccc"; // Light red
    const mediumConfidenceColor = "#fce5cd"; // Light orange
    const lowConfidenceColor = "#fff2cc"; // Light yellow
    
    // Group anomalies by row for more efficient processing
    const anomaliesByRow = new Map();
    
    anomalies.forEach(anomaly => {
      if (!anomaly.row) return; // Skip anomalies without row information
      
      if (!anomaliesByRow.has(anomaly.row)) {
        anomaliesByRow.set(anomaly.row, []);
      }
      anomaliesByRow.get(anomaly.row).push(anomaly);
    });
    
    // Process each row with anomalies
    anomaliesByRow.forEach((rowAnomalies, rowNum) => {
      try {
        // Determine the highest confidence for this row's anomalies
        let maxConfidence = 0;
        let allErrors = [];
        let specificColumns = new Set();
        
        rowAnomalies.forEach(anomaly => {
          const confidence = anomaly.confidence || 0.5;
          maxConfidence = Math.max(maxConfidence, confidence);
          
          // Collect all error messages
          if (Array.isArray(anomaly.errors)) {
            allErrors = [...allErrors, ...anomaly.errors];
          }
          
          // Track specific columns mentioned in anomalies
          if (anomaly.column) {
            specificColumns.add(anomaly.column);
          }
        });
        
        // Choose color based on confidence level
        let backgroundColor;
        if (maxConfidence > 0.8) {
          backgroundColor = highConfidenceColor;
        } else if (maxConfidence > 0.5) {
          backgroundColor = mediumConfidenceColor;
        } else {
          backgroundColor = lowConfidenceColor;
        }
        
        // Apply highlighting to the row
        const rowRange = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn());
        
        // If specific columns were mentioned, only highlight those
        if (specificColumns.size > 0) {
          // Highlight specific cells
          specificColumns.forEach(colIdx => {
            const cellRange = sheet.getRange(rowNum, colIdx);
            cellRange.setBackground(backgroundColor);
            
            // Add a note with the error description if available
            const relatedAnomaly = rowAnomalies.find(a => a.column === colIdx);
            if (relatedAnomaly && Array.isArray(relatedAnomaly.errors)) {
              cellRange.setNote(relatedAnomaly.errors.join('\n'));
            } else {
              cellRange.setNote('Anomaly detected');
            }
          });
        } else {
          // Highlight entire row
          rowRange.setBackground(backgroundColor);
          
          // Add a note to the first cell with all error messages
          const firstCell = sheet.getRange(rowNum, 1);
          firstCell.setNote(allErrors.length > 0 
              ? allErrors.join('\n') 
              : 'Anomaly detected');
        }
        
        // Add a small visual indicator in the first cell
        const firstCell = sheet.getRange(rowNum, 1);
        const currentValue = firstCell.getValue();
        if (currentValue === "" || currentValue == null) {
          firstCell.setValue("⚠️"); // Add warning symbol if cell is empty
        }
      } catch (rowError) {
        logError(`Error highlighting row ${rowNum}: ${rowError.message}`);
      }
    });
    
    // Add a filter to the header row for easier analysis
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).createFilter();
    
    return true;
  } catch (error) {
    logError(`Error in highlightAndAnnotateAnomalies: ${error.message}`);
    return false;
  }
}

/**
 * Creates an error report sheet with detected anomalies
 * @param {Array<Object>} anomalies Detected anomalies
 */
function createErrorReportSheet(anomalies) {
  if (!anomalies || anomalies.length === 0) return false;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if the Error Report sheet already exists
    let errorSheet = ss.getSheetByName('Error Report');
    if (errorSheet) {
      // If it exists, clear it and reuse
      errorSheet.clear();
    } else {
      // Otherwise create a new sheet
      errorSheet = ss.insertSheet('Error Report');
    }
    
    // Set up headers
    const headers = [
      'Row', 'Column', 'Column Name', 'Value', 'Confidence', 
      'Errors', 'Mean', 'Std Dev', 'Z-Score', 'IQR', 'Duplicate Of'
    ];
    
    // Apply header formatting
    const headerRange = errorSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#eee');
    
    // Format the data for display
    const reportData = anomalies.map(anomaly => {
      return [
        anomaly.row || 'N/A', 
        anomaly.column || 'N/A', 
        anomaly.columnName || 'N/A',
        typeof anomaly.value !== 'undefined' ? String(anomaly.value) : 'N/A',
        anomaly.confidence ? `${(anomaly.confidence * 100).toFixed(0)}%` : 'N/A',
        Array.isArray(anomaly.errors) ? anomaly.errors.join(', ') : 'Unknown issue',
        typeof anomaly.mean !== 'undefined' ? anomaly.mean.toFixed(2) : 'N/A',
        typeof anomaly.stdDev !== 'undefined' ? anomaly.stdDev.toFixed(2) : 'N/A',
        typeof anomaly.zScore !== 'undefined' ? anomaly.zScore.toFixed(2) : 'N/A',
        typeof anomaly.iqr !== 'undefined' ? anomaly.iqr.toFixed(2) : 'N/A',
        anomaly.duplicateOf || 'N/A'
      ];
    });
    
    // Insert the data
    if (reportData.length > 0) {
      errorSheet.getRange(2, 1, reportData.length, headers.length)
        .setValues(reportData);
    }
    
    // Apply conditional formatting based on confidence
    const dataRange = errorSheet.getRange(2, 5, reportData.length, 1); // Confidence column
    const highConfidenceRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('8') // 80% or higher
        .setBackground('#f4cccc')
        .setRanges([dataRange])
        .build();
    
    const mediumConfidenceRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('5') // 50-79%
        .setBackground('#fce5cd')
        .setRanges([dataRange])
        .build();
    
    const lowConfidenceRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains('3') // Below 50%
        .setBackground('#fff2cc')
        .setRanges([dataRange])
        .build();
    
    const rules = errorSheet.getConditionalFormatRules();
    rules.push(highConfidenceRule, mediumConfidenceRule, lowConfidenceRule);
    errorSheet.setConditionalFormatRules(rules);
    
    // Auto-resize columns for better readability
    errorSheet.autoResizeColumns(1, headers.length);
    
    // Add filters to make it easier to sort and filter the anomalies
    errorSheet.getRange(1, 1, reportData.length + 1, headers.length).createFilter();
    
    // Add summary at the top
    const summarySheet = ss.getSheetByName('Summary') || ss.insertSheet('Summary');
    summarySheet.clear();
    
    // Calculate summary statistics
    const highConfidence = anomalies.filter(a => (a.confidence || 0) > 0.8).length;
    const mediumConfidence = anomalies.filter(a => {
      const conf = a.confidence || 0;
      return conf <= 0.8 && conf > 0.5;
    }).length;
    const lowConfidence = anomalies.filter(a => (a.confidence || 0) <= 0.5).length;
    
    // Insert summary data
    summarySheet.getRange('A1').setValue('Anomaly Detection Summary');
    summarySheet.getRange('A1:C1').merge().setFontWeight('bold').setBackground('#eee');
    
    const summaryData = [
      ['Total Anomalies', anomalies.length, ''],
      ['High Confidence Issues', highConfidence, 'Require immediate attention'],
      ['Medium Confidence Issues', mediumConfidence, 'Should be reviewed'],
      ['Low Confidence Issues', lowConfidence, 'May be false positives']
    ];
    
    summarySheet.getRange(2, 1, summaryData.length, 3).setValues(summaryData);
    summarySheet.autoResizeColumns(1, 3);
    
    // Return success
    return true;
  } catch (error) {
    logError(`Error creating error report sheet: ${error.message}`);
    return false;
  }
}

/**
 * Fetches data from QuickBooks using OAuth2
 * @param {string} companyId QuickBooks company ID
 * @param {string} query Query to execute
 * @returns {Promise<Array<Array<any>>>} The fetched data
 */
async function fetchQuickBooksData(companyId, query) {
  try {
    // Get QuickBooks configuration
    const clientId = getQuickbooksClientId();
    const clientSecret = getQuickbooksClientSecret();
    const environment = getQuickBooksEnvironment();
    
    if (!clientId || !clientSecret) {
      throw new Error('QuickBooks API credentials not configured');
    }
    
    // Check if we have a valid token
    const token = getQuickBooksToken();
    
    // Set API endpoints based on environment
    const apiBase = environment === 'PRODUCTION' 
      ? 'https://quickbooks.api.intuit.com'
      : 'https://sandbox-quickbooks.api.intuit.com';
    
    // Prepare the API request
    const endpoint = `${apiBase}/v3/company/${companyId}/query`;
    
    const headers = {
      'Authorization': `Bearer ${token}`,
      'Accept': 'application/json',
      'Content-Type': 'application/text'
    };
    
    const options = {
      'method': 'post',
      'headers': headers,
      'payload': query,
      'muteHttpExceptions': true
    };
    
    // Make the API request
    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    
    // Handle the response
    if (responseCode === 200) {
      const jsonResponse = JSON.parse(response.getContentText());
      
      // Process the response into a 2D array format
      if (jsonResponse.QueryResponse) {
        // Get the entity type (first property in QueryResponse)
        const entityType = Object.keys(jsonResponse.QueryResponse)[0];
        
        if (jsonResponse.QueryResponse[entityType] && 
            Array.isArray(jsonResponse.QueryResponse[entityType])) {
          
          const entities = jsonResponse.QueryResponse[entityType];
          
          if (entities.length === 0) {
            return [['No data found']];
          }
          
          // Extract all available field names from the first entity
          const firstEntity = entities[0];
          const fields = extractFieldsRecursive(firstEntity);
          
          // Create headers row
          const headers = fields.map(field => field.name);
          
          // Create data rows
          const rows = entities.map(entity => {
            return fields.map(field => {
              return getNestedValue(entity, field.path) || '';
            });
          });
          
          // Return headers + data rows
          return [headers, ...rows];
        }
      }
      
      return [['No data found or unrecognized response format']];
    } else if (responseCode === 401) {
      // Token expired, refresh and retry
      const newToken = refreshQuickBooksToken();
      if (newToken) {
        // Retry with the new token
        return fetchQuickBooksData(companyId, query);
      } else {
        throw new Error('Failed to refresh authentication token');
      }
    } else {
      const errorResponse = JSON.parse(response.getContentText());
      throw new Error(`QuickBooks API error: ${errorResponse.Fault?.Error?.[0]?.Message || 'Unknown error'}`);
    }
  } catch (error) {
    logError(`Error fetching QuickBooks data: ${error.message}`);
    
    // Return mock data for development/testing
    return [
      ['Date', 'Description', 'Amount', 'Category'],
      ['2023-01-15', 'Office Supplies', 125.99, 'Expenses'],
      ['2023-01-22', 'Client Payment', 1500.00, 'Revenue'],
      ['2023-02-05', 'Software Subscription', 49.99, 'Expenses'],
      ['2023-02-10', 'Consulting Services', 2500.00, 'Revenue'],
      ['2023-02-28', 'Office Rent', 1200.00, 'Expenses']
    ];
  }
}

/**
 * Helper function to extract fields recursively from an entity
 * @param {Object} entity The entity to extract fields from
 * @param {string} parentPath The parent path for nested objects
 * @param {Array<Object>} fields Accumulated fields
 * @returns {Array<Object>} Array of field objects with name and path
 */
function extractFieldsRecursive(entity, parentPath = '', fields = []) {
  for (const key in entity) {
    const path = parentPath ? `${parentPath}.${key}` : key;
    const value = entity[key];
    
    if (value !== null && typeof value === 'object' && !Array.isArray(value)) {
      // Recursive call for nested objects
      extractFieldsRecursive(value, path, fields);
    } else if (!Array.isArray(value)) {
      // Add only primitive fields, not arrays
      fields.push({ name: path, path: path });
    }
  }
  
  return fields;
}

/**
 * Helper function to get nested value from an object using path
 * @param {Object} obj The object to extract value from
 * @param {string} path The path to the value (e.g. 'user.name')
 * @returns {any} The extracted value
 */
function getNestedValue(obj, path) {
  return path.split('.').reduce((o, k) => (o || {})[k], obj);
}

/**
 * Gets QuickBooks access token
 * @returns {string} The access token
 */
function getQuickBooksToken() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const accessToken = scriptProperties.getProperty('QB_ACCESS_TOKEN');
  const tokenExpiry = scriptProperties.getProperty('QB_TOKEN_EXPIRY');
  
  // Check if token is still valid
  if (accessToken && tokenExpiry) {
    const expiryDate = new Date(parseInt(tokenExpiry));
    const now = new Date();
    
    if (expiryDate > now) {
      return accessToken;
    }
  }
  
  // Token expired or doesn't exist, refresh it
  return refreshQuickBooksToken();
}

/**
 * Refreshes the QuickBooks access token
 * @returns {string} The new access token
 */
function refreshQuickBooksToken() {
  try {
    const clientId = getQuickbooksClientId();
    const clientSecret = getQuickbooksClientSecret();
    const refreshToken = PropertiesService.getScriptProperties().getProperty('QB_REFRESH_TOKEN');
    
    if (!clientId || !clientSecret || !refreshToken) {
      throw new Error('Missing QuickBooks credentials or refresh token');
    }
    
    // OAuth2 token endpoint
    const tokenUrl = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer';
    
    // Create authorization header
    const authHeader = 'Basic ' + Utilities.base64Encode(`${clientId}:${clientSecret}`);
    
    // Set up request parameters
    const payload = {
      'grant_type': 'refresh_token',
      'refresh_token': refreshToken
    };
    
    const options = {
      'method': 'post',
      'headers': {
        'Authorization': authHeader,
        'Content-Type': 'application/x-www-form-urlencoded',
        'Accept': 'application/json'
      },
      'payload': payload,
      'muteHttpExceptions': true
    };
    
    // Make the request
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const jsonResponse = JSON.parse(response.getContentText());
      
      // Save the new tokens
      const scriptProperties = PropertiesService.getScriptProperties();
      scriptProperties.setProperty('QB_ACCESS_TOKEN', jsonResponse.access_token);
      scriptProperties.setProperty('QB_REFRESH_TOKEN', jsonResponse.refresh_token);
      
      // Calculate expiry time
      const expiryTime = new Date().getTime() + (jsonResponse.expires_in * 1000);
      scriptProperties.setProperty('QB_TOKEN_EXPIRY', expiryTime.toString());
      
      return jsonResponse.access_token;
    } else {
      throw new Error(`Failed to refresh token. Status: ${responseCode}`);
    }
  } catch (error) {
    logError(`Error refreshing QuickBooks token: ${error.message}`);
    return null;
  }
}

/**
 * Inserts QuickBooks data into a sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to insert data into
 * @param {Array<Array<any>>} data The data to insert
 * @returns {boolean} Whether the operation was successful
 */
function insertQuickBooksData(sheet, data) {
  if (!sheet || !data || data.length === 0) {
    return false;
  }

  try {
    // Clear existing content
    sheet.clear();
    
    // Insert all data at once for better performance
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Format headers
    const headers = sheet.getRange(1, 1, 1, data[0].length);
    headers.setFontWeight('bold');
    headers.setBackground('#f3f3f3');
    
    // Format date columns
    const headerRow = data[0];
    headerRow.forEach((header, index) => {
      // If it looks like a date column
      if (header.toString().toLowerCase().includes('date')) {
        const columnLetter = columnToLetter(index + 1);
        const columnRange = sheet.getRange(`${columnLetter}2:${columnLetter}${data.length}`);
        columnRange.setNumberFormat('yyyy-mm-dd');
      }
      
      // If it looks like a monetary column
      if (header.toString().toLowerCase().includes('amount') || 
          header.toString().toLowerCase().includes('price') || 
          header.toString().toLowerCase().includes('cost') ||
          header.toString().toLowerCase().includes('total')) {
        const columnLetter = columnToLetter(index + 1);
        const columnRange = sheet.getRange(`${columnLetter}2:${columnLetter}${data.length}`);
        columnRange.setNumberFormat('$#,##0.00');
      }
    });
    
    // Auto-resize columns for better readability
    sheet.autoResizeColumns(1, data[0].length);
    
    // Add filtering capability
    sheet.getRange(1, 1, data.length, data[0].length).createFilter();
    
    // Create data validation for categorical columns if appropriate
    identifyAndCreateDataValidation(sheet, data);
    
    return true;
  } catch (error) {
    logError(`Error inserting QuickBooks data: ${error.message}`);
    return false;
  }
}

/**
 * Converts column number to letter
 * @param {number} column The column number (1-based)
 * @returns {string} The column letter(s)
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Creates data validation for categorical columns
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet
 * @param {Array<Array<any>>} data The data
 */
function identifyAndCreateDataValidation(sheet, data) {
  if (data.length < 3) return; // Need at least header + 2 rows
  
  const headers = data[0];
  const dataRows = data.slice(1);
  
  headers.forEach((header, colIndex) => {
    // Skip if header is empty
    if (!header) return;
    
    // Check if the column might contain categories
    const columnValues = dataRows.map(row => row[colIndex]);
    const uniqueValues = [...new Set(columnValues.filter(val => val !== null && val !== ''))];
    
    // If number of unique values is reasonable for a category and less than 50% of rows
    if (uniqueValues.length > 1 && uniqueValues.length <= 20 && uniqueValues.length <= dataRows.length * 0.5) {
      // This looks like a categorical column
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(uniqueValues, true)
        .setAllowInvalid(true)
        .build();
      
      const columnLetter = columnToLetter(colIndex + 1);
      sheet.getRange(`${columnLetter}2:${columnLetter}${data.length}`).setDataValidation(rule);
    }
  });
}

/**
 * Saves the user's Gemini model selections
 * @param {string} textModel The selected text model name
 * @param {string} visionModel The selected vision model name
 * @returns {boolean} Whether the save was successful
 */
function saveGeminiModelSelections(textModel, visionModel) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperties({
      'USER_SELECTED_TEXT_MODEL': textModel,
      'USER_SELECTED_VISION_MODEL': visionModel
    });
    
    logMessage(`Saved model selections - Text: ${textModel}, Vision: ${visionModel}`);
    return true;
  } catch (error) {
    logError(`Error saving model selections: ${error.message}`);
    throw new Error('Failed to save model selections');
  }
}

/**
 * Gets the user-selected text model or default if not set
 * @returns {string} The text model name
 */
function getUserSelectedTextModel() {
  const userModel = PropertiesService.getUserProperties().getProperty('USER_SELECTED_TEXT_MODEL');
  if (userModel) {
    return userModel;
  }
  
  const scriptModel = PropertiesService.getScriptProperties().getProperty('GEMINI_TEXT_MODEL');
  if (scriptModel) {
    return scriptModel;
  }
  
  return "gemini-1.5-pro-latest";
}

/**
 * Gets the user-selected vision model or default if not set
 * @returns {string} The vision model name
 */
function getUserSelectedVisionModel() {
  const userModel = PropertiesService.getUserProperties().getProperty('USER_SELECTED_VISION_MODEL');
  if (userModel) {
    return userModel;
  }
  
  const scriptModel = PropertiesService.getScriptProperties().getProperty('GEMINI_VISION_MODEL');
  if (scriptModel) {
    return scriptModel;
  }
  
  return "gemini-1.5-pro-vision-latest";
}

/**
 * Gets cached Gemini models from the user cache
 * @returns {Array<Object>} Array of model objects or empty array if none cached
 */
function getCachedGeminiModels() {
  try {
    const cache = CacheService.getUserCache();
    const cachedModels = cache.get('GEMINI_MODELS');
    
    if (cachedModels) {
      return JSON.parse(cachedModels);
    }
    return [];
  } catch (error) {
    logError(`Error retrieving cached models: ${error.message}`);
    return [];
  }
}

/**
 * Gets user email for report author fields
 * @returns {string} User's email address or empty string
 */
function getUserEmail() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (error) {
    logError(`Error getting user email: ${error.message}`);
    return '';
  }
}

/**
 * Creates a rich tooltip with helpful information
 * @param {string} text The tooltip text content
 * @param {string} title Optional tooltip title
 * @returns {string} HTML for the tooltip
 */
function createTooltip(text, title = null) {
  const titleHtml = title ? `<strong>${sanitizeHtml(title)}</strong><br>` : '';
  return `
    <span class="tooltip">?
      <span class="tooltiptext">${titleHtml}${sanitizeHtml(text)}</span>
    </span>
  `;
}

/**
 * Formats an anomaly confidence score as a percentage with color coding
 * @param {number} confidence The confidence score (0-1)
 * @returns {string} HTML-formatted confidence display
 */
function formatConfidence(confidence) {
  if (typeof confidence !== 'number' || isNaN(confidence)) {
    return '<span class="confidence low">Unknown</span>';
  }
  
  const percent = Math.round(confidence * 100);
  let levelClass = 'low';
  
  if (percent > 80) {
    levelClass = 'high';
  } else if (percent > 50) {
    levelClass = 'medium';
  }
  
  return `<span class="confidence ${levelClass}">${percent}%</span>`;
}

/**
 * Creates standardized section headers for reports
 * @param {string} title Section title
 * @param {number} level Header level (1-6)
 * @returns {string} Formatted HTML header
 */
function createReportSectionHeader(title, level = 2) {
  if (level < 1) level = 1;
  if (level > 6) level = 6;
  
  return `<h${level} class="report-section-header">${sanitizeHtml(title)}</h${level}>`;
}

/**
 * Gets file details from a URL
 * @param {string} fileUrl Google Drive file URL
 * @returns {Object} File details including name, ID, type, and owner
 */
function getFileDetailsFromUrl(fileUrl) {
  try {
    // Extract the file ID from the URL
    const regex = /\/d\/([a-zA-Z0-9-_]+)/;
    const match = fileUrl.match(regex);
    
    if (!match || !match[1]) {
      throw new Error('Invalid file URL format');
    }
    
    const fileId = match[1];
    const file = DriveApp.getFileById(fileId);
    
    return {
      id: fileId,
      name: file.getName(),
      type: file.getMimeType(),
      owner: file.getOwner().getEmail(),
      url: fileUrl,
      lastUpdated: file.getLastUpdated()
    };
  } catch (error) {
    logError(`Error getting file details: ${error.message}`);
    return {
      id: null,
      name: 'Unknown File',
      type: 'Unknown',
      url: fileUrl
    };
  }
}

/**
 * Validates an email address format
 * @param {string} email Email address to validate
 * @returns {boolean} Whether the email is valid
 */
function isValidEmail(email) {
  if (!email) return false;
  
  // Basic email validation regex
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Calculates summary statistics for financial data
 * @param {Array<Array<any>>} data The data to analyze
 * @param {number} valueColumnIndex The index of the column with numeric values
 * @returns {Object} Summary statistics
 */
function calculateFinancialStats(data, valueColumnIndex) {
  if (!data || data.length <= 1) {
    return {
      count: 0,
      sum: 0,
      mean: 0,
      median: 0,
      min: 0,
      max: 0,
      stdDev: 0
    };
  }
  
  // Extract values from the specified column, skipping header row
  const values = data.slice(1)
    .map(row => row[valueColumnIndex])
    .filter(val => typeof val === 'number' && !isNaN(val));
  
  if (values.length === 0) {
    return {
      count: 0,
      sum: 0,
      mean: 0,
      median: 0,
      min: 0,
      max: 0,
      stdDev: 0
    };
  }
  
  // Calculate statistics
  const count = values.length;
  const sum = values.reduce((acc, val) => acc + val, 0);
  const mean = sum / count;
  
  // Sort values for median and min/max
  const sortedValues = [...values].sort((a, b) => a - b);
  const min = sortedValues[0];
  const max = sortedValues[sortedValues.length - 1];
  
  // Calculate median
  const midPoint = Math.floor(sortedValues.length / 2);
  const median = sortedValues.length % 2 === 0
    ? (sortedValues[midPoint - 1] + sortedValues[midPoint]) / 2
    : sortedValues[midPoint];
  
  // Calculate standard deviation
  const variance = values.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / count;
  const stdDev = Math.sqrt(variance);
  
  return {
    count,
    sum,
    mean,
    median,
    min,
    max,
    stdDev,
    range: max - min
  };
}

/**
 * Identify potential outliers in financial data using Z-scores
 * @param {Array<number>} values Array of numeric values
 * @param {number} threshold Z-score threshold (default: 3)
 * @returns {Array<{value: number, index: number, zScore: number}>} Identified outliers
 */
function identifyOutliers(values, threshold = 3) {
  if (!values || values.length <= 2) {
    return [];
  }
  
  // Calculate mean
  const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
  
  // Calculate standard deviation
  const variance = values.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / values.length;
  const stdDev = Math.sqrt(variance);
  
  // If standard deviation is 0 or very small, no outliers can be determined
  if (stdDev < 0.0001) {
    return [];
  }
  
  // Identify outliers using Z-scores
  const outliers = [];
  values.forEach((value, index) => {
    const zScore = Math.abs((value - mean) / stdDev);
    if (zScore > threshold) {
      outliers.push({
        value,
        index,
        zScore
      });
    }
  });
  
  return outliers;
}

/**
 * Groups financial data by a category column
 * @param {Array<Array<any>>} data The data to group
 * @param {number} categoryColIndex The index of the category column
 * @param {number} valueColIndex The index of the value column
 * @returns {Object} Grouped data with calculated totals
 */
function groupFinancialData(data, categoryColIndex, valueColIndex) {
  if (!data || data.length <= 1) {
    return {};
  }
  
  // Skip header row
  const rows = data.slice(1);
  const groups = {};
  let total = 0;
  
  // Group values by category
  rows.forEach(row => {
    const category = row[categoryColIndex]?.toString() || 'Unknown';
    const value = typeof row[valueColIndex] === 'number' ? row[valueColIndex] : 0;
    total += value;
    
    if (!groups[category]) {
      groups[category] = {
        values: [],
        total: 0,
        count: 0
      };
    }
    
    groups[category].values.push(value);
    groups[category].total += value;
    groups[category].count++;
  });
  
  // Calculate additional statistics for each group
  Object.keys(groups).forEach(category => {
    const group = groups[category];
    
    // Calculate mean for the group
    group.mean = group.total / group.count;
    
    // Calculate percentage of total
    group.percentage = total !== 0 ? (group.total / total * 100) : 0;
  });
  
  return {
    groups,
    total,
    categoryCount: Object.keys(groups).length
  };
}

/**
 * Gets the user-selected text model or default if not set
 * @returns {string} The text model name
 */
function getUserSelectedTextModel() {
  const userModel = PropertiesService.getUserProperties().getProperty('USER_SELECTED_TEXT_MODEL');
  if (userModel) {
    return userModel;
  }
  
  const scriptModel = PropertiesService.getScriptProperties().getProperty('DEFAULT_TEXT_MODEL');
  if (scriptModel) {
    return scriptModel;
  }
  
  return "gemini-1.5-pro-latest";
}

/**
 * Gets the user-selected vision model or default if not set
 * @returns {string} The vision model name
 */
function getUserSelectedVisionModel() {
  const userModel = PropertiesService.getUserProperties().getProperty('USER_SELECTED_VISION_MODEL');
  if (userModel) {
    return userModel;
  }
  
  const scriptModel = PropertiesService.getScriptProperties().getProperty('DEFAULT_VISION_MODEL');
  if (scriptModel) {
    return scriptModel;
  }
  
  return "gemini-1.5-pro-vision-latest";
}

/**
 * Validates if a string is a properly formatted Google Document URL
 * @param {string} url The URL to validate
 * @returns {boolean} Whether the URL is valid
 */
function isValidDocUrl(url) {
  if (!url) return false;
  
  // Check for Google Docs URL format
  const regex = /https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9-_]+)\/edit/;
  return regex.test(url);
}