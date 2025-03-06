function onOpen(e) {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('Fiscal Eye');

    if (e && e.authMode === ScriptApp.AuthMode.LIMITED) {
        menu.addItem('Grant File Access', 'requestFullAuthorization');
        ui.alert("Fiscal Eye requires additional authorization to work with this file. Please select 'Grant File Access' and authorize to continue.");
    } else if (e && e.authMode === ScriptApp.AuthMode.FULL) {
        addMenuItems(menu);
    } else {
        addMenuItems(menu);
    }

    menu.addToUi();
}

function addMenuItems(menu) {
    menu.addItem('Analyze Sheet', 'analyzeSheet')
        .addItem('Open Chat', 'showChatSidebar')
        .addItem('QuickBooks Integration', 'showQuickBooksConfig')
        .addItem('Generate Report', 'showReportDialog')
        .addItem('Set Gemini Model', 'showGeminiModelDialog');
}

function requestFullAuthorization() {
    // This function is intentionally left empty. It's a placeholder for triggering authorization if needed.
    // The onOpen function itself handles the authorization request and UI display.
}

function onInstall(e) {
    onOpen(e);
}

function onHomepage(e) {
    return HtmlService.createHtmlOutputFromFile('UI_Main').setTitle('Fiscal Eye');
}

function analyzeSheet() {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const anomalies = detectAnomalies(sheet); // Call AnomalyDetection.gs function
        highlightAnomalies(sheet, anomalies);
        createErrorReportSheet(anomalies);
        SpreadsheetApp.getUi().alert(`${anomalies.length} anomalies found and highlighted.  Error report sheet created.`);
    } catch (error) {
        showError(error.message); // Use Utilities.gs function
    }
}

function showChatSidebar() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('UI_Main').setTitle('Fiscal Eye Chat');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function showQuickBooksConfig() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('UI_QuickBooksConfig').setWidth(500).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'QuickBooks Integration');
}

function showReportDialog() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('UI_ReportMenu').setWidth(500).setHeight(600);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generate Report');
}

function showGeminiModelDialog() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('UI_GeminiModelSelection').setWidth(300).setHeight(200);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Set Gemini Model');
}

async function handleUserQuery(userQuery) {
    if (!userQuery) {
        return "Please enter a query.";
    }

    try {
        const sheetData = getSheetData();
        return await generateResponse(userQuery, sheetData); // Call Gemini.gs function
    } catch (error) {
        logError(error.message); // Use Utilities.gs function
        return "An error occurred: " + error.message;
    }
}

async function generateReportWithOptions(options) {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        let errorSheet = spreadsheet.getSheetByName('Error Report');
        let anomalies = [];

        if (!errorSheet) {
            logMessage("'Error Report' sheet not found. Processing entire spreadsheet."); // Use Utilities.gs function
            anomalies = await detectAnomalies(spreadsheet.getActiveSheet()); // Call AnomalyDetection.gs function to analyze current sheet
            errorSheet = spreadsheet.getSheetByName('Error Report'); // Re-fetch error sheet after analysis
        }

        if (errorSheet) {
            anomalies = getAnomaliesFromErrorSheet(errorSheet);
        } else {
            return 'Failed to create or find the Error Report sheet.';
        }

        if (options.includeAIResults && options.includeDetailedAnalysis) { // Only call AI if detailed analysis is requested
            const sheetData = getSheetData();
            try {
                const aiAnalysisResult = await analyzeSheetWithGemini(sheetData); // Call Gemini.gs function
                if (aiAnalysisResult && aiAnalysisResult.anomalies) {
                    anomalies = anomalies.concat(aiAnalysisResult.anomalies);
                }
            } catch (aiError) {
                logError("AI Analysis for Report Failed: " + aiError.message); // Log AI error, but continue report generation
                return 'Report generated with a warning: AI analysis step failed. Check logs for details. Report will be generated without AI analysis.';
            }
        }

        const validAnomalies = anomalies.filter(a => !onlyHasNAOrZero(a));

        if (validAnomalies.length > 0) {
            const reportUrl = await createReport(validAnomalies, options); // Call ReportGeneration.gs function
            return `Report generated successfully.  Open it here: ${reportUrl}`;
        } else {
            return 'No anomalies found to generate a report.';
        }

    } catch (error) {
        logError("Error Generating Report with options:" + error); // Use Utilities.gs function
        return 'Failed to generate report: ' + error.message;
    }
}


function getAnomaliesFromErrorSheet(errorSheet) {
    const data = errorSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    return rows.map(r => headers.reduce((obj, h, idx) => { obj[h.toLowerCase()] = r[idx]; return obj; }, {}));
}

function onlyHasNAOrZero(anomaly) {
    return Object.values(anomaly).every(val => !val || val.toString().trim().toUpperCase() === 'N/A' || (typeof val === 'number' && val === 0));
}

function setUserSelectedGeminiModel(modelName) {
    return setUserSelectedModel(modelName); // Call Gemini.gs function
}

/**
 * Main entry point for the application.
 * Handles user queries and routes them to the appropriate tools.
 * 
 * @param {string} userQuery The user's query
 * @returns {string} The response to the user
 */
function handleUserQuery(userQuery) {
  if (!userQuery) {
    return "Please enter a query.";
  }
  
  try {
    // Step 1: Call Gemini with tools to determine what function to call
    const toolUse = callGeminiWithTools(userQuery, WORKSPACE_TOOLS);
    console.log("Tool selected:", toolUse);
    
    // Step 2: Route to the appropriate function based on the function call
    if (toolUse.name === "analyzeTransactions") {
      return analyzeTransactions(toolUse.args.period, toolUse.args.type);
    } else if (toolUse.name === "generateReport") {
      return generateFinancialReport(toolUse.args.reportType, toolUse.args.includeAI);
    } else if (toolUse.name === "detectAnomalies") {
      return detectAndSummarizeAnomalies(toolUse.args.threshold);
    } else if (toolUse.name === "monthlyComparison") {
      return performMonthlyComparison();
    } else {
      // If no function matches, use general Gemini text response
      const sheetData = getSheetData();
      return generateResponse(userQuery, sheetData);
    }
  } catch (error) {
    logError(`Error handling user query: ${error}`);
    return `I encountered an error processing your request: ${error.message}. Please try again or rephrase your query.`;
  }
}

/**
 * Analyzes transactions for a given period and type.
 * 
 * @param {string} period The time period to analyze
 * @param {string} type The type of analysis
 * @returns {string} Analysis results
 */
function analyzeTransactions(period, type) {
  const sheetData = getSheetData();
  const prompt = `Analyze the financial transactions for ${period} focusing on ${type}. 
  Provide insights, trends, and recommendations based on the data.`;
  
  return generateResponse(prompt, sheetData);
}

/**
 * Generates a financial report of the specified type.
 * 
 * @param {string} reportType The type of report to generate
 * @param {boolean} includeAI Whether to include AI-powered insights
 * @returns {string} Confirmation message with report URL
 */
function generateFinancialReport(reportType, includeAI) {
  const options = {
    includeTitle: true,
    reportTitle: `${reportType} Financial Report`,
    includeIntroduction: true,
    includeSummary: true,
    includeChart: true,
    includeNumericAnalysis: true,
    includeCategoryBreakdown: true,
    includeDetailedAnalysis: true,
    includeAIResults: includeAI,
    locale: getDefaultLocale()
  };
  
  return generateReportWithOptions(options);
}

/**
 * Detects and summarizes anomalies in the transaction data.
 * 
 * @param {number} threshold The threshold for anomaly detection
 * @returns {string} Summary of detected anomalies
 */
function detectAndSummarizeAnomalies(threshold) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  try {
    const anomalies = detectAnomalies(sheet);
    if (anomalies.length === 0) {
      return "No anomalies detected in the current data.";
    }
    
    const prompt = `Summarize the following ${anomalies.length} anomalies detected in financial transactions:
    ${JSON.stringify(anomalies)}
    
    Provide a concise summary of the key issues, patterns, and recommended actions.`;
    
    return generateReportAnalysis(prompt);
  } catch (error) {
    logError(`Error detecting anomalies: ${error}`);
    return `Error detecting anomalies: ${error.message}`;
  }
}

/**
 * Performs a monthly comparison of financial data.
 * 
 * @returns {string} Monthly comparison analysis
 */
function performMonthlyComparison() {
  const sheetData = getSheetData();
  const prompt = `Compare the financial data month-by-month. 
  Include spending trends, major changes in categories, and insights on monthly financial health.
  Format the response with clear headings and bullet points for easy reading.`;
  
  return generateResponse(prompt, sheetData);
}

/**
 * Handles the report generation with the given options.
 * 
 * @param {object} options The report configuration options
 * @returns {string} Confirmation message with report URL
 */
function generateReportWithOptions(options) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const anomalies = detectAnomalies(sheet);
    
    if (anomalies.length === 0) {
      return "No anomalies or data issues to include in the report.";
    }
    
    const reportUrl = createReport(anomalies, options);
    return `Report generated successfully! You can view it at: ${reportUrl}`;
  } catch (error) {
    logError(`Error generating report: ${error}`);
    return `Error generating report: ${error.message}`;
  }
}