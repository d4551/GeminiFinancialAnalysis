/**
 * Main application entry point for Gemini Financial Analysis
 * Handles all core functionality and UI interactions
 */

const MENU_NAME = "Gemini Financial AI";
const APP_VERSION = "1.2.0"; // Incremented version due to improvements

// Tools configuration for Gemini AI function calling
const WORKSPACE_TOOLS = [
  {
    name: "analyzeTransactions",
    description: "Analyze financial transactions for a specific time period",
    parameters: {
      period: "string - The time period (monthly, quarterly, annually)",
      type: "string - The type of analysis (overview, trends, categories)"
    }
  },
  {
    name: "generateReport",
    description: "Generate a financial report",
    parameters: {
      reportType: "string - Type of report (standard, executive, monthly, anomaly, budget)",
      includeAI: "boolean - Whether to include AI-enhanced analysis"
    }
  },
  {
    name: "detectAnomalies",
    description: "Detect anomalies in transaction data",
    parameters: {
      threshold: "number - Threshold for anomaly detection (optional)"
    }
  },
  {
    name: "monthlyComparison",
    description: "Compare financial data month by month",
    parameters: {}
  },
  {
    name: "categoryAnalysis",
    description: "Analyze spending by categories",
    parameters: {
      topCategories: "number - Number of top categories to analyze (optional)",
      period: "string - Time period to analyze (optional)"
    }
  }
];

/**
 * Runs when the spreadsheet is opened
 */
function onOpen(e) {
    createMenu(e);
    
    // Load available Gemini models in background
    try {
        listGeminiModels();
    } catch (error) {
        logError("Failed to list Gemini models on startup: " + error.message);
    }
}

/**
 * Creates the application menu
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e Optional event object
 */
function createMenu(e) {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu(MENU_NAME);
    
    // Always add all menu items
    addMenuItems(menu);
    
    // Add authorization item if needed
    if (e && e.authMode === ScriptApp.AuthMode.LIMITED) {
        menu.addSeparator()
            .addItem('🔐 Grant Full Access', 'requestFullAuthorization');
            
        // Display notification about limited access
        ui.alert("Limited Access Mode", 
            "Some features require additional permissions. Click 'Grant Full Access' in the menu if you encounter any permission errors.", 
            ui.ButtonSet.OK);
    }

    menu.addToUi();
}

/**
 * Requests full authorization and rebuilds the menu with all options
 */
function requestFullAuthorization() {
    // Display a message to the user
    const ui = SpreadsheetApp.getUi();
    ui.alert('Authorization Required', 
        'Please click OK to authorize access to your sheet. The menu will be updated after authorization.',
        ui.ButtonSet.OK);
    
    // Force a refresh by calling a simple function that requires authorization
    try {
        // This call will force the authorization dialog if needed
        SpreadsheetApp.getActiveSpreadsheet().getName();
        
        // If we get here without error, rebuild the menu
        refreshMenu();
        
        ui.alert('Authorization Successful', 
            'Thank you for authorizing. You now have access to all features.',
            ui.ButtonSet.OK);
    } catch (error) {
        logError("Authorization error: " + error.message);
        ui.alert('Authorization Error', 
            'There was a problem with authorization. Please try again or reload the sheet.',
            ui.ButtonSet.OK);
    }
}

/**
 * Refreshes the menu with all available options
 */
function refreshMenu() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu(MENU_NAME);
    addMenuItems(menu);
    menu.addToUi();
}

/**
 * Adds menu items to the menu
 * @param {GoogleAppsScript.Base.Menu} menu The menu to add items to
 */
function addMenuItems(menu) {
    const ui = SpreadsheetApp.getUi();
    
    menu.addItem('Analyze Sheet', 'analyzeSheet')
        .addItem('Open Chat Assistant', 'showChatSidebar')
        .addSeparator()
        .addSubMenu(ui.createMenu('Reports')
            .addItem('Generate Standard Report', 'showReportDialog')
            .addItem('Generate Executive Summary', 'generateExecutiveSummary')
            .addItem('Email Report', 'showEmailReportDialog'))
        .addSeparator()
        .addSubMenu(ui.createMenu('Data Analysis')
            .addItem('Analyze Selected Data', 'analyzeSelectedData')
            .addItem('Monthly Comparison', 'runMonthlyComparison')
            .addItem('Transaction Pattern Analysis', 'showPatternAnalysisDialog'))
        .addSeparator()
        .addSubMenu(ui.createMenu('Integrations')
            .addItem('QuickBooks Integration', 'showQuickBooksConfig')
            .addItem('Import From QuickBooks', 'showQuickBooksImportDialog'))
        .addSeparator()
        .addSubMenu(ui.createMenu('Settings')
            .addItem('Configuration', 'showConfigDialog') 
            .addItem('Gemini AI Settings', 'showGeminiConfigDialog')
            .addItem('Set Gemini Models', 'showGeminiModelDialog'))
        .addSeparator()
        .addItem('About', 'showAboutDialog');
}

function requestFullAuthorization() {
    // This function is intentionally left empty. It's a placeholder for triggering authorization if needed.
    // The onOpen function itself handles the authorization request and UI display.
}

function onInstall(e) {
    onOpen(e);
}

function onHomepage(e) {
    return HtmlService.createHtmlOutputFromFile('UI_Main').setTitle('Gemini Financial AI');
}

/**
 * Runs anomaly detection on the current sheet
 */
async function analyzeSheet() {
    const ui = SpreadsheetApp.getUi();
    try {
        ui.alert('Starting anomaly detection...', 
            'This may take a moment, especially for large sheets.', 
            ui.ButtonSet.OK);
            
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const anomalies = await detectAnomalies(sheet);
        
        if (anomalies.length === 0) {
            ui.alert('Analysis Complete', 'No anomalies were detected in the data.', ui.ButtonSet.OK);
            return;
        }
        
        // Use the enhanced highlighting method that includes confidence scores
        await highlightAndAnnotateAnomalies(sheet, anomalies);
        
        // Create the enhanced error report sheet
        createErrorReportSheet(anomalies);
        
        // Group by confidence for better reporting
        const highConfidence = anomalies.filter(a => (a.confidence || 1.0) > 0.8).length;
        const mediumConfidence = anomalies.filter(a => {
            const conf = a.confidence || 1.0;
            return conf <= 0.8 && conf > 0.5;
        }).length;
        const lowConfidence = anomalies.filter(a => (a.confidence || 1.0) <= 0.5).length;
        
        const message = `Analysis complete. Found ${anomalies.length} potential issues:\n` +
            `• ${highConfidence} high confidence issues\n` +
            `• ${mediumConfidence} medium confidence issues\n` +
            `• ${lowConfidence} low confidence issues\n\n` +
            `Issues are highlighted in the sheet and detailed in the "Error Report" sheet.`;
            
        const response = ui.alert('Analysis Complete', message, ui.ButtonSet.OK_CANCEL);
        
        // If user clicks OK, offer to generate a report
        if (response === ui.Button.OK) {
            const generateReport = ui.alert('Generate Report', 
                'Would you like to generate a detailed report of these findings?', 
                ui.ButtonSet.YES_NO);
                
            if (generateReport === ui.Button.YES) {
                showReportDialog();
            }
        }
    } catch (error) {
        logError(`Error in analyzeSheet: ${error.message}`);
        showError("Error analyzing sheet: " + error.message);
    }
}

function showChatSidebar() {
    try {
        // Use the templated dialog with proper styling
        const htmlOutput = createTemplatedDialog(
            'UI_Main', 
            { version: APP_VERSION }, 
            { width: 400, title: 'Gemini Financial AI' }
        );
        
        SpreadsheetApp.getUi().showSidebar(htmlOutput);
    } catch (error) {
        logError(`Error showing chat sidebar: ${error.message}`);
        showError("Could not open chat assistant: " + error.message);
    }
}

/**
 * Shows the QuickBooks configuration dialog
 */
function showQuickBooksConfig() {
    try {
        const htmlOutput = createTemplatedDialog(
            'UI_QuickBooksConfig', 
            { 
                clientId: getQuickbooksClientId() || '',
                clientSecret: getQuickbooksClientSecret() || '',
                environment: getQuickBooksEnvironment() || 'SANDBOX'
            },
            { width: 500, height: 600, title: 'QuickBooks Integration' }
        );
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'QuickBooks Integration');
    } catch (error) {
        logError(`Error showing QuickBooks config: ${error.message}`);
        showError("Could not open QuickBooks configuration: " + error.message);
    }
}

/**
 * Shows the report generation dialog with enhanced options
 */
function showReportDialog() {
    try {
        const htmlOutput = createTemplatedDialog(
            'UI_ReportMenu',
            { 
                defaultLocale: getDefaultLocale(),
                defaultCurrency: getDefaultCurrency(), // Add this line to pass currency
                version: APP_VERSION
            },
            { width: 500, height: 650, title: 'Generate Financial Report' }
        );
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generate Financial Report');
    } catch (error) {
        logError(`Error showing report dialog: ${error.message}`);
        showError("Could not open report generator: " + error.message);
    }
}

/**
 * Shows dialog for selecting Gemini AI models with improved UI handling
 */
function showGeminiModelDialog() {
    try {
        const htmlOutput = createTemplatedDialog(
            'UI_GeminiModelSelection', 
            { currentVersion: APP_VERSION },
            { width: 500, height: 500, title: 'Select Gemini Models' }
        );
        
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Gemini Models');
    } catch (error) {
        logError(`Error showing Gemini model dialog: ${error.message}`);
        showError("Could not open model selector: " + error.message);
    }
}

/**
 * Shows the about dialog with app information
 */
function showAboutDialog() {
    const content = `
        <div style="text-align: center;">
            <div class="app-title" style="margin: 20px 0;">
                <h2>Gemini Financial AI</h2>
                <div style="color: #666; margin-bottom: 20px;">Version ${APP_VERSION}</div>
            </div>
            <p>An open source, AI-powered financial analysis tool integrated with Google Sheets</p>
            
            <div style="text-align: left; margin: 20px auto; max-width: 400px;">
                <h3>Key Features:</h3>
                <ul>
                    <li>AI-powered anomaly detection</li>
                    <li>Automated financial reporting</li>
                    <li>Interactive AI chat assistant</li>
                    <li>QuickBooks integration</li>
                    <li>Transaction pattern analysis</li>
                </ul>
            </div>
            
            <p>Powered by Google Gemini AI</p>
            <p><a href="https://github.com/d4551/GeminiFinancialAnalysis" target="_blank">View on GitHub</a></p>
        </div>
    `;
    
    const htmlOutput = createStandardDialog('About Gemini Financial AI', content, 400, 450);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About Gemini Financial AI');
}

/**
 * Handles user query and routes it to the appropriate function
 * @param {string} userQuery The user's query
 * @returns {string} The response
 */
async function handleUserQuery(userQuery) {
    if (!userQuery) {
        return "Please enter a query.";
    }
    
    logMessage(`Processing user query: "${userQuery}"`);
    
    try {
        // Check if AI is enabled in configuration
        const enableAI = PropertiesService.getScriptProperties().getProperty('ENABLE_AI');
        if (enableAI === 'false') {
            return "AI features are currently disabled. Please enable them in the Gemini AI Settings menu.";
        }
        
        // Check if the user is asking about categories specifically
        if (userQuery.toLowerCase().includes('categor') && 
            (userQuery.toLowerCase().includes('spending') || userQuery.toLowerCase().includes('expense'))) {
            return await analyzeCategoriesWithGemini();
        }
        
        // Step 1: Call Gemini with tools to determine what function to call
        const toolUse = callGeminiWithTools(userQuery, WORKSPACE_TOOLS);
        logMessage(`Tool selected for query "${userQuery}": ${toolUse.name}`);
        
        // Step 2: Route to the appropriate function based on the function call
        if (toolUse.name === "analyzeTransactions") {
            return await analyzeTransactions(
                toolUse.args.period || "monthly", 
                toolUse.args.type || "overview"
            );
        } else if (toolUse.name === "generateReport") {
            return await generateFinancialReport(
                toolUse.args.reportType || "standard", 
                toolUse.args.includeAI !== false
            );
        } else if (toolUse.name === "detectAnomalies") {
            return await detectAndSummarizeAnomalies(
                toolUse.args.threshold || 3
            );
        } else if (toolUse.name === "monthlyComparison") {
            return await performMonthlyComparison();
        } else if (toolUse.name === "categoryAnalysis") {
            return await analyzeCategories(
                toolUse.args.topCategories || 5,
                toolUse.args.period || "all"
            );
        } else {
            // If no function matches, use general Gemini text response
            const sheetData = getSheetData();
            return await generateResponse(userQuery, sheetData);
        }
    } catch (error) {
        logError(`Error handling user query: ${error.message}`);
        return `I encountered an error processing your request: ${error.message}. Please try again or rephrase your query.`;
    }
}

/**
 * Gets the current sheet data for analysis
 * @returns {Array<Array<any>>} The sheet data
 */
function getSheetData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    return sheet.getDataRange().getValues();
}

/**
 * Generates a financial report of the specified type - delegating to ReportGeneration.gs
 * @param {string} reportType The type of report to generate
 * @param {boolean} includeAI Whether to include AI-powered insights
 * @returns {Promise<string>} Confirmation message with report URL
 */
async function generateFinancialReport(reportType, includeAI) {
    return ReportGeneration.generateFinancialReport(reportType, includeAI);
}

/**
 * Generates an executive summary report - delegating to ReportGeneration.gs
 * @returns {Promise<string>} Confirmation message with report URL
 */
async function generateExecutiveSummary() {
    return ReportGeneration.generateExecutiveSummary();
}

/**
 * Shows dialog to send a report by email
 **/
function showEmailReportDialog() {
    try {
        const htmlOutput = createTemplatedDialog(
            'UI_EmailReportDialog',
            { userEmail: Session.getActiveUser().getEmail() || "" },
            { width: 400, height: 400, title: 'Email Financial Report' }
        );
        
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Email Financial Report');
    } catch (error) {
        logError(`Error showing email report dialog: ${error.message}`);
        showError("Could not open email dialog: " + error.message);
    }
}

/**
 * Generates and emails a financial report - delegating to ReportGeneration.gs
 * @param {string} recipient The recipient email
 * @param {string} reportType The type of report
 * @param {string} customMessage Optional custom message
 * @returns {string} Status message
 */
async function emailFinancialReport(recipient, reportType, customMessage) {
    try {
        // Delegate to specialized implementation in ReportGeneration.gs
        return await emailReport(recipient, reportType, customMessage);
    } catch (error) {
        logError(`Error in emailFinancialReport: ${error.message}`);
        return `Error: ${error.message}`;
    }
}

/**
 * Generates a report with the specified options
 * @param {Object} options Report options
 * @returns {Promise<string>} Status message with report URL
 */
async function generateReportWithOptions(options) {
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        let errorSheet = spreadsheet.getSheetByName('Error Report');
        let anomalies = [];

        // If no Error Report sheet exists, run anomaly detection
        if (!errorSheet) {
            logMessage("'Error Report' sheet not found. Processing entire spreadsheet.");
            anomalies = await detectAnomalies(spreadsheet.getActiveSheet());
            errorSheet = spreadsheet.getSheetByName('Error Report');
        }

        if (errorSheet) {
            anomalies = getAnomaliesFromErrorSheet(errorSheet);
        } else {
            return 'Failed to create or find the Error Report sheet.';
        }

        if (anomalies.length === 0) {
            // If no anomalies are found in the sheet, try to detect them directly
            anomalies = await detectAnomalies(spreadsheet.getActiveSheet());
            
            if (anomalies.length === 0) {
                return 'No anomalies found to generate a report.';
            }
        }

        // Enhance with AI analysis if requested
        if (options.includeAIResults && options.includeDetailedAnalysis) {
            try {
                const sheetData = getSheetData();
                const aiAnalysisResult = await analyzeSheetWithGemini(sheetData);
                
                if (aiAnalysisResult && aiAnalysisResult.anomalies) {
                    anomalies = mergeAnomalies(anomalies, aiAnalysisResult.anomalies);
                }
            } catch (aiError) {
                logError("AI Analysis for Report Failed: " + aiError.message);
                // Continue report generation without AI analysis
            }
        }

        // Filter out empty anomalies
        const validAnomalies = anomalies.filter(a => !onlyHasNAOrZero(a));

        if (validAnomalies.length > 0) {
            // Set default author if not provided
            if (!options.author) {
                const userEmail = Session.getActiveUser().getEmail();
                options.author = userEmail || "Financial Analysis System";
            }
            
            // Add department if not specified
            if (!options.department) {
                options.department = "Finance Department";
            }

            let reportUrl;
            if (options.reportType === 'executive' || options.includeExecutiveSummary) {
                // Use enhanced report for executive summaries
                reportUrl = await createEnhancedReport(validAnomalies, options);
            } else {
                reportUrl = await createReport(validAnomalies, options);
            }
            
            // Return a formatted response with HTML link instead of opening document directly
            return `Report generated successfully. <a href="${reportUrl}" target="_blank">Open Report</a>`;
        } else {
            return 'No valid anomalies found to generate a report.';
        }

    } catch (error) {
        logError("Error Generating Report with options:" + error);
        return 'Failed to generate report: ' + error.message;
    }
}

/**
 * Extract anomalies from the Error Report sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} errorSheet The error report sheet
 * @returns {Array<Object>} Array of anomaly objects
 */
function getAnomaliesFromErrorSheet(errorSheet) {
    const data = errorSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
        return []; // No anomalies, just headers
    }
    
    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    const rows = data.slice(1);
    
    return rows.map(r => {
        // Create basic anomaly object
        const anomaly = headers.reduce((obj, h, idx) => { 
            obj[h.toLowerCase()] = r[idx]; 
            return obj; 
        }, {});
        
        // Convert errors string to array if needed
        if (typeof anomaly.errors === 'string') {
            anomaly.errors = anomaly.errors.split(',').map(e => e.trim());
        }
        
        // Convert confidence string percentage to number if needed
        if (typeof anomaly.confidence === 'string' && anomaly.confidence.includes('%')) {
            anomaly.confidence = parseInt(anomaly.confidence) / 100;
        }
        
        return anomaly;
    });
}

/**
 * Checks if an anomaly only has N/A or zero values
 * @param {Object} anomaly The anomaly to check
 * @returns {boolean} True if only has N/A or zero values
 */
function onlyHasNAOrZero(anomaly) {
    return Object.values(anomaly).every(val => 
        !val || 
        val.toString().trim().toUpperCase() === 'N/A' || 
        (typeof val === 'number' && val === 0)
    );
}

/**
 * Analyzes transactions for a given period and type
 * @param {string} period The time period to analyze
 * @param {string} type The type of analysis
 * @returns {string} Analysis results
 */
async function analyzeTransactions(period, type) {
    const sheetData = getSheetData();
    
    const prompt = `Analyze the financial transactions for ${period} focusing on ${type}. 
    Provide insights, trends, and recommendations based on the data.
    
    Format your response with proper sections, bullet points, and highlight key insights.
    Include quantitative analysis where possible.`;
    
    return await generateResponse(prompt, sheetData);
}

/**
 * Generates a financial report of the specified type
 * @param {string} reportType The type of report to generate
 * @param {boolean} includeAI Whether to include AI-powered insights
 * @returns {string} Confirmation message with report URL
 */
async function generateFinancialReport(reportType, includeAI) {
    // Get template configuration if available
    const template = REPORT_CONFIG.defaultTemplates[reportType.toLowerCase()] || {};
    
    const options = {
        reportType: reportType,
        includeTitle: true,
        reportTitle: template.title || `${reportType.charAt(0).toUpperCase() + reportType.slice(1)} Financial Report`,
        includeIntroduction: true,
        includeSummary: template.sections?.includes('summary') ?? true,
        includeChart: true,
        includeNumericAnalysis: true,
        includeCategoryBreakdown: true,
        includeDetailedAnalysis: template.sections?.includes('recommendations') ?? true,
        includeAIResults: includeAI,
        includeExecutiveSummary: reportType.toLowerCase() === 'executive' || template.sections?.includes('highlights'),
        includeRecommendations: template.sections?.includes('recommendations') ?? includeAI,
        locale: getDefaultLocale(),
        currency: 'USD'
    };
    
    return await generateReportWithOptions(options);
}

/**
 * Generates an executive summary report
 * @returns {string} Confirmation message with report URL
 */
async function generateExecutiveSummary() {
    return await generateFinancialReport('executive', true);
}

/**
 * Shows dialog to send a report by email
 */
function showEmailReportDialog() {
    try {
        const htmlOutput = createTemplatedDialog(
            'UI_EmailReportDialog',
            { userEmail: Session.getActiveUser().getEmail() || "" },
            { width: 400, height: 400, title: 'Email Financial Report' }
        );
        
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Email Financial Report');
    } catch (error) {
        logError(`Error showing email report dialog: ${error.message}`);
        showError("Could not open email dialog: " + error.message);
    }
}

/**
 * Generates and emails a financial report
 * @param {string} recipient The recipient email
 * @param {string} reportType The type of report
 * @param {string} customMessage Optional custom message
 * @returns {string} Status message
 */
async function emailFinancialReport(recipient, reportType, customMessage) {
    try {
        const options = {
            reportType: reportType,
            includeTitle: true,
            reportTitle: `${reportType.charAt(0).toUpperCase() + reportType.slice(1)} Financial Report`,
            includeIntroduction: true,
            includeSummary: true,
            includeChart: true,
            includeDetailedAnalysis: true,
            includeAIResults: true,
            locale: getDefaultLocale()
        };
        
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getActiveSheet();
        const anomalies = await detectAnomalies(sheet);
        
        if (anomalies.length === 0) {
            return "No anomalies found to include in the report.";
        }
        
        // Generate the report
        const reportUrl = await createReport(anomalies, options);
        
        // Create and send email
        const success = await sendReportEmail(recipient, reportUrl, anomalies, customMessage);
        
        if (success) {
            return `Report sent successfully to ${recipient}`;
        } else {
            return "Error sending email. Please check the logs.";
        }
    } catch (error) {
        logError(`Error in emailFinancialReport: ${error.message}`);
        return `Error: ${error.message}`;
    }
}

/**
 * Detects and summarizes anomalies in the transaction data.
 * @param {number} threshold The threshold for anomaly detection
 * @returns {string} Summary of detected anomalies
 */
async function detectAndSummarizeAnomalies(threshold = 3) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    try {
        // Adjust config for this specific run if threshold is provided
        let config = getConfig();
        if (threshold) {
            config = {...config, 
                outliers: {
                    ...config.outliers, 
                    threshold: threshold
                }
            };
        }
        
        const anomalies = await detectAnomalies(sheet, config);
        if (anomalies.length === 0) {
            return "No anomalies detected in the current data.";
        }
        
        // Group anomalies by confidence level for better reporting
        const highConfidence = anomalies.filter(a => (a.confidence || 1.0) > 0.8);
        const mediumConfidence = anomalies.filter(a => {
            const conf = a.confidence || 1.0;
            return conf <= 0.8 && conf > 0.5;
        });
        const lowConfidence = anomalies.filter(a => (a.confidence || 1.0) <= 0.5);
        
        // Create a summary prompt that includes confidence levels
        const prompt = `Summarize the following financial anomalies:
        
        High confidence issues (${highConfidence.length}):
        ${JSON.stringify(highConfidence.slice(0, 5))}
        
        Medium confidence issues (${mediumConfidence.length}):
        ${JSON.stringify(mediumConfidence.slice(0, 5))}
        
        Low confidence issues (${lowConfidence.length}):
        ${JSON.stringify(lowConfidence.slice(0, 5))}
        
        Provide a concise summary of the key issues, patterns, and recommended actions.
        Include severity assessment and prioritize which issues to address first.`;
        
        return await generateReportAnalysis(prompt);
    } catch (error) {
        logError(`Error detecting anomalies: ${error.message}`);
        return `Error detecting anomalies: ${error.message}`;
    }
}

/**
 * Performs a monthly comparison of financial data.
 * @returns {string} Monthly comparison analysis
 */
async function performMonthlyComparison() {
    const sheetData = getSheetData();
    
    // Create a more detailed prompt for monthly comparison
    const prompt = `Compare the financial data month-by-month using this data:
    
    ${JSON.stringify(sheetData.slice(0, 100))}
    
    In your analysis, please include:
    1. Month-over-month spending trends
    2. Major changes in expense categories between months
    3. Unusual patterns or anomalies in monthly spending
    4. Revenue trends if present in the data
    5. Key metrics like monthly totals, averages, and growth rates
    6. Monthly financial health assessment
    
    Format your response with clear headings and bullet points for easy reading.
    Include quantitative analysis where possible with percentages and trend directions.`;
    
    return await generateResponse(prompt, sheetData);
}

/**
 * Runs the monthly comparison and displays the results
 */
async function runMonthlyComparison() {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert('Monthly Comparison', 
        'Analyzing monthly data. This may take a moment...', 
        ui.ButtonSet.OK);
    
    try {
        const analysis = await performMonthlyComparison();
        
        // Create a formatted HTML dialog to display the results
        const htmlContent = `
            <html>
                <head>
                    <base target="_top">
                    <style>
                        body { font-family: Arial, sans-serif; margin: 20px; }
                        h1 { color: #4CAF50; }
                        h2 { color: #2E7D32; border-bottom: 1px solid #ddd; padding-bottom: 8px; }
                        .analysis { line-height: 1.5; }
                    </style>
                </head>
                <body>
                    <h1>Monthly Comparison Analysis</h1>
                    <div class="analysis">${analysis.replace(/\n/g, '<br>')}</div>
                    
                    <div style="margin-top: 30px;">
                        <button onclick="google.script.run.withSuccessHandler(function() { 
                            google.script.host.close(); 
                        }).generateMonthlyComparisonReport()">
                            Generate Full Report
                        </button>
                    </div>
                </body>
            </html>
        `;
        
        const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
            .setWidth(800)
            .setHeight(600);
            
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Monthly Comparison Results');
    } catch (error) {
        logError(`Error in runMonthlyComparison: ${error.message}`);
        ui.alert('Error', `An error occurred during monthly comparison: ${error.message}`, ui.ButtonSet.OK);
    }
}

/**
 * Generates a full monthly comparison report - delegating to ReportGeneration.gs
 * @returns {string} URL to the created document
 */
async function generateMonthlyComparisonReport() {
    return await generateReportFromTemplate('monthly', true);
}

/**
 * Shows dialog for pattern analysis configuration
 */
function showPatternAnalysisDialog() {
    try {
        const htmlOutput = createTemplatedDialog(
            'UI_PatternAnalysisDialog',
            {},
            { width: 400, height: 350, title: 'Transaction Pattern Analysis' }
        );
        
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Transaction Pattern Analysis');
    } catch (error) {
        logError(`Error showing pattern analysis dialog: ${error.message}`);
        showError("Could not open pattern analysis: " + error.message);
    }
}

/**
 * Generates a pattern analysis report - delegating to ReportGeneration.gs
 * @param {string} analysisType Type of pattern analysis to perform
 * @param {boolean} includeVisuals Whether to include visualizations
 * @param {boolean} includeAI Whether to include AI insights
 * @returns {Promise<string>} HTML link to the created report
 */
async function generatePatternAnalysisReport(analysisType, includeVisuals, includeAI) {
    try {
        // Delegate to specialized function in ReportGeneration.gs
        return await generatePatternReport(analysisType, includeVisuals, includeAI);
    } catch (error) {
        logError(`Error generating pattern analysis report: ${error.message}`);
        throw new Error(`Failed to generate pattern analysis: ${error.message}`);
    }
}

/**
 * Handles QuickBooks configuration settings
 * @param {string} clientId QuickBooks client ID
 * @param {string} clientSecret QuickBooks client secret
 * @param {string} environment QuickBooks environment (SANDBOX/PRODUCTION)
 * @returns {string} Status message
 */
function setQuickBooksConfig(clientId, clientSecret, environment) {
    try {
        PropertiesService.getScriptProperties().setProperties({
            'QB_CLIENT_ID': clientId,
            'QB_CLIENT_SECRET': clientSecret,
            'QUICKBOOKS_ENV': environment
        });
        return 'QuickBooks configuration saved successfully.';
    } catch (error) {
        logError(`Error setting QuickBooks config: ${error.message}`);
        throw new Error('Failed to save QuickBooks configuration.');
    }
}

/**
 * Imports data from QuickBooks
 * @param {string} companyId QuickBooks company ID
 * @param {string} query Query to execute
 * @returns {Promise<string>} Status message
 */
async function importDataFromQuickBooks(companyId, query) {
    try {
        // Validate QuickBooks configuration
        const clientId = getQuickbooksClientId();
        const clientSecret = getQuickbooksClientSecret();
        
        if (!clientId || !clientSecret) {
            throw new Error('QuickBooks configuration is incomplete.');
        }

        // Initialize QuickBooks connection here
        // Note: Actual QuickBooks API implementation would go here
        const data = await fetchQuickBooksData(companyId, query);
        if (!data || data.length === 0) {
            throw new Error('No data returned from QuickBooks');
        }
        const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Imported QuickBooks Data');
        insertQuickBooksData(sheet, data);

        return `Imported ${data.length - 1} rows from QuickBooks.`;
    } catch (error) {
        logError(`QuickBooks import error: ${error.message}`);
        throw new Error(`Failed to import from QuickBooks: ${error.message}`);
    }
}

/**
 * Resets QuickBooks authorization
 * @returns {string} Status message
 */
function resetQuickBooksAuth() {
    try {
        // Clear OAuth tokens and other QuickBooks-specific properties
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.deleteProperty('QB_ACCESS_TOKEN');
        scriptProperties.deleteProperty('QB_REFRESH_TOKEN');
        scriptProperties.deleteProperty('QB_TOKEN_EXPIRY');
        
        return 'QuickBooks authorization has been reset.';
    } catch (error) {
        logError(`Error resetting QuickBooks auth: ${error.message}`);
        throw new Error('Failed to reset QuickBooks authorization.');
    }
}

/**
 * Shows configuration dialog using the enhanced HTML templating
 */
function showConfigDialog() {
    try {
        // Get current configuration settings
        const config = getConfig();
        
        const htmlOutput = createTemplatedDialog(
            'UI_ConfigDialog',
            {
                locale: getDefaultLocale(),
                currency: PropertiesService.getScriptProperties().getProperty('DEFAULT_CURRENCY') || 'USD',
                enableAI: PropertiesService.getScriptProperties().getProperty('ENABLE_AI') !== 'false',
                outlierThreshold: config.outliers?.threshold || 3,
                detectionAlgorithm: config.detectionAlgorithm || 'hybrid'
            },
            { width: 500, height: 400, title: 'Configuration Settings' }
        );
        
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Configuration Settings');
    } catch (error) {
        logError(`Error showing config dialog: ${error.message}`);
        showError("Could not open configuration: " + error.message);
    }
}

/**
 * Saves user configuration
 * @param {Object} config Configuration object
 * @returns {string} Status message
 */
function saveConfiguration(config) {
    try {
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperties({
            'DEFAULT_LOCALE': config.locale,
            'DEFAULT_CURRENCY': config.currency,
            'ENABLE_AI': config.enableAI.toString()
        });
        
        return 'Configuration saved successfully.';
    } catch (error) {
        logError(`Error saving configuration: ${error.message}`);
        throw new Error('Failed to save configuration.');
    }
}

/**
 * Handles GET requests when the application is deployed as a web app
 * @param {GoogleAppsScript.Events.DoGet} e The event object
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output
 */
function doGet(e) {
    try {
        // Initialize if needed (without checking user email)
        initializeIfNeeded();
        
        // Return the main UI with proper template evaluation
        return createTemplatedDialog(
            'UI_Main',
            {
                version: APP_VERSION,
                isWebApp: true
            },
            {
                title: 'Gemini Financial AI'
            }
        ).setFaviconUrl('https://www.google.com/images/chrome/apps/1x/sheets_32dp.png')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    } catch (error) {
        logError(`Error in doGet: ${error.message}`);
        return HtmlService.createHtmlOutput(`
            <div style="color: #c5221f; padding: 20px; font-family: Arial, sans-serif;">
                <h3>Error</h3>
                <p>An error occurred: ${sanitizeHTML(error.message)}</p>
                <p><a href="javascript:window.location.reload()">Reload page</a></p>
            </div>
        `);
    }
}

/**
 * Creates a new graph of data values in a new sheet
 */
function createDataGraph() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getActiveSheet();
    const range = sourceSheet.getDataRange();
    const data = range.getValues();
    
    // Check if we have enough data
    if (data.length <= 1) {
        SpreadsheetApp.getUi().alert("Error", "Not enough data to create a graph. Please select a range with headers and data.", SpreadsheetApp.getUi().ButtonSet.OK);
        return;
    }
    
    try {
        // Create a new sheet for the chart
        const chartSheet = spreadsheet.insertSheet("Data Graph " + new Date().toLocaleTimeString());
        
        // Copy the data to the new sheet for the chart
        chartSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
        
        // Create a chart
        const chart = chartSheet.newChart()
            .asColumnChart()
            .addRange(chartSheet.getRange(1, 1, data.length, data[0].length))
            .setPosition(5, 5, 0, 0)
            .setOption('title', 'Data Visualization')
            .setOption('legend', {position: 'bottom'})
            .setOption('hAxis.title', data[0][0])
            .setOption('vAxis.title', 'Value')
            .setOption('colors', ['#4285F4', '#34A853', '#FBBC05', '#EA4335'])
            .build();
            
        chartSheet.insertChart(chart);
        
        SpreadsheetApp.getUi().alert("Success", "Chart created in a new sheet!", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (error) {
        logError(`Error creating chart: ${error.message}`);
        SpreadsheetApp.getUi().alert("Error", `Failed to create chart: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * Logs the redirect URI for OAuth2 configuration
 */
function logRedirectUri() {
    return getRedirectUri();
}

/**
 * Initializes the application if it hasn't been initialized yet
 */
function initializeIfNeeded() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const isInitialized = scriptProperties.getProperty('IS_INITIALIZED');
    
    if (!isInitialized) {
        try {
            // Set default configuration
            setConfig({
                locale: 'en-US',
                currency: 'USD',
                enableAI: true
            });
            
            // Set default model selections
            PropertiesService.getScriptProperties().setProperties({
                'DEFAULT_TEXT_MODEL': 'gemini-1.5-pro-latest',
                'DEFAULT_VISION_MODEL': 'gemini-1.5-pro-vision-latest'
            });
            
            // Mark as initialized
            scriptProperties.setProperty('IS_INITIALIZED', 'true');
            logMessage('Application initialized with default settings');
        } catch (error) {
            logError(`Error in initialization: ${error.message}`);
            // Continue anyway to avoid blocking the web app from loading
        }
    }
}

/**
 * Shows the Gemini configuration dialog
 */
function showGeminiConfigDialog() {
    try {
        const apiKey = getGeminiAPIKey();
        const htmlOutput = createTemplatedDialog(
            'UI_GeminiConfig',
            {
                apiKey: apiKey || '',
                apiConfigured: Boolean(apiKey),
                useFunctionCalling: PropertiesService.getScriptProperties().getProperty('USE_FUNCTION_CALLING_API') === 'true'
            },
            { width: 450, height: 350, title: 'Gemini AI Configuration' }
        );
        
        SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gemini AI Configuration');
    } catch (error) {
        logError(`Error showing Gemini config dialog: ${error.message}`);
        showError("Could not open Gemini configuration: " + error.message);
    }
}

/**
 * Saves the Gemini API key and function calling preference
 * @param {string} apiKey The Gemini API key
 * @param {boolean} useFunctionCalling Whether to use function calling API
 * @returns {string} Status message
 */
function saveGeminiConfig(apiKey, useFunctionCalling) {
    try {
        PropertiesService.getScriptProperties().setProperties({
            'GEMINI_API_KEY': apiKey,
            'USE_FUNCTION_CALLING_API': useFunctionCalling.toString()
        });
        
        // Test the API key with a simple request
        if (apiKey) {
            const testResult = testGeminiApiKey(apiKey);
            if (!testResult.success) {
                return `API key saved, but test failed: ${testResult.message}`;
            }
        }
        
        return "Gemini configuration saved successfully.";
    } catch (error) {
        logError(`Error saving Gemini config: ${error.message}`);
        throw new Error("Failed to save Gemini configuration: " + error.message);
    }
}

/**
 * Tests if a Gemini API key is valid
 * @param {string} apiKey The API key to test
 * @returns {Object} Result with success status and message
 */
function testGeminiApiKey(apiKey) {
    try {
        const endpoint = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
        const response = UrlFetchApp.fetch(endpoint, {
            method: 'get',
            muteHttpExceptions: true
        });
        
        const code = response.getResponseCode();
        if (code !== 200) {
            return {
                success: false,
                message: `API test failed with code ${code}`
            };
        }
        
        return {
            success: true,
            message: "API key validated successfully"
        };
    } catch (error) {
        return {
            success: false,
            message: error.message
        };
    }
}

/**
 * Gets default config with helpful defaults
 * @returns {Object} Default configuration
 */
function getDefaultConfig() {
    return {
        locale: 'en-US',
        currency: 'USD',
        detectionAlgorithm: 'hybrid',
        enableAIDetection: true,
        outliers: {
            threshold: 3,
            method: 'zscore'
        },
        amount: {
            min: 0,
            max: 10000,
            allowNegative: true
        },
        date: {
            allowFuture: false,
            datePatterns: [
                /^\d{4}-\d{2}-\d{2}$/,  // yyyy-mm-dd
                /^\d{2}\/\d{2}\/\d{4}$/ // mm/dd/yyyy
            ]
        },
        description: {
            required: true
        },
        category: {
            required: false,
            validCategories: ['Sales', 'Marketing', 'Development', 'HR', 'Operations', 'Other']
        },
        email: {
            required: false,
            format: /^[^\s@]+@[^\s@]+\.[^\s@]+$/
        },
        duplicates: {
            check: true,
            uniqueColumns: ['amount', 'date', 'description']
        },
        mandatoryFields: ['amount', 'date', 'description'],
        roundNumberThreshold: 100,
        flagWeekendTransactions: true,
        includeAIExplanations: true
    };
}

/**
 * Displays an error message to the user
 * @param {string} errorMessage The error message to display
 */
function showError(errorMessage) {
    if (!errorMessage) return;
    
    try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Error', errorMessage, ui.ButtonSet.OK);
    } catch (e) {
        // If UI isn't available, just log the error
        Logger.log('ERROR: ' + errorMessage);
    }
}

/**
 * Log a message for debugging
 * @param {string} message The message to log
 */
function logMessage(message) {
    Logger.log(message);
}

/**
 * Log an error message for debugging
 * @param {string} errorMessage The error message to log
 */
function logError(errorMessage) {
    Logger.log('ERROR: ' + errorMessage);
}

/**
 * Saves Gemini model selections to script properties and refreshes available models
 * @param {string} textModel The selected text model 
 * @param {string} visionModel The selected vision model
 * @returns {string} Status message
 */
function saveGeminiModelSelections(textModel, visionModel) {
    try {
        PropertiesService.getScriptProperties().setProperties({
            'GEMINI_TEXT_MODEL': textModel,
            'GEMINI_VISION_MODEL': visionModel
        });
        
        // Try refreshing models in the background
        refreshModelsInBackground();
        
        return "Model selections saved successfully.";
    } catch (error) {
        logError(`Error saving model selections: ${error.message}`);
        throw new Error("Failed to save model selections: " + error.message);
    }
}

/**
 * Shows the QuickBooks Import dialog using the improved template handling
 */
function showQuickBooksImportDialog() {
  try {
    // Create the HTML template with style inclusion
    const html = createTemplatedDialog(
      'UI_QuickBooksImportDialog', 
      {
        isConfigured: isQuickBooksConfigured(),
        environment: getQuickBooksEnvironment()
      }, 
      { width: 400, height: 350, title: 'QuickBooks Import' }
    );
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (error) {
    logError(`Error showing QuickBooks import dialog: ${error.message}`);
    showError("Could not open QuickBooks import: " + error.message);
  }
}

/**
 * Processes the CSV text from QuickBooks export.
 * @param {string} csvContent The CSV file content.
 * @return {string} Simple confirmation or error message.
 */
function processQuickBooksImport(csvContent) {
  try {
    // Simple validation
    if (!csvContent || csvContent.trim() === '') {
      throw new Error('No CSV data provided');
    }
    
    // Parse the CSV content
    const lines = csvContent.split(/\r?\n/);
    if (lines.length < 2) {
      throw new Error('CSV file contains insufficient data');
    }
    
    // First line should be headers
    const headers = lines[0].split(',');
    if (headers.length < 3) {
      throw new Error('CSV file should have at least 3 columns for valid transaction data');
    }
    
    // Parse the remaining lines as rows
    const rows = lines.slice(1)
      .filter(line => line.trim() !== '') // Skip empty lines
      .map(line => line.split(','));
    
    // Get the active spreadsheet and create a new sheet for imported data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'QB Import ' + new Date().toLocaleDateString();
    
    // Create a new sheet and insert data
    let sheet;
    try {
      sheet = ss.insertSheet(sheetName);
    } catch (e) {
      // If sheet with same name exists, create with timestamp
      sheet = ss.insertSheet(`${sheetName} ${new Date().toLocaleTimeString()}`);
    }
    
    // Prepare all data (headers + rows)
    const allData = [headers, ...rows];
    
    // Insert data into the sheet
    sheet.getRange(1, 1, allData.length, headers.length).setValues(allData);
    
    // Format the header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, headers.length);
    
    // Format date and currency columns if we can detect them
    headers.forEach((header, colIndex) => {
      const headerLower = header.toLowerCase();
      if (headerLower.includes('date')) {
        // Format as date
        sheet.getRange(2, colIndex + 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
      } else if (headerLower.includes('amount') || 
                headerLower.includes('total') || 
                headerLower.includes('price') ||
                headerLower.includes('cost')) {
        // Format as currency
        sheet.getRange(2, colIndex + 1, rows.length, 1).setNumberFormat('$#,##0.00');
      }
    });
    
    return `Successfully imported ${rows.length} transactions from QuickBooks to "${sheet.getName()}"`;
  } catch (err) {
    logError('QuickBooks Import parsing error: ' + err.message);
    throw new Error('Failed to parse QuickBooks CSV: ' + err.message);
  }
}

/**
 * Formats a value (number, date) according to locale settings
 * @param {any} value The value to format
 * @param {string} type Format type (currency, date, number, percent)
 * @param {Object} options Additional formatting options
 * @returns {string} Formatted value as string
 */
function formatLocalized(value, type = 'number', options = {}) {
  switch (type) {
    case 'date':
      return formatLocalizedDate(value, options.locale, options.format);
    case 'currency':
    case 'number':
    case 'percent':
      return formatLocalizedValue(value, type, options);
    default:
      return String(value);
  }
}

/**
 * Analyzes selected data range and provides insights
 */
function analyzeSelectedData() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getSelection();
  
  // Check if there's an active range selected
  const activeRange = selection.getActiveRange();
  if (!activeRange) {
    ui.alert('No Selection', 'Please select a range of data to analyze.', ui.ButtonSet.OK);
    return;
  }
  
  // Get the selected data
  const data = activeRange.getValues();
  
  // Check if there's enough data to analyze
  if (data.length <= 1 || data[0].length === 0) {
    ui.alert('Insufficient Data', 'Please select a larger range of data to analyze.', ui.ButtonSet.OK);
    return;
  }
  
  // Show loading message
  ui.alert('Analyzing Selection', 'Analyzing the selected data. This may take a moment...', ui.ButtonSet.OK);
  
  // Run the analysis
  analyzeDataSelection(data)
    .then(analysis => {
      // Show results in a dialog
      const htmlContent = `
        <html>
          <head>
            <base target="_top">
            <style>
              body { font-family: Arial, sans-serif; margin: 20px; line-height: 1.6; }
              h1 { color: #4285f4; }
              .analysis { background-color: #f8f9fa; padding: 15px; border-radius: 8px; }
              .actions { margin-top: 20px; }
              button { background-color: #4285f4; color: white; border: none; padding: 8px 16px;
                      border-radius: 4px; cursor: pointer; }
              button:hover { background-color: #2b76e5; }
            </style>
          </head>
          <body>
            <h1>Selection Analysis</h1>
            <div class="analysis">${analysis.replace(/\n/g, '<br>')}</div>
            <div class="actions">
              <button onclick="google.script.run.createSelectionVisualization(); google.script.host.close();">
                Create Visualization
              </button>
              <button onclick="google.script.host.close();">Close</button>
            </div>
          </body>
        </html>
      `;
      
      const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
        .setWidth(600)
        .setHeight(500);
        
      ui.showModalDialog(htmlOutput, 'Data Analysis Results');
    })
    .catch(error => {
      ui.alert('Error', `An error occurred during analysis: ${error.message}`, ui.ButtonSet.OK);
      logError(`Error analyzing selection: ${error.message}`);
    });
}

/**
 * Creates visualizations based on the currently selected data
 */
function createSelectionVisualization() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getSelection();
  const activeRange = selection.getActiveRange();
  
  if (!activeRange) {
    ui.alert('No Selection', 'Please select data to visualize.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const data = activeRange.getValues();
    const headers = data[0];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const chartSheet = ss.insertSheet('Visualization ' + new Date().toLocaleTimeString());
    
    // Copy the data
    chartSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Create a chart - determine the best type based on data
    let chartType;
    if (data[0].length === 2) {
      chartType = Charts.ChartType.PIE;
    } else {
      chartType = Charts.ChartType.COLUMN;
    }
    
    const chartBuilder = chartSheet.newChart()
      .setChartType(chartType)
      .addRange(chartSheet.getRange(1, 1, data.length, Math.min(5, data[0].length)))
      .setPosition(data.length + 5, 1, 0, 0)
      .setOption('title', 'Data Visualization')
      .setOption('width', 600)
      .setOption('height', 400)
      .setOption('legend', { position: 'bottom' });
      
    // Add the chart
    chartSheet.insertChart(chartBuilder.build());
    
    ui.alert('Success', 'Visualization created in a new sheet!', ui.ButtonSet.OK);
  } catch (error) {
    logError(`Error creating visualization: ${error.message}`);
    ui.alert('Error', `Failed to create visualization: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Analyzes the selected data and returns insights with enhanced date handling
 * @param {Array<Array<any>>} data The selected data to analyze
 * @returns {Promise<string>} A promise that resolves to the analysis text
 */
async function analyzeDataSelection(data) {
  try {
    // Create a more sophisticated data analysis that accounts for cell types
    const headers = data[0];
    const rows = data.slice(1);
    const rowCount = rows.length;
    
    // Detect column types
    const columnTypes = detectColumnTypes(data);
    
    // Create column summaries based on detected types
    const columnSummaries = [];
    const dateRangeSummaries = [];
    
    headers.forEach((header, index) => {
      const columnType = columnTypes[index];
      const columnValues = rows.map(row => row[index]);
      
      let summary = `Column "${header}" (${columnType}): `;
      
      switch(columnType) {
        case 'date':
          // Handle dates separately to avoid including them in numeric calculations
          const rawDates = columnValues.filter(v => v !== null && v !== undefined && v !== '');
          const validDates = rawDates.filter(v => v instanceof Date && !isNaN(v.getTime()))
                            .map(d => new Date(d));
          
          // Add valid date conversions for values that might be stored as numbers or strings
          columnValues.forEach(val => {
            if (!(val instanceof Date) && val !== null && val !== undefined && val !== '') {
              try {
                const potentialDate = new Date(val);
                if (!isNaN(potentialDate.getTime())) {
                  validDates.push(potentialDate);
                }
              } catch (e) {
                // Not convertible to date, ignore
              }
            }
          });
          
          if (validDates.length > 0) {
            const minDate = new Date(Math.min(...validDates.map(d => d.getTime())));
            const maxDate = new Date(Math.max(...validDates.map(d => d.getTime())));
            
            // Format dates in a readable way
            const minDateStr = formatLocalizedDate(minDate, 'en-US', 'short');
            const maxDateStr = formatLocalizedDate(maxDate, 'en-US', 'short');
            
            // Calculate date span in days
            const daySpan = Math.round((maxDate - minDate) / (1000 * 60 * 60 * 24));
            
            summary += `Range from ${minDateStr} to ${maxDateStr} (${daySpan} days), ${validDates.length} valid dates`;
            
            // Add to specific date summaries collection
            dateRangeSummaries.push({
              header,
              minDate: minDateStr,
              maxDate: maxDateStr,
              daySpan,
              count: validDates.length
            });
          } else {
            summary += 'No valid dates found';
          }
          break;
        
        case 'number':
          // Ensure we're only using actual numbers, not dates stored as numbers
          const numbers = columnValues.filter(v => {
            // Skip if it's a date or convertible to a valid date
            if (v instanceof Date) return false;
            
            // Check if it's a number that could represent a date timestamp
            // (extremely large or precise numbers often indicate timestamps)
            if (typeof v === 'number' && v > 946684800000) { // Jan 1, 2000 timestamp
              try {
                const asDate = new Date(v);
                // If it converts to a reasonable date in the last ~50 years, it's likely a timestamp
                if (!isNaN(asDate.getTime()) && 
                    asDate.getFullYear() > 1970 && 
                    asDate.getFullYear() < 2050) {
                  return false;
                }
              } catch (e) {
                // Not a date, continue
              }
            }
            
            return typeof v === 'number' || (!isNaN(parseFloat(v)) && isFinite(v));
          }).map(v => typeof v === 'number' ? v : parseFloat(v));
          
          if (numbers.length > 0) {
            const sum = numbers.reduce((a, b) => a + b, 0);
            const avg = sum / numbers.length;
            const min = Math.min(...numbers);
            const max = Math.max(...numbers);
            
            // Format large numbers for better readability
            const formatNumber = (num) => {
              if (Math.abs(num) >= 1000000) {
                return `${(num / 1000000).toFixed(2)}M`;
              } else if (Math.abs(num) >= 1000) {
                return `${(num / 1000).toFixed(2)}K`;
              }
              return num.toFixed(2);
            };
            
            summary += `Range: ${formatNumber(min)} to ${formatNumber(max)}, ` +
                      `Average: ${formatNumber(avg)}, ` + 
                      `Sum: ${formatNumber(sum)}, ` +
                      `Count: ${numbers.length}`;
          } else {
            summary += 'No valid numbers found';
          }
          break;
          
        // Rest of the cases remain the same
        case 'boolean':
          // ...existing code...
          const trueCount = columnValues.filter(v => v === true || v === 'TRUE' || v === 'true').length;
          const falseCount = columnValues.filter(v => v === false || v === 'FALSE' || v === 'false').length;
          summary += `${trueCount} true values, ${falseCount} false values`;
          break;
          
        case 'text':
        default:
          // ...existing code...
          const uniqueValues = new Set(columnValues.filter(v => v !== null && v !== undefined).map(String));
          summary += `${uniqueValues.size} unique values out of ${columnValues.length} total values`;
          
          // Word count analysis for text
          if (columnType === 'text') {
            const wordCounts = columnValues
              .filter(v => typeof v === 'string')
              .map(v => v.split(/\s+/).filter(word => word.trim().length > 0).length);
            
            if (wordCounts.length > 0) {
              const totalWords = wordCounts.reduce((a, b) => a + b, 0);
              const avgWords = totalWords / wordCounts.length;
              summary += `, Avg words per cell: ${avgWords.toFixed(1)}`;
            }
          }
          break;
      }
      
      columnSummaries.push(summary);
    });
    
    // Create date range specific section if we have date columns
    let dateRangeSection = '';
    if (dateRangeSummaries.length > 0) {
      dateRangeSection = `\nDate Ranges:\n` + 
        dateRangeSummaries.map(dr => 
          `- ${dr.header}: ${dr.minDate} to ${dr.maxDate} (${dr.daySpan} days, ${dr.count} entries)`
        ).join('\n');
    }
    
    // Create an enhanced prompt including data types and separate date section
    const prompt = `
      Analyze the following spreadsheet data (${rowCount} rows):
      
      Column summaries:
      ${columnSummaries.join('\n')}
      ${dateRangeSection}
      
      Sample data:
      Headers: ${headers.join(', ')}
      ${JSON.stringify(rows.slice(0, Math.min(5, rows.length)))}
      
      Please provide a comprehensive analysis considering the data types and patterns, including:
      1. A descriptive overview of what the data represents
      2. Key statistical insights tailored to each data type (dates, numbers, text, etc.)
      3. For date columns, analyze time periods and any temporal patterns
      4. For numeric columns, analyze distributions, trends and outliers
      5. Patterns, trends, or correlations between columns
      6. Data quality issues (missing values, inconsistencies)
      7. Brief conclusions and recommendations for further analysis
      
      Format your response with clear headings for each section and bullet points where appropriate.
    `;
    
    // Generate analysis using Gemini
    const analysis = await generateResponse(prompt, []);
    return analysis;
  } catch (error) {
    logError(`Error analyzing selection: ${error.message}`);
    throw new Error(`Failed to analyze selection: ${error.message}`);
  }
}

/**
 * Enhanced date detection to better identify different date formats
 * @param {Array<Array<any>>} data The data to analyze
 * @returns {Array<string>} Array of column types
 */
function detectColumnTypes(data) {
  // Skip the header row
  const rows = data.slice(1);
  if (rows.length === 0) return [];
  
  const columnCount = data[0].length;
  const columnTypes = [];
  
  for (let col = 0; col < columnCount; col++) {
    const values = rows.map(row => row[col]).filter(val => val !== null && val !== undefined && val !== '');
    
    if (values.length === 0) {
      columnTypes.push('unknown');
      continue;
    }
    
    // Count occurrences of each type
    let dateCount = 0;
    let numberCount = 0;
    let booleanCount = 0;
    let textCount = 0;
    
    values.forEach(val => {
      // Check if it's already a date object
      if (val instanceof Date) {
        dateCount++;
        return;
      } 
      
      // Check common date patterns
      if (typeof val === 'string') {
        // ISO dates: 2023-01-15, MM/DD/YYYY, DD/MM/YYYY, etc.
        const datePatterns = [
          /^\d{4}-\d{2}-\d{2}$/,
          /^\d{1,2}\/\d{1,2}\/\d{4}$/,
          /^\d{1,2}-\d{1,2}-\d{4}$/,
          /^\d{1,2}\.\d{1,2}\.\d{4}$/,
          /^\d{2}-[A-Za-z]{3}-\d{4}$/,
          /^\d{2}\/[A-Za-z]{3}\/\d{4}$/
        ];
        
        for (const pattern of datePatterns) {
          if (pattern.test(val)) {
            try {
              const testDate = new Date(val);
              if (!isNaN(testDate.getTime())) {
                dateCount++;
                return;
              }
            } catch (e) {
              // Not a valid date, continue
            }
          }
        }
      }
      
      // Check if it might be a timestamp (large number that converts to a reasonable date)
      if (typeof val === 'number' && val > 946684800000) { // Jan 1, 2000 timestamp
        try {
          const asDate = new Date(val);
          if (!isNaN(asDate.getTime()) && 
             asDate.getFullYear() > 1970 && 
             asDate.getFullYear() < 2050) {
            dateCount++;
            return;
          }
        } catch (e) {
          // Not a date, continue
        }
      }
      
      // Check if it's a boolean
      if (val === true || val === false || val === 'TRUE' || val === 'FALSE' || val === 'true' || val === 'false') {
        booleanCount++;
        return;
      }
      
      // Check if it's a number
      if (typeof val === 'number' || (!isNaN(parseFloat(val)) && isFinite(val))) {
        numberCount++;
        return;
      }
      
      // Must be text
      textCount++;
    });
    
    // Determine the most common type
    const typeCountMap = {
      'date': dateCount,
      'number': numberCount,
      'boolean': booleanCount,
      'text': textCount
    };
    
    // Find the type with the highest count
    let maxType = 'text'; // Default to text
    let maxCount = 0;
    
    Object.entries(typeCountMap).forEach(([type, count]) => {
      if (count > maxCount) {
        maxType = type;
        maxCount = count;
      }
    });
    
    // If more than 70% of values match the max type, use that
    if (maxCount / values.length >= 0.7) {
      columnTypes.push(maxType);
    } else {
      // Mixed type, default to text
      columnTypes.push('text');
    }
  }
  
  return columnTypes;
}

/**
 * Creates visualizations based on the currently selected data with enhanced type detection
 */
function createSelectionVisualization() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getSelection();
  const activeRange = selection.getActiveRange();
  
  if (!activeRange) {
    ui.alert('No Selection', 'Please select data to visualize.', ui.ButtonSet.OK);
    return;
  }
  
  try {
    const data = activeRange.getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    if (rows.length === 0) {
      ui.alert('Insufficient Data', 'Please select data with at least one row of content.', ui.ButtonSet.OK);
      return;
    }
    
    // Detect column types
    const columnTypes = detectColumnTypes(data);
    
    // Create a new sheet for the visualization
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const chartSheet = ss.insertSheet('Visualization ' + new Date().toLocaleTimeString());
    
    // Copy the data
    chartSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    
    // Auto-format columns based on detected types
    columnTypes.forEach((type, index) => {
      const column = index + 1;
      
      switch(type) {
        case 'date':
          chartSheet.getRange(2, column, rows.length, 1).setNumberFormat('yyyy-mm-dd');
          break;
        case 'number':
          chartSheet.getRange(2, column, rows.length, 1).setNumberFormat('#,##0.00');
          break;
        case 'boolean':
          // No special formatting for booleans
          break;
        default:
          // No special formatting for text
          break;
      }
    });
    
    // Determine the most appropriate chart type based on data types
    let chartBuilder;
    
    // Find numeric columns
    const numericColumns = columnTypes.map((type, index) => ({ type, index }))
                                      .filter(col => col.type === 'number');
    
    // Find date columns
    const dateColumns = columnTypes.map((type, index) => ({ type, index }))
                                  .filter(col => col.type === 'date');
    
    // Find text/category columns
    const categoryColumns = columnTypes.map((type, index) => ({ type, index }))
                                      .filter(col => col.type === 'text' || col.type === 'boolean');
    
    if (dateColumns.length >= 1 && numericColumns.length >= 1) {
      // Create a time series chart (line chart) with dates on X-axis
      const dateColIndex = dateColumns[0].index;
      const numericColIndex = numericColumns[0].index;
      
      chartBuilder = chartSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartSheet.getRange(1, dateColIndex + 1, rows.length + 1, 1)) // Date column
        .addRange(chartSheet.getRange(1, numericColIndex + 1, rows.length + 1, 1)) // Value column
        .setOption('title', `${headers[numericColIndex]} Over Time`)
        .setOption('hAxis.title', headers[dateColIndex])
        .setOption('vAxis.title', headers[numericColIndex]);
    }
    else if (numericColumns.length >= 2) {
      // Create a scatter chart for two numeric columns
      chartBuilder = chartSheet.newChart()
        .setChartType(Charts.ChartType.SCATTER)
        .addRange(chartSheet.getRange(1, numericColumns[0].index + 1, rows.length + 1, 1))
        .addRange(chartSheet.getRange(1, numericColumns[1].index + 1, rows.length + 1, 1))
        .setOption('title', `${headers[numericColumns[0].index]} vs ${headers[numericColumns[1].index]}`)
        .setOption('hAxis.title', headers[numericColumns[0].index])
        .setOption('vAxis.title', headers[numericColumns[1].index]);
    }
    else if (categoryColumns.length >= 1 && numericColumns.length >= 1) {
      // Create a column chart for category vs numeric
      chartBuilder = chartSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(chartSheet.getRange(1, categoryColumns[0].index + 1, rows.length + 1, 1))
        .addRange(chartSheet.getRange(1, numericColumns[0].index + 1, rows.length + 1, 1))
        .setOption('title', `${headers[numericColumns[0].index]} by ${headers[categoryColumns[0].index]}`)
        .setOption('hAxis.title', headers[categoryColumns[0].index])
        .setOption('vAxis.title', headers[numericColumns[0].index]);
    }
    else if (categoryColumns.length >= 1) {
      // Create a pie chart for single category column (frequency)
      // First calculate category frequencies
      const categoryIndex = categoryColumns[0].index;
      const categoryValues = rows.map(row => String(row[categoryIndex]));
      
      // Count frequencies
      const frequencies = {};
      categoryValues.forEach(value => {
        frequencies[value] = (frequencies[value] || 0) + 1;
      });
      
      // Add frequency data to sheet
      const freqSheet = ss.insertSheet('Category Frequencies');
      freqSheet.appendRow(['Category', 'Frequency']);
      
      Object.entries(frequencies).forEach(([category, frequency], index) => {
        freqSheet.appendRow([category, frequency]);
      });
      
      chartBuilder = freqSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(freqSheet.getDataRange())
        .setOption('title', `${headers[categoryIndex]} Distribution`)
        .setOption('pieSliceText', 'percentage');
    }
    else if (numericColumns.length === 1) {
      // Create a histogram for single numeric column
      chartBuilder = chartSheet.newChart()
        .setChartType(Charts.ChartType.HISTOGRAM)
        .addRange(chartSheet.getRange(2, numericColumns[0].index + 1, rows.length, 1))
        .setOption('title', `Distribution of ${headers[numericColumns[0].index]}`)
        .setOption('hAxis.title', headers[numericColumns[0].index])
        .setOption('vAxis.title', 'Frequency');
    }
    else {
      // Default to simple column chart if no special case is detected
      chartBuilder = chartSheet.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(chartSheet.getDataRange())
        .setOption('title', 'Data Visualization')
        .setOption('legend', { position: 'bottom' });
    }
    
    // Set common chart properties
    chartBuilder.setPosition(data.length + 5, 1, 0, 0)
                .setOption('width', 700)
                .setOption('height', 500);
                
    // Add the chart
    chartSheet.insertChart(chartBuilder.build());
    
    // Give feedback to the user
    ui.alert('Success', 'Visualization created in a new sheet!', ui.ButtonSet.OK);
  } catch (error) {
    logError(`Error creating visualization: ${error.message}`);
    ui.alert('Error', `Failed to create visualization: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Safely appends content to a report document
 * @param {string} docUrl The URL of the document
 * @param {string} content The content to append
 * @returns {boolean} True if successful, false otherwise
 */
function appendPatternAnalysisToReport(docUrl, content) {
  try {
    // First validate the URL
    const validUrl = validateAndNormalizeDocUrl(docUrl);
    if (!validUrl) {
      logError(`Invalid document URL format: ${docUrl}`);
      return false;
    }
    
    // Try to append content
    return appendToReport(validUrl, content);
  } catch (error) {
    logError(`Error appending to report: ${error.message}`);
    return false;
  }
}

/**
 * Schedules automatic reports based on user settings
 * @param {string} frequency How often to run reports ('daily', 'weekly', 'monthly')
 * @param {Object} options Scheduling options including email, report type, etc.
 * @returns {string} Status message confirming the schedule
 */
function scheduleAutomaticReports(frequency, options = {}) {
  try {
    // Delete any existing report triggers
    deleteExistingTriggers('generateScheduledReport');
    
    if (frequency === 'none') {
      return 'Automatic reports have been disabled.';
    }
    
    // Store report options
    PropertiesService.getScriptProperties().setProperties({
      'SCHEDULED_REPORT_FREQUENCY': frequency,
      'SCHEDULED_REPORT_TYPE': options.reportType || 'standard',
      'SCHEDULED_REPORT_EMAIL': options.email || '',
      'SCHEDULED_REPORT_INCLUDE_AI': (options.includeAI !== false).toString()
    });
    
    let trigger;
    switch (frequency.toLowerCase()) {
      case 'daily':
        trigger = ScriptApp.newTrigger('generateScheduledReport')
          .timeBased()
          .everyDays(1)
          .atHour(options.hour || 6)
          .create();
        break;
      case 'weekly':
        trigger = ScriptApp.newTrigger('generateScheduledReport')
          .timeBased()
          .everyWeeks(1)
          .onWeekDay(options.weekDay || ScriptApp.WeekDay.MONDAY)
          .atHour(options.hour || 6)
          .create();
        break;
      case 'monthly':
        trigger = ScriptApp.newTrigger('generateScheduledReport')
          .timeBased()
          .everyDays(30)
          .atHour(options.hour || 6)
          .create();
        break;
      default:
        throw new Error(`Invalid frequency: ${frequency}`);
    }
    
    return `Reports scheduled to run ${frequency}. Next report will be generated automatically.`;
    
  } catch (error) {
    logError(`Error scheduling reports: ${error.message}`);
    throw new Error(`Failed to schedule reports: ${error.message}`);
  }
}

/**
 * Generates scheduled reports and emails them to recipients
 */
async function generateScheduledReport() {
  try {
    logMessage('Starting scheduled report generation');
    
    // Get stored report options
    const scriptProps = PropertiesService.getScriptProperties();
    const reportType = scriptProps.getProperty('SCHEDULED_REPORT_TYPE') || 'standard';
    const recipientEmail = scriptProps.getProperty('SCHEDULED_REPORT_EMAIL');
    const includeAI = scriptProps.getProperty('SCHEDULED_REPORT_INCLUDE_AI') !== 'false';
    
    if (!recipientEmail) {
      logError('No recipient email configured for scheduled reports');
      return;
    }
    
    // Generate the report using the specialized function
    const reportUrl = await generateReportFromTemplate(reportType, includeAI);
    
    // Send the email with the report
    const subject = `${reportType.charAt(0).toUpperCase() + reportType.slice(1)} Financial Report - ${new Date().toLocaleDateString()}`;
    const message = `Your scheduled financial report is ready.

View the report here: ${reportUrl}

This is an automated message from Gemini Financial AI.`;
    
    GmailApp.sendEmail(recipientEmail, subject, message);
    
    logMessage(`Scheduled report (${reportType}) sent to ${recipientEmail}`);
  } catch (error) {
    logError(`Error in scheduled report generation: ${error.message}`);
    
    // Try to send error notification to admin
    try {
      const adminEmail = Session.getActiveUser().getEmail();
      if (adminEmail) {
        GmailApp.sendEmail(
          adminEmail,
          'Error: Scheduled Report Failed',
          `The scheduled report generation failed with error: ${error.message}\n\nPlease check the application logs.`
        );
      }
    } catch (emailError) {
      logError(`Failed to send error notification: ${emailError.message}`);
    }
  }
}

/**
 * Creates a scheduled anomaly detection job based on user settings
 * @param {Object} schedulingOptions Options for scheduling (frequency, email, etc)
 * @returns {string} Status message
 */
function createScheduledAnomalyDetection(schedulingOptions) {
  try {
    return setupScheduledAnomalyDetection(
      schedulingOptions.frequency,
      {
        hour: schedulingOptions.hour,
        weekDay: schedulingOptions.weekDay,
        notificationEmail: schedulingOptions.email
      }
    );
  } catch (error) {
    logError(`Error creating scheduled anomaly detection: ${error.message}`);
    throw new Error(`Failed to create scheduled job: ${error.message}`);
  }
}

/**
 * Shows dialog for scheduling automatic reports
 */
function showScheduleReportsDialog() {
  try {
    // Get current scheduling configuration
    const scriptProps = PropertiesService.getScriptProperties();
    const frequency = scriptProps.getProperty('SCHEDULED_REPORT_FREQUENCY') || 'none';
    const reportType = scriptProps.getProperty('SCHEDULED_REPORT_TYPE') || 'standard';
    const email = scriptProps.getProperty('SCHEDULED_REPORT_EMAIL') || '';
    const includeAI = scriptProps.getProperty('SCHEDULED_REPORT_INCLUDE_AI') !== 'false';
    
    const htmlOutput = createTemplatedDialog(
      'UI_ScheduleReports',
      {
        frequency,
        reportType,
        email,
        includeAI,
        userEmail: Session.getActiveUser().getEmail()
      },
      { width: 450, height: 500, title: 'Schedule Automatic Reports' }
    );
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Schedule Automatic Reports');
  } catch (error) {
    logError(`Error showing schedule dialog: ${error.message}`);
    showError("Could not open scheduling dialog: " + error.message);
  }
}

/**
 * Analyzes categories with provided options
 * @param {number} topCategories Number of top categories to analyze
 * @param {string} period Time period to analyze
 * @returns {Promise<string>} Analysis results
 */
async function analyzeCategories(topCategories = 5, period = 'all') {
  try {
    const sheetData = getSheetData();
    
    // Find if there's a category column
    const headers = sheetData[0].map(h => String(h).toLowerCase());
    const categoryIndex = headers.findIndex(h => h === 'category' || h.includes('categor'));
    
    if (categoryIndex === -1) {
      return "No category column found in your sheet. Please ensure you have a column labeled 'Category' or similar.";
    }
    
    // Extract dates if period filtering is needed
    const dateIndex = headers.findIndex(h => h.includes('date'));
    let filteredData = sheetData.slice(1); // Skip header row
    
    // Filter by period if specified
    if (period !== 'all' && dateIndex !== -1) {
      const now = new Date();
      let startDate;
      
      switch(period.toLowerCase()) {
        case 'this month':
          startDate = new Date(now.getFullYear(), now.getMonth(), 1);
          break;
        case 'last month':
          startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
          const endOfLastMonth = new Date(now.getFullYear(), now.getMonth(), 0);
          filteredData = filteredData.filter(row => {
            const rowDate = new Date(row[dateIndex]);
            return rowDate >= startDate && rowDate <= endOfLastMonth;
          });
          break;
        case 'this year':
          startDate = new Date(now.getFullYear(), 0, 1);
          filteredData = filteredData.filter(row => {
            const rowDate = new Date(row[dateIndex]);
            return rowDate >= startDate;
          });
          break;
        case 'last year':
          startDate = new Date(now.getFullYear() - 1, 0, 1);
          const endOfLastYear = new Date(now.getFullYear(), 0, 0);
          filteredData = filteredData.filter(row => {
            const rowDate = new Date(row[dateIndex]);
            return rowDate >= startDate && rowDate <= endOfLastYear;
          });
          break;
        case 'q1':
        case 'q2':
        case 'q3':
        case 'q4':
          const quarter = parseInt(period.charAt(1));
          const startMonth = (quarter - 1) * 3;
          startDate = new Date(now.getFullYear(), startMonth, 1);
          const endDate = new Date(now.getFullYear(), startMonth + 3, 0);
          filteredData = filteredData.filter(row => {
            const rowDate = new Date(row[dateIndex]);
            return rowDate >= startDate && rowDate <= endDate;
          });
          break;
      }
    }
    
    // Find amount column if it exists
    const amountIndex = headers.findIndex(h => 
      h === 'amount' || h.includes('amount') || 
      h === 'value' || h.includes('price') || 
      h === 'cost'
    );
    
    // Aggregate by category
    const categories = {};
    filteredData.forEach(row => {
      const category = row[categoryIndex] || 'Uncategorized';
      if (!categories[category]) {
        categories[category] = {
          count: 0,
          total: 0
        };
      }
      
      categories[category].count++;
      
      if (amountIndex !== -1) {
        const amount = parseFloat(row[amountIndex]);
        if (!isNaN(amount)) {
          categories[category].total += amount;
        }
      }
    });
    
    // Sort categories by count or amount
    let sortedCategories;
    if (amountIndex !== -1) {
      sortedCategories = Object.entries(categories).sort((a, b) => b[1].total - a[1].total);
    } else {
      sortedCategories = Object.entries(categories).sort((a, b) => b[1].count - a[1].count);
    }
    
    // Limit to top N categories
    const topCategoriesList = sortedCategories.slice(0, topCategories);
    
    // Calculate totals
    const totalTransactions = filteredData.length;
    let totalAmount = 0;
    if (amountIndex !== -1) {
      totalAmount = filteredData.reduce((sum, row) => {
        const amount = parseFloat(row[amountIndex]);
        return sum + (isNaN(amount) ? 0 : amount);
      }, 0);
    }
    
    // Create a detailed prompt for Gemini analysis
    let prompt = `Analyze the following category breakdown for financial data`;
    
    if (period !== 'all') {
      prompt += ` for ${period}`;
    }
    
    prompt += `:\n\n`;
    
    prompt += `Top ${topCategoriesList.length} Categories:\n`;
    topCategoriesList.forEach(([category, data], index) => {
      prompt += `${index + 1}. ${category}: ${data.count} transactions`;
      
      if (amountIndex !== -1) {
        prompt += `, $${data.total.toFixed(2)}`;
        const percentage = (data.total / totalAmount) * 100;
        prompt += ` (${percentage.toFixed(1)}% of total)`;
      }
      
      prompt += '\n';
    });
    
    prompt += `\nTotal Transactions: ${totalTransactions}`;
    if (amountIndex !== -1) {
      prompt += `\nTotal Amount: $${totalAmount.toFixed(2)}`;
    }
    
    prompt += `\n\nPlease provide:\n`;
    prompt += `1. An analysis of spending patterns by category\n`;
    prompt += `2. Insights about the concentration or distribution of transactions\n`;
    prompt += `3. Any recommendations based on this category breakdown\n`;
    prompt += `4. Brief summary of key findings\n\n`;
    prompt += `Keep your response concise and focused on actionable insights.`;
    
    // Generate the analysis
    return await generateResponse(prompt, []);
    
  } catch (error) {
    logError(`Error analyzing categories: ${error.message}`);
    return `Error analyzing categories: ${error.message}`;
  }
}

/**
 * Creates a documentation dashboard for all available functionality
 * @returns {string} URL to the created documentation
 */
function createDocumentation() {
  try {
    const docName = 'Gemini Financial AI Documentation';
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    
    // Add title
    body.appendParagraph(docName)
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Add version and date
    body.appendParagraph(`Version ${APP_VERSION} - Generated on ${new Date().toLocaleDateString()}`)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setItalic(true);
    
    // Add introduction
    body.appendParagraph('Introduction')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(
      'Gemini Financial AI is an advanced financial analysis tool that ' +
      'integrates with Google Sheets to provide AI-powered insights, ' +
      'anomaly detection, and reporting capabilities.'
    );
    
    // Add key features
    body.appendParagraph('Key Features')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    const features = [
      'AI-powered anomaly detection in financial data',
      'Automated financial reporting with customizable templates',
      'Interactive chat interface for natural language queries',
      'Transaction pattern analysis and visualization',
      'QuickBooks data integration',
      'Scheduled analysis and reporting',
      'Advanced data visualization'
    ];
    
    features.forEach(feature => {
      body.appendListItem(feature).setGlyphType(DocumentApp.GlyphType.BULLET);
    });
    
    // Add available functions
    body.appendParagraph('Available Functions')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    // Define functions to document
    const functions = [
      {
        name: 'Analyze Sheet',
        description: 'Analyzes the entire sheet for anomalies and potential issues.',
        usage: 'Select from the Gemini Financial AI menu or use the analyzeSheet() function.'
      },
      {
        name: 'Generate Reports',
        description: 'Creates comprehensive financial reports with customizable options.',
        usage: 'Use "Generate Standard Report" or "Generate Executive Summary" from the Reports submenu.'
      },
      {
        name: 'Chat Assistant',
        description: 'Natural language interface to analyze your financial data.',
        usage: 'Open from the menu and type queries like "Analyze my spending by category".'
      }
      // Add more functions here
    ];
    
    functions.forEach(func => {
      body.appendParagraph(func.name)
          .setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.appendParagraph('Description: ' + func.description);
      body.appendParagraph('Usage: ' + func.usage);
      body.appendParagraph(''); // Spacer
    });
    
    // Add dynamic TOC
    insertDynamicTableOfContents(doc);
    
    // Add report configuration section
    body.appendParagraph('Report Configuration')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(
      'Reports can be customized with various options including charts, AI analysis, ' +
      'and executive summaries. Configure reports through the Reports menu.'
    );
    
    // Add scheduled automation section
    body.appendParagraph('Scheduled Automation')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.appendParagraph(
      'Set up automatic anomaly detection and reporting on a schedule (daily, weekly, or monthly). ' +
      'Results can be automatically emailed to specified recipients.'
    );
    
    // Make the document accessible
    const file = DriveApp.getFileById(doc.getId());
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return doc.getUrl();
  } catch (error) {
    logError(`Error creating documentation: ${error.message}`);
    throw new Error(`Failed to create documentation: ${error.message}`);
  }
}

/**
 * Shows documentation of available features
 */
function showDocumentation() {
  try {
    // Generate or retrieve documentation
    const docUrl = createDocumentation();
    
    // Show a dialog with a link to the documentation
    const htmlContent = `
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h2 style="color: #4285f4;">Gemini Financial AI Documentation</h2>
        <p>A comprehensive documentation has been created with information about all available features and how to use them.</p>
        <p>Click the button below to open the documentation:</p>
        <div style="text-align: center; margin-top: 20px;">
          <a href="${docUrl}" target="_blank" style="
            background-color: #4285f4;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 4px;
            font-weight: bold;">
            Open Documentation
          </a>
        </div>
      </div>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(450)
      .setHeight(300);
      
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Documentation');
  } catch (error) {
    logError(`Error showing documentation: ${error.message}`);
    showError("Could not create documentation: " + error.message);
  }
}