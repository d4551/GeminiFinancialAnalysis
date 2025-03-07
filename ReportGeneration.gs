/**
 * Report Generation with AI-powered insights and flexible reporting options.
 */

// Report generator configuration
const REPORT_CONFIG = {
    defaultTemplates: {
        'monthly': {
      title: 'Monthly Financial Analysis',
      sections: ['summary', 'transactions', 'anomalies', 'trends', 'recommendations'],
      aiPrompt: 'Provide a monthly financial analysis focusing on spending patterns, unusual transactions, and recommendations for improvement.'
    },
    'anomaly': {
      title: 'Anomaly Detection Report',
      sections: ['summary', 'anomalies', 'patterns', 'recommendations'],
      aiPrompt: 'Analyze these financial anomalies, explain their patterns, and recommend corrective actions.'
    },
    'budget': {
      title: 'Budget Performance Report',
      sections: ['summary', 'categories', 'variances', 'forecast'],
      aiPrompt: 'Compare actual spending against budget, identify variances, and provide a forecast for the coming period.'
    },
    'executive': {
      title: 'Executive Financial Summary',
      sections: ['highlights', 'risks', 'opportunities', 'decisions'],
      aiPrompt: 'Create an executive summary of the financial data with key metrics, risks, opportunities, and recommended decisions.'
    }
  }
};

function generateReportId() {
    return 'REPORT-' + new Date().toISOString().replace(/[^0-9]/g, '');
}

/**
 * Retrieves custom report configuration, merging defaults with user overrides.
 */
function getReportConfig() {
  const defaultConfig = {
    includeAIInsights: true,
    includeTableOfContents: false,
    dateRange: { start: null, end: null }
  };
  // Merge with stored user config if available
  const userConfig = getConfig().reportConfig || {};
  return { ...defaultConfig, ...userConfig };
}

/**
 * Creates a report based on specified anomalies and options.
 * @param {Array<Object>} anomalies The anomalies to include in the report.
 * @param {Object} options Report configuration options.
 * @returns {Promise<string>} URL to the created document.
 */
async function createReport(anomalies, options) {
    const reportId = generateReportId();
    const docName = (options && options.docName) ? options.docName : 'Transaction Error Report';
    const doc = DocumentApp.create(`${docName} - ${reportId}`);
    const body = doc.getBody();

    try {
        // Start with basic structure
        if (options && options.includeTitle) {
            addReportTitle(body, options.reportTitle || docName);
        }
        addReportMetadata(body, reportId, options);
        
        // Optional introduction
        if (options && options.includeIntroduction) {
            const introText = options.introText || await generateIntroduction(anomalies, options);
            addIntroduction(body, introText);
        }
        
        // Add main content
        await addMainReportContent(body, anomalies, options);
        
        // Add footer with metadata
        addReportFooter(body, doc);
        
        // Set document properties for better organization
        doc.setDescription(`Financial analysis report generated on ${new Date().toLocaleString()}`);
        
        return doc.getUrl();
    } catch (error) {
        logError(`Error generating report: ${error.message}`);
        // Add error notification to document
        body.appendParagraph('An error occurred during report generation:')
            .setForegroundColor('#FF0000');
        body.appendParagraph(error.message)
            .setForegroundColor('#FF0000');
            
        return doc.getUrl();
    }
}

/**
 * Adds the main content to the report based on options and data.
 * @param {GoogleAppsScript.Document.Body} body The document body.
 * @param {Array<Object>} anomalies The anomalies to report on.
 * @param {Object} options Report configuration options.
 * @returns {Promise<void>}
 */
async function addMainReportContent(body, anomalies, options) {
    // Core data display
    addAnomaliesTable(body, anomalies, options);
    
    // Optional sections based on options
    if (options && options.includeSummary) {
        addSummary(body, anomalies, options);
    }
    
    if (options && options.includeChart) {
        addChart(body, anomalies);
    }
    
    if (options && options.includeNumericAnalysis) {
        addNumericAnalysis(body, anomalies, options);
    }
    
    if (options && options.includeDetailedAnalysis) {
        if (options.includeAIResults) {
            await addGeminiAnalysis(body, anomalies, options);
        } else {
            addStandardAnalysis(body, anomalies);
        }
    }
    
    // Add executive summary if requested
    if (options && options.includeExecutiveSummary) {
        await addExecutiveSummary(body, anomalies, options);
    }
    
    // Add recommendations if requested
    if (options && options.includeRecommendations) {
        await addRecommendations(body, anomalies, options);
    }
}

/**
 * Generates an introduction for the report using AI.
 * @param {Array<Object>} anomalies The anomalies to report on.
 * @param {Object} options Report configuration options.
 * @returns {Promise<string>} Introduction text.
 */
async function generateIntroduction(anomalies, options) {
    try {
        const prompt = `Generate a brief introduction paragraph for a financial report with the following characteristics:
        - Report title: ${options.reportTitle || 'Financial Analysis'}
        - Contains data on ${anomalies.length} financial anomalies or transactions
        - ${options.reportContext || 'Standard financial analysis'}
        
        The introduction should be professional, briefly explain the purpose of the report and what the reader will find inside.`;
        
        const response = await generateReportAnalysis(prompt);
        return response;
    } catch (error) {
        logError(`Error generating introduction: ${error.message}`);
        return 'This report provides a detailed analysis of anomalies in the transactions. Each anomaly is listed to help identify and correct errors in the data.';
    }
}

function addReportTitle(body, reportTitle) {
    body.appendParagraph(reportTitle)
        .setHeading(DocumentApp.ParagraphHeading.TITLE)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
}

function addReportMetadata(body, reportId, options) {
    const dateStr = options && options.locale
        ? Utilities.formatDate(new Date(), options.locale, 'yyyy-MM-dd HH:mm:ss')
        : new Date().toLocaleString();

    const metadataParagraph = body.appendParagraph(`Report ID: ${reportId}\nDate: ${dateStr}`);
    
    if (options && options.author) {
        metadataParagraph.appendText(`\nPrepared by: ${options.author}`);
    }
    
    if (options && options.department) {
        metadataParagraph.appendText(`\nDepartment: ${options.department}`);
    }
    
    metadataParagraph.setHeading(DocumentApp.ParagraphHeading.HEADING2);
}

function addIntroduction(body, introText) {
    body.appendParagraph('Introduction')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);
    body.appendParagraph(introText);
}

/**
 * Adds standard analysis section without AI assistance.
 * @param {GoogleAppsScript.Document.Body} body The document body.
 * @param {Array<Object>} anomalies The anomalies to analyze.
 */
function addStandardAnalysis(body, anomalies) {
    body.appendParagraph('Analysis')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);
    
    // Group anomalies by error type
    const errorGroups = {};
    anomalies.forEach(anomaly => {
        if (Array.isArray(anomaly.errors)) {
            anomaly.errors.forEach(error => {
                if (!errorGroups[error]) {
                    errorGroups[error] = [];
                }
                errorGroups[error].push(anomaly);
            });
        }
    });
    
    // Add analysis for each error group
    Object.entries(errorGroups).forEach(([errorType, items]) => {
        body.appendParagraph(`${errorType} (${items.length} occurrences)`)
            .setHeading(DocumentApp.ParagraphHeading.HEADING3)
            .setBold(true);
        
        // Add standard explanation based on error type
        let explanation = '';
        if (errorType.includes('missing')) {
            explanation = 'Missing data can lead to incomplete financial records and potentially incorrect calculations or decisions.';
        } else if (errorType.includes('Invalid')) {
            explanation = 'Invalid data entries may indicate input errors or misunderstandings of data requirements.';
        } else if (errorType.includes('duplicate')) {
            explanation = 'Duplicate entries can lead to double-counting and inflated financial metrics.';
        } else if (errorType.includes('outlier')) {
            explanation = 'Outliers may represent legitimate large transactions or potential errors in data entry.';
        } else {
            explanation = 'These anomalies should be reviewed for potential data quality issues.';
        }
        
        body.appendParagraph(explanation);
    });
}

function addAnomaliesTable(body, anomalies, options) {
    if (!Array.isArray(anomalies) || anomalies.length === 0) {
        body.appendParagraph('No anomalies found.').setBold(true);
        return;
    }

    body.appendParagraph('Anomalies Detected')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);

    const table = body.appendTable();
    
    // Enhanced headers with additional fields if they exist in anomalies
    const baseHeaders = ['Row', 'Amount', 'Date', 'Description', 'Category'];
    
    // Check if any anomalies have transaction_type or confidence score
    const hasTransactionType = anomalies.some(a => a.transaction_type && a.transaction_type !== 'N/A');
    const hasConfidence = anomalies.some(a => a.confidence !== undefined);
    
    let headers = [...baseHeaders];
    if (hasTransactionType) {
        headers.push('Type');
    }
    if (hasConfidence) {
        headers.push('Confidence');
    }
    headers.push('Errors');
    
    const headerRow = table.appendTableRow();
    headers.forEach(header => {
        headerRow.appendTableCell(header).setBold(true).setBackgroundColor('#cccccc');
    });

    anomalies.forEach(anomaly => {
        const row = table.appendTableRow();
        
        // Format amount
        const amountStr = (options && options.locale && typeof anomaly.amount === 'number')
            ? Utilities.formatString('%s', anomaly.amount.toLocaleString(options.locale))
            : (anomaly.amount !== undefined ? anomaly.amount.toString() : 'N/A');
        
        // Format date
        const dateStr = (anomaly.date instanceof Date)
            ? ((options && options.locale)
                ? Utilities.formatDate(anomaly.date, options.locale, 'yyyy-MM-dd')
                : anomaly.date.toLocaleString())
            : anomaly.date || 'N/A';
        
        // Add base cells
        row.appendTableCell(anomaly.row ? anomaly.row.toString() : 'N/A');
        row.appendTableCell(amountStr);
        row.appendTableCell(dateStr);
        row.appendTableCell(anomaly.description || 'N/A');
        row.appendTableCell(anomaly.category || 'N/A');
        
        // Add transaction_type if any anomalies have it
        if (hasTransactionType) {
            row.appendTableCell(anomaly.transaction_type || 'N/A');
        }
        
        // Add confidence if any anomalies have it
        if (hasConfidence) {
            const confidence = anomaly.confidence || 1.0;
            const confidenceCell = row.appendTableCell(`${(confidence * 100).toFixed(0)}%`);
            
            // Color code confidence
            if (confidence > 0.8) {
                confidenceCell.setBackgroundColor('#FFCDD2'); // Light red
            } else if (confidence > 0.5) {
                confidenceCell.setBackgroundColor('#FFE0B2'); // Light orange
            } else {
                confidenceCell.setBackgroundColor('#FFF9C4'); // Light yellow
            }
        }
        
        // Add errors
        row.appendTableCell(Array.isArray(anomaly.errors) ? anomaly.errors.join(', ') : 'N/A');
    });
    
    // Add footnote explanation of confidence scores if needed
    if (hasConfidence) {
        body.appendParagraph('\nConfidence Score Legend:')
            .setItalic(true)
            .setFontSize(9);
        
        const legendTable = body.appendTable();
        const legendRow = legendTable.appendTableRow();
        
        const highCell = legendRow.appendTableCell('High (>80%)');
        highCell.setBackgroundColor('#FFCDD2');
        
        const mediumCell = legendRow.appendTableCell('Medium (50-80%)');
        mediumCell.setBackgroundColor('#FFE0B2');
        
        const lowCell = legendRow.appendTableCell('Low (<50%)');
        lowCell.setBackgroundColor('#FFF9C4');
        
        legendTable.setWidth(400);
    }
}

function addSummary(body, anomalies, options) {
    body.appendParagraph('Summary')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);

    const totalErrors = anomalies.length;
    body.appendParagraph(`Total Anomalies: ${totalErrors}`);

    if (options.includeNumericAnalysis) {
        const numericStats = computeNumericStats(anomalies, 'amount');
        if (numericStats) {
            body.appendParagraph('Numeric Analysis of Amounts:');
            body.appendParagraph(`- Count (non-N/A): ${numericStats.count}`);
            body.appendParagraph(`- Sum: ${numericStats.sum}`);
            body.appendParagraph(`- Average: ${numericStats.average}`);
            body.appendParagraph(`- Min: ${numericStats.min}`);
            body.appendParagraph(`- Max: ${numericStats.max}`);
        }
    }

    const errorBreakdown = breakdownErrors(anomalies);
    if (Object.keys(errorBreakdown).length > 0) {
        body.appendParagraph('Error Breakdown:');
        Object.entries(errorBreakdown).forEach(([error, count]) => {
            body.appendParagraph(`- ${error}: ${count}`);
        });
    }

    if (options.includeCategoryBreakdown) {
        const categoryBreakdown = breakdownCategories(anomalies);
        if (Object.keys(categoryBreakdown).length > 0) {
            body.appendParagraph('Category Breakdown:');
            Object.entries(categoryBreakdown).forEach(([category, count]) => {
                body.appendParagraph(`- ${category}: ${count}`);
            });
        }
    }

    // Add confidence level breakdown if available
    if (anomalies.some(a => a.confidence !== undefined)) {
        const confidenceLevels = {
            high: anomalies.filter(a => (a.confidence || 1.0) > 0.8).length,
            medium: anomalies.filter(a => {
                const confidence = a.confidence || 1.0;
                return confidence <= 0.8 && confidence > 0.5;
            }).length,
            low: anomalies.filter(a => (a.confidence || 1.0) <= 0.5).length
        };

        body.appendParagraph('Confidence Level Breakdown:');
        body.appendParagraph(`- High Confidence: ${confidenceLevels.high} issues`);
        body.appendParagraph(`- Medium Confidence: ${confidenceLevels.medium} issues`);
        body.appendParagraph(`- Low Confidence: ${confidenceLevels.low} issues`);
    }
}

/**
 * Adds a dedicated numeric analysis section with visualizations.
 * @param {GoogleAppsScript.Document.Body} body The document body.
 * @param {Array<Object>} anomalies The anomalies to analyze.
 * @param {Object} options Report configuration options.
 */
function addNumericAnalysis(body, anomalies, options) {
    body.appendParagraph('Numeric Analysis')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);
    
    const numericStats = computeNumericStats(anomalies, 'amount');
    if (!numericStats) {
        body.appendParagraph('No numeric data available for analysis.');
        return;
    }
    
    // Add formatted numeric stats
    const locale = options.locale || 'en-US';
    const formatCurrency = (value) => {
        if (typeof value !== 'number') return value;
        return value.toLocaleString(locale, {
            style: 'currency',
            currency: options.currency || 'USD'
        });
    };
    
    const statsTable = body.appendTable();
    
    // Add header row
    const headerRow = statsTable.appendTableRow();
    ['Metric', 'Value'].forEach(header => {
        headerRow.appendTableCell(header).setBold(true).setBackgroundColor('#cccccc');
    });
    
    // Add data rows
    [
        ['Count', numericStats.count],
        ['Sum', formatCurrency(numericStats.sum)],
        ['Average', formatCurrency(numericStats.average)],
        ['Minimum', formatCurrency(numericStats.min)],
        ['Maximum', formatCurrency(numericStats.max)],
        ['Range', formatCurrency(numericStats.max - numericStats.min)]
    ].forEach(([metric, value]) => {
        const row = statsTable.appendTableRow();
        row.appendTableCell(metric);
        row.appendTableCell(value.toString());
    });
    
    // Add distribution information
    const amounts = anomalies
        .map(a => typeof a.amount === 'number' ? a.amount : null)
        .filter(a => a !== null);
    
    if (amounts.length > 0) {
        body.appendParagraph('Distribution Analysis')
            .setHeading(DocumentApp.ParagraphHeading.HEADING3)
            .setBold(true);
        
        // Calculate quartiles for box plot description
        const sortedAmounts = [...amounts].sort((a, b) => a - b);
        const { q1, q3 } = calculateQuartiles(sortedAmounts);
        const iqr = q3 - q1;
        
        body.appendParagraph(`Quartile Information:
- First Quartile (Q1): ${formatCurrency(q1)}
- Median: ${formatCurrency(sortedAmounts[Math.floor(sortedAmounts.length / 2)])}
- Third Quartile (Q3): ${formatCurrency(q3)}
- Interquartile Range (IQR): ${formatCurrency(iqr)}`);
    }
}

function addChart(body, anomalies) {
    body.appendParagraph('Error Frequency Chart')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    const chartBlob = createErrorFrequencyChart(anomalies);
    if (chartBlob) {
        body.appendImage(chartBlob);
    } else {
        body.appendParagraph('No chart available (insufficient data or no errors).');
    }
    
    // Add an amount distribution chart if we have numeric amounts
    const amountChartBlob = createAmountDistributionChart(anomalies);
    if (amountChartBlob) {
        body.appendParagraph('Amount Distribution')
            .setHeading(DocumentApp.ParagraphHeading.HEADING3);
        body.appendImage(amountChartBlob);
    }
}

/**
 * Creates a chart showing the distribution of anomaly amounts.
 * @param {Array<Object>} anomalies The anomalies to chart.
 * @returns {Blob|null} The chart as a blob, or null if chart couldn't be created.
 */
function createAmountDistributionChart(anomalies) {
    // Extract numeric amounts
    const amounts = anomalies
        .map(a => typeof a.amount === 'number' ? a.amount : null)
        .filter(a => a !== null);
    
    if (amounts.length < 3) {
        return null;  // Need at least a few values for a meaningful chart
    }
    
    try {
        // Create a histogram by grouping amounts into bins
        const min = Math.min(...amounts);
        const max = Math.max(...amounts);
        const range = max - min;
        const binCount = Math.min(10, Math.max(5, Math.ceil(Math.sqrt(amounts.length))));
        const binWidth = range / binCount;
        
        const bins = Array(binCount).fill(0);
        const binLabels = [];
        
        for (let i = 0; i < binCount; i++) {
            const binStart = min + (i * binWidth);
            const binEnd = binStart + binWidth;
            binLabels.push(`${binStart.toFixed(2)} - ${binEnd.toFixed(2)}`);
        }
        
        amounts.forEach(amount => {
            const binIndex = Math.min(binCount - 1, Math.floor((amount - min) / binWidth));
            bins[binIndex]++;
        });
        
        // Create chart
        const dataTable = Charts.newDataTable()
            .addColumn(Charts.ColumnType.STRING, 'Amount Range')
            .addColumn(Charts.ColumnType.NUMBER, 'Frequency');
        
        binLabels.forEach((label, index) => {
            dataTable.addRow([label, bins[index]]);
        });
        
        const chart = Charts.newColumnChart()
            .setDataTable(dataTable)
            .setTitle('Amount Distribution')
            .setXAxisTitle('Amount Range')
            .setYAxisTitle('Frequency')
            .setDimensions(600, 400)
            .build();
        
        return chart.getBlob();
    } catch (error) {
        logError(`Error creating amount distribution chart: ${error.message}`);
        return null;
    }
}

async function addGeminiAnalysis(body, anomalies, options) {
    try {
        body.appendParagraph('AI-Powered Detailed Analysis')
            .setHeading(DocumentApp.ParagraphHeading.HEADING2)
            .setBold(true);

        const analysisPrompt = createGeminiAnalysisPrompt(anomalies, options);
        
        // Use the Gemini API for report analysis
        const analysisResponse = await generateReportAnalysis(analysisPrompt);

        if (analysisResponse) {
            // Format the response for better readability in the document
            const formattedResponse = formatGeminiResponse(analysisResponse);
            body.appendParagraph(formattedResponse);
        } else {
            body.appendParagraph('AI analysis could not be generated.');
        }

    } catch (error) {
        logError("Error generating Gemini Analysis for Report: " + error.message);
        body.appendParagraph('Error generating AI-powered detailed analysis: ' + error.message);
    }
}

function formatGeminiResponse(response) {
    if (!response) return '';
    
    // Replace markdown-style headers with proper formatting
    let formatted = response
        .replace(/^# (.*$)/gm, '$1\n') // h1 headers
        .replace(/^## (.*$)/gm, '$1\n') // h2 headers
        .replace(/^### (.*$)/gm, '$1\n') // h3 headers
        .replace(/\*\*(.*?)\*\*/g, '$1') // bold text - we'll handle this with proper Doc styling
        .replace(/\*(.*?)\*/g, '$1'); // italic text - we'll handle with proper Doc styling
    
    return formatted;
}

function createGeminiAnalysisPrompt(anomalies, options) {
    const locale = options.locale || 'en-US';
    const formattedAnomalies = anomalies.map(anomaly => {
        const amountStr = (typeof anomaly.amount === 'number')
            ? anomaly.amount.toLocaleString(locale)
            : (anomaly.amount !== undefined ? anomaly.amount.toString() : 'N/A');

        const dateStr = (anomaly.date instanceof Date)
            ? Utilities.formatDate(anomaly.date, locale, 'yyyy-MM-dd')
            : anomaly.date || 'N/A';

        return `- Row: ${anomaly.row}, Amount: ${amountStr}, Date: ${dateStr}, Description: ${anomaly.description}, Category: ${anomaly.category}, Email: ${anomaly.email}, Errors: ${anomaly.errors.join(', ')}`;
    }).join('\n');

    // Create a more comprehensive prompt that leverages Gemini's financial analysis capabilities
    return `As a financial analysis expert with experience in fraud detection and accounting, analyze these financial anomalies:

${formattedAnomalies}

Provide a comprehensive analysis including:
1. A summary of the key patterns and issues discovered
2. Root cause analysis of why these anomalies are occurring
3. Potential financial impact of these issues
4. Risk assessment (categorize issues as high/medium/low risk)
5. Recommended actions to address each category of anomaly
6. Procedural changes that could prevent similar issues in the future

Format your analysis with clear headings for each section. Include specific examples from the data to illustrate your points.`;
}

function addReportFooter(body, doc) {
    const footer = doc.getFooter() || doc.addFooter();
    const reportUrl = doc.getUrl();
    
    // Create a more complete footer with multiple lines
    const footerText = footer.appendParagraph(`Report generated by Gemini Financial AI on ${new Date().toLocaleString()}`);
    footerText.setFontSize(8).setForegroundColor('#777777');
    
    const urlText = footer.appendParagraph(`Report URL: ${reportUrl}`);
    urlText.setFontSize(8).setForegroundColor('#777777');
    
    const disclaimerText = footer.appendParagraph('CONFIDENTIAL: This report is for internal use only and contains sensitive financial information.');
    disclaimerText.setFontSize(8).setForegroundColor('#777777').setItalic(true);
}

function createErrorFrequencyChart(anomalies) {
    if (!anomalies || anomalies.length === 0) {
        return null;
    }

    const errorCounts = {};
    anomalies.forEach(anomaly => {
        if (Array.isArray(anomaly.errors)) {
            anomaly.errors.forEach(error => {
                errorCounts[error] = (errorCounts[error] || 0) + 1;
            });
        }
    });

    if (Object.keys(errorCounts).length === 0) {
        return null;
    }

    const dataTable = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, 'Error Type')
        .addColumn(Charts.ColumnType.NUMBER, 'Frequency');

    for (const error in errorCounts) {
        dataTable.addRow([error, errorCounts[error]]);
    }

    const chart = Charts.newBarChart()
        .setDataTable(dataTable)
        .setTitle('Error Frequency')
        .setXAxisTitle('Error Type')
        .setYAxisTitle('Frequency')
        .setDimensions(600, 400)
        .build();

    return chart.getBlob();
}

function computeNumericStats(anomalies, field) {
    let sum = 0;
    let count = 0;
    let min = Number.POSITIVE_INFINITY;
    let max = Number.NEGATIVE_INFINITY;

    anomalies.forEach(a => {
        const val = a[field];
        if (typeof val === 'number' && !isNaN(val)) {
            sum += val;
            count++;
            min = Math.min(min, val);
            max = Math.max(max, val);
        }
    });

    if (count === 0) return null;

    return {
        count,
        sum,
        average: sum / count,
        min,
        max
    };
}

function breakdownErrors(anomalies) {
    const counts = {};
    anomalies.forEach(a => {
        if (Array.isArray(a.errors)) {
            a.errors.forEach(err => {
                if (err && err !== 'N/A') {
                    counts[err] = (counts[err] || 0) + 1;
                }
            });
        } else if (a.errors && a.errors !== 'N/A') {
            counts[a.errors] = (counts[a.errors] || 0) + 1;
        }
    });
    return counts;
}

function breakdownCategories(anomalies) {
    const counts = {};
    anomalies.forEach(a => {
        const category = a.category;
        if (category && category !== 'N/A') {
            counts[category] = (counts[category] || 0) + 1;
        }
    });
    return counts;
}

/**
 * Creates an enhanced report with additional AI-powered insights.
 * @param {Array<Object>} anomalies Anomalies to include in the report.
 * @param {Object} options Report configuration options.
 * @returns {Promise<string>} URL to the created document.
 */
async function createEnhancedReport(anomalies, options) {
    // First create the basic report structure
    const doc = DocumentApp.create(`Enhanced ${options.reportTitle || 'Financial Analysis'} Report`);
    const body = doc.getBody();
    
    // Add standard report sections
    addReportTitle(body, options.reportTitle || 'Enhanced Financial Analysis');
    addReportMetadata(body, generateReportId(), options);
    
    if (options.includeIntroduction) {
        addIntroduction(body, options.introText);
    }
    
    // Add executive summary from Gemini
    if (options.includeExecutiveSummary) {
        await addExecutiveSummary(body, anomalies, options);
    }
    
    // Add anomalies data
    addAnomaliesTable(body, anomalies, options);
    
    if (options.includeSummary) {
        addSummary(body, anomalies, options);
    }
    
    if (options.includeChart) {
        addChart(body, anomalies);
    }
    
    // Add AI-powered detailed analysis
    if (options.includeDetailedAnalysis && options.includeAIResults) {
        await addGeminiAnalysis(body, anomalies, options);
    }
    
    // Add recommendations from Gemini
    if (options.includeRecommendations) {
        await addRecommendations(body, anomalies, options);
    }
    
    addReportFooter(body, doc);
    
    return doc.getUrl();
}

async function addExecutiveSummary(body, anomalies, options) {
    body.appendParagraph('Executive Summary')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);
    
    try {
        const summaryPrompt = `Create a concise executive summary (3-4 sentences) of these financial anomalies:
        ${JSON.stringify(anomalies.slice(0, 5))}
        
        Focus on the most critical issues and their potential business impact.`;
        
        const summary = await generateReportAnalysis(summaryPrompt);
        body.appendParagraph(summary || 'Executive summary could not be generated.');
    } catch (error) {
        logError("Error generating executive summary: " + error.message);
        body.appendParagraph('Executive summary could not be generated due to an error.');
    }
}

async function addRecommendations(body, anomalies, options) {
    body.appendParagraph('Recommendations')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);
    
    try {
        const recommendationsPrompt = `Based on these financial anomalies:
        ${JSON.stringify(anomalies)}
        
        Provide 3-5 specific, actionable recommendations for addressing these issues.
        Format each recommendation as a bullet point with a brief explanation.`;
        
        const recommendations = await generateReportAnalysis(recommendationsPrompt);
        body.appendParagraph(recommendations || 'Recommendations could not be generated.');
    } catch (error) {
        logError("Error generating recommendations: " + error.message);
        body.appendParagraph('Recommendations could not be generated due to an error.');
    }
}

/**
 * Sends a report by email to the specified recipient
 * @param {string} recipient The email address to send to
 * @param {string} reportUrl The URL of the generated report
 * @param {Array<Object>} anomalies The anomalies included in the report
 * @param {string} customMessage Optional custom message to include in the email
 * @returns {Promise<boolean>} True if email was sent successfully
 */
async function sendReportEmail(recipient, reportUrl, anomalies, customMessage = '') {
    try {
        // Get basic information for the email
        const anomalyCount = anomalies.length;
        const reportDate = new Date().toLocaleDateString();
        const spreadsheetName = SpreadsheetApp.getActiveSpreadsheet().getName();
        
        // Group anomalies by type for summary
        const errorTypes = {};
        anomalies.forEach(anomaly => {
            if (Array.isArray(anomaly.errors)) {
                anomaly.errors.forEach(error => {
                    errorTypes[error] = (errorTypes[error] || 0) + 1;
                });
            } else if (typeof anomaly.errors === 'string') {
                errorTypes[anomaly.errors] = (errorTypes[anomaly.errors] || 0) + 1;
            }
        });
        
        // Get top 3 most common anomaly types
        const topErrors = Object.entries(errorTypes)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 3)
            .map(([type, count]) => `${type} (${count})`);
            
        // If AI is enabled, try to get a professionally formatted email body
        let emailBody = '';
        try {
            const emailPrompt = `Write a professional email summarizing a financial anomaly report with these details:
            - ${anomalyCount} anomalies were detected in "${spreadsheetName}"
            - Most common issues: ${topErrors.join(', ')}
            - Report date: ${reportDate}
            
            Include a professional greeting, brief introduction explaining what the report contains,
            summary of the anomalies, and professional closing. Keep it concise and business-appropriate.
            Don't use markdown formatting, just plain text.`;
            
            emailBody = await generateReportAnalysis(emailPrompt);
        } catch (aiError) {
            logError(`Error generating email body with AI: ${aiError.message}`);
            // Fall back to standard message if AI fails
            emailBody = `Dear Colleague,

This email contains the financial anomaly report you requested for "${spreadsheetName}".

Summary:
- ${anomalyCount} potential anomalies were detected
- Main issue types: ${topErrors.join(', ')}

Please review the full report at your earliest convenience.

Kind regards,
Financial Analysis System`;
        }
        
        // Add custom message if provided
        if (customMessage && customMessage.trim()) {
            emailBody += `\n\nAdditional message from sender:\n${customMessage}\n\n`;
        }
        
        // Always add the report URL at the end for clarity
        emailBody += `\nView the complete report here: ${reportUrl}\n`;
        
        // Send the email
        const subject = `Financial Anomaly Report - ${spreadsheetName} - ${reportDate}`;
        GmailApp.sendEmail(
            recipient,
            subject,
            emailBody,
            {
                name: 'Financial Analysis System',
                replyTo: Session.getActiveUser().getEmail()
            }
        );
        
        // Log successful sending
        logMessage(`Report email sent successfully to ${recipient}`);
        return true;
    } catch (error) {
        logError(`Error sending report email: ${error.message}`);
        return false;
    }
}

/**
 * Generates a financial report with optional AI highlights and table of contents.
 */
async function generateFinancialReport(reportType, includeAI = true, dateRange = null) {
  const config = getReportConfig();
  // Overwrite from function args
  config.includeAIInsights = includeAI;
  if (dateRange) {
    config.dateRange = dateRange;
  }

  const doc = DocumentApp.create(`Financial Report - ${reportType}`); 

  // ...existing code to gather data, create doc, etc...

  // Insert table of contents if requested
  if (config.includeTableOfContents) {
    insertTableOfContents(doc);
  }

  // If AI insights are enabled, add a summary passage
  if (config.includeAIInsights) {
    const aiSummary = await generateReportAnalysis(`Briefly summarize financial report (${reportType}).`);
    doc.getBody().appendParagraph(aiSummary).setHeading(DocumentApp.ParagraphHeading.HEADING3);
  }

  // ...existing code...
  return doc.getUrl();
}

/**
 * Inserts a table of contents at the beginning of the document.
 * @param {GoogleAppsScript.Document.Document} doc
 */
function insertTableOfContents(doc) {
  const body = doc.getBody();
  body.insertParagraph(0, 'Table of Contents').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.insertParagraph(1, '')
      .asParagraph().addPositionedImage(body.getImages()[0] || null); // Example: or dynamic ToC generation
}

/**
 * Safely appends pattern analysis text to the given document.
 * Throws an error if the document URL is invalid.
 */
function appendPatternAnalysisToReport(docUrl, patternAnalysis) {
  if (!validateAndNormalizeDocUrl(docUrl)) {
    throw new Error("Invalid document URL format");
  }
  
  return appendToReport(docUrl, patternAnalysis);
}

/**
 * Generates a pattern analysis report, appending the analysis text to the document.
 */
function generatePatternAnalysisReport(docUrl, analysisType) {
  try {
    // ...existing code...
    const analysisText = `Pattern analysis results for type: ${analysisType}`;
    appendPatternAnalysisToReport(docUrl, analysisText);
    // ...existing code...
  } catch (error) {
    throw new Error('Failed to generate pattern analysis: ' + error.message);
  }
}

// ...existing code...

/**
 * Generates a report from a predefined template
 * @param {string} reportType The type of report to generate
 * @param {boolean} includeAI Whether to include AI-powered insights
 * @returns {Promise<string>} Report URL or status message
 */
async function generateReportFromTemplate(reportType, includeAI = true) {
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
        currency: getDefaultCurrency()
    };
    
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

    // Filter out empty anomalies
    const validAnomalies = anomalies.filter(a => !isEmptyAnomaly(a));

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
        
        // Return a formatted response with HTML link
        return `Report generated successfully. <a href="${reportUrl}" target="_blank">Open Report</a>`;
    } else {
        return 'No valid anomalies found to generate a report.';
    }
}

/**
 * Checks if an anomaly object is empty or just has placeholder values
 * @param {Object} anomaly The anomaly to check
 * @returns {boolean} True if it's an empty anomaly
 */
function isEmptyAnomaly(anomaly) {
    return Object.values(anomaly).every(val => 
        !val || 
        val.toString().trim().toUpperCase() === 'N/A' || 
        (typeof val === 'number' && val === 0)
    );
}

/**
 * Extracts anomalies from the error report sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} errorSheet The error sheet
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
 * Generates and emails a report
 * @param {string} recipient Email recipient
 * @param {string} reportType Report type to generate
 * @param {string} customMessage Optional custom message
 * @returns {Promise<string>} Status message
 */
async function emailReport(recipient, reportType, customMessage) {
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
        const success = await ReportUtils.sendReportEmail(recipient, reportUrl, anomalies, customMessage);
        
        if (success) {
            return `Report sent successfully to ${recipient}`;
        } else {
            return "Error sending email. Please check the logs.";
        }
    } catch (error) {
        logError(`Error in emailReport: ${error.message}`);
        return `Error: ${error.message}`;
    }
}

/**
 * Generates a pattern analysis report
 * @param {string} analysisType The type of pattern analysis
 * @param {boolean} includeVisuals Whether to include visualizations
 * @param {boolean} includeAI Whether to include AI insights
 * @returns {Promise<string>} HTML link to the generated report
 */
async function generatePatternReport(analysisType, includeVisuals, includeAI) {
  try {
    const sheetData = getSheetData();
    const anomalies = await detectAnomalies(SpreadsheetApp.getActiveSheet());
    
    const options = {
      reportType: 'pattern',
      reportTitle: `${analysisType.charAt(0).toUpperCase() + analysisType.slice(1)} Pattern Analysis`,
      includeTitle: true,
      includeIntroduction: true,
      includeSummary: true,
      includeChart: includeVisuals,
      includeDetailedAnalysis: true,
      includeAIResults: includeAI,
      analysisType: analysisType,
      locale: getDefaultLocale()
    };

    // Generate the report with focus on patterns
    let reportUrl;
    try {
      reportUrl = await createReport(anomalies, options);
      
      // Validate report URL immediately
      if (!reportUrl || typeof reportUrl !== 'string') {
        throw new Error('Invalid report URL returned from createReport');
      }
      
      // Log for debugging
      logMessage(`Pattern report generated at URL: ${reportUrl}`);
    } catch (reportError) {
      logError(`Error creating pattern report: ${reportError.message}`);
      throw new Error(`Failed to create pattern report: ${reportError.message}`);
    }

    // If AI insights were requested, enhance the report with pattern-specific analysis
    if (includeAI && reportUrl) {
      try {
        const patternPrompt = `Analyze the financial data for ${analysisType} patterns: 
        ${analysisType === 'frequency' ? 'Focus on transaction frequency and recurring patterns.' :
          analysisType === 'temporal' ? 'Focus on time-based patterns and seasonal trends.' :
          analysisType === 'value' ? 'Focus on transaction value distributions and outliers.' :
          'Focus on category-based analysis and spending patterns.'}`;

        const patternInsights = await generateResponse(patternPrompt, sheetData);
        
        // First validate URL
        const validatedUrl = validateAndNormalizeDocUrl(reportUrl);
        
        if (!validatedUrl) {
          logError(`Invalid document URL format: ${reportUrl}`);
          // Continue without appending - we'll still return the report URL
        } else {
          // Try to append, but don't fail if this part doesn't work
          const appendSuccess = await appendToReport(validatedUrl, patternInsights);
          if (!appendSuccess) {
            logError(`Failed to append AI insights to report: ${validatedUrl}`);
          }
        }
      } catch (aiError) {
        // Log but don't fail the whole operation
        logError(`Error generating AI insights: ${aiError.message}`);
      }
    }

    return `Pattern analysis report generated successfully. <a href="${reportUrl}" target="_blank">Open Pattern Analysis Report</a>`;
  } catch (error) {
    logError(`Error generating pattern analysis report: ${error.message}`);
    throw new Error(`Failed to generate pattern analysis: ${error.message}`);
  }
}

// ...existing code...

/**
 * Core routines for generating advanced financial reports.
 * These functions serve as the central hub for all reporting functionality.
 */

/**
 * Generates a financial report of the specified type
 * @param {string} reportType The type of report to generate
 * @param {boolean} includeAI Whether to include AI-powered insights
 * @returns {Promise<string>} Confirmation message with report URL
 */
async function generateFinancialReport(reportType, includeAI = true) {
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
    currency: getDefaultCurrency()
  };
  
  return await generateReportWithOptions(options);
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
    const validAnomalies = anomalies.filter(a => !isEmptyAnomaly(a));

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
 * Generates an executive summary report
 * @returns {Promise<string>} Confirmation message with report URL
 */
async function generateExecutiveSummary() {
  try {
    // Use the financial report function with executive type and AI insights
    return await generateFinancialReport('executive', true);
  } catch (error) {
    logError(`Error generating executive summary: ${error.message}`);
    return `Failed to generate executive summary: ${error.message}`;
  }
}

/**
 * Emails a report to the specified recipient
 * @param {string} recipient Email recipient
 * @param {string} reportType Report type to generate
 * @param {string} customMessage Optional custom message
 * @returns {Promise<string>} Status message
 */
async function emailReport(recipient, reportType, customMessage) {
  try {
    if (!recipient) {
      throw new Error("Recipient email is required");
    }

    // Validate email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(recipient)) {
      throw new Error("Invalid email format");
    }
    
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
    const success = await ReportUtils.sendReportEmail(recipient, reportUrl, anomalies, customMessage);
    
    if (success) {
      return `Report sent successfully to ${recipient}`;
    } else {
      return "Error sending email. Please check the logs.";
    }
  } catch (error) {
    logError(`Error in emailReport: ${error.message}`);
    return `Error: ${error.message}`;
  }
}

/**
 * Generates a pattern analysis report
 * @param {string} analysisType The type of pattern analysis
 * @param {boolean} includeVisuals Whether to include visualizations
 * @param {boolean} includeAI Whether to include AI insights
 * @returns {Promise<string>} HTML link to the generated report
 */
async function generatePatternReport(analysisType, includeVisuals, includeAI) {
  try {
    const sheetData = getSheetData();
    const anomalies = await detectAnomalies(SpreadsheetApp.getActiveSheet());
    
    const options = {
      reportType: 'pattern',
      reportTitle: `${analysisType.charAt(0).toUpperCase() + analysisType.slice(1)} Pattern Analysis`,
      includeTitle: true,
      includeIntroduction: true,
      includeSummary: true,
      includeChart: includeVisuals,
      includeDetailedAnalysis: true,
      includeAIResults: includeAI,
      analysisType: analysisType,
      locale: getDefaultLocale()
    };

    // Generate the report with focus on patterns
    let reportUrl;
    try {
      reportUrl = await createReport(anomalies, options);
      
      // Validate report URL immediately
      if (!reportUrl || typeof reportUrl !== 'string') {
        throw new Error('Invalid report URL returned from createReport');
      }
      
      // Log for debugging
      logMessage(`Pattern report generated at URL: ${reportUrl}`);
    } catch (reportError) {
      logError(`Error creating pattern report: ${reportError.message}`);
      throw new Error(`Failed to create pattern report: ${reportError.message}`);
    }

    // If AI insights were requested, enhance the report with pattern-specific analysis
    if (includeAI && reportUrl) {
      try {
        const patternPrompt = `Analyze the financial data for ${analysisType} patterns: 
        ${analysisType === 'frequency' ? 'Focus on transaction frequency and recurring patterns.' :
          analysisType === 'temporal' ? 'Focus on time-based patterns and seasonal trends.' :
          analysisType === 'value' ? 'Focus on transaction value distributions and outliers.' :
          'Focus on category-based analysis and spending patterns.'}`;

        const patternInsights = await generateResponse(patternPrompt, sheetData);
        
        // First validate URL
        const validatedUrl = validateAndNormalizeDocUrl(reportUrl);
        
        if (!validatedUrl) {
          logError(`Invalid document URL format: ${reportUrl}`);
          // Continue without appending - we'll still return the report URL
        } else {
          // Try to append, but don't fail if this part doesn't work
          const appendSuccess = await appendToReport(validatedUrl, patternInsights);
          if (!appendSuccess) {
            logError(`Failed to append AI insights to report: ${validatedUrl}`);
          }
        }
      } catch (aiError) {
        // Log but don't fail the whole operation
        logError(`Error generating AI insights: ${aiError.message}`);
      }
    }

    return `Pattern analysis report generated successfully. <a href="${reportUrl}" target="_blank">Open Pattern Analysis Report</a>`;
  } catch (error) {
    logError(`Error generating pattern analysis report: ${error.message}`);
    throw new Error(`Failed to generate pattern analysis: ${error.message}`);
  }
}

/**
 * Generate a report from a template and returns HTML with link
 * @param {string} templateName The name of the template to use
 * @param {boolean} includeAI Whether to include AI-generated insights
 * @returns {Promise<string>} The HTML link to the report
 */
async function generateReportFromTemplate(templateName, includeAI = true) {
  try {
    // Check if template exists
    const template = REPORT_CONFIG.defaultTemplates[templateName.toLowerCase()];
    
    if (!template) {
      return `Error: Template "${templateName}" not found`;
    }

    const result = await generateFinancialReport(templateName, includeAI);
    return result;
  } catch (error) {
    logError(`Error generating report from template ${templateName}: ${error.message}`);
    return `Failed to generate report: ${error.message}`;
  }
}

/**
 * Creates and sends a scheduled report based on saved settings
 * This function is intended to be called by a time-based trigger
 */
async function generateScheduledReport() {
  try {
    // Get report settings from script properties
    const scriptProps = PropertiesService.getScriptProperties();
    const reportType = scriptProps.getProperty('SCHEDULED_REPORT_TYPE') || 'standard';
    const recipientEmail = scriptProps.getProperty('SCHEDULED_REPORT_EMAIL');
    const includeAI = scriptProps.getProperty('SCHEDULED_REPORT_INCLUDE_AI') !== 'false';
    
    if (!recipientEmail) {
      logError('No recipient email configured for scheduled reports');
      return;
    }
    
    // Generate the report
    const result = await generateFinancialReport(reportType, includeAI);
    
    // Extract the URL from the result HTML
    const urlMatch = result.match(/href="([^"]+)"/);
    const reportUrl = urlMatch ? urlMatch[1] : null;
    
    if (!reportUrl) {
      throw new Error('Could not extract report URL from result');
    }
    
    // Send email with the report URL
    const subject = `Scheduled ${reportType.charAt(0).toUpperCase() + reportType.slice(1)} Financial Report - ${new Date().toLocaleDateString()}`;
    const message = `Your scheduled financial report is ready.

View the report here: ${reportUrl}

This is an automated message from Gemini Financial AI.`;
    
    GmailApp.sendEmail(recipientEmail, subject, message);
    
    logMessage(`Scheduled report (${reportType}) sent to ${recipientEmail}`);
    return `Report sent to ${recipientEmail}`;
  } catch (error) {
    logError(`Error in scheduled report generation: ${error.message}`);
    throw error;
  }
}

// ...existing code...