function generateReportId() {
    return 'REPORT-' + new Date().toISOString().replace(/[^0-9]/g, '');
}

async function createReport(anomalies, options) {
    const reportId = generateReportId();
    const docName = (options && options.docName) ? options.docName : 'Transaction Error Report';
    const doc = DocumentApp.create(`${docName} - ${reportId}`);
    const body = doc.getBody();

    if (options && options.includeTitle) {
        addReportTitle(body, options.reportTitle || docName);
    }
    addReportMetadata(body, reportId, options);
    if (options && options.includeIntroduction) {
        addIntroduction(body, options.introText);
    }
    addAnomaliesTable(body, anomalies, options);
    if (options && options.includeSummary) {
        addSummary(body, anomalies, options);
    }
    if (options && options.includeChart) {
        addChart(body, anomalies);
    }
     if (options && options.includeDetailedAnalysis && options.includeAIResults) { // Conditionally add AI Analysis section
        await addGeminiAnalysis(body, anomalies, options);
    }
    addReportFooter(body, doc);

    return doc.getUrl();
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

    body.appendParagraph(`Report ID: ${reportId}\nDate: ${dateStr}`)
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
}

function addIntroduction(body, customIntroText) {
    body.appendParagraph('Introduction')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setBold(true);
    const introText = customIntroText || (
        'This report provides a detailed analysis of anomalies in the transactions. ' +
        'Each anomaly is listed to help identify and correct errors in the data. ' +
        'Charts and summaries are included to give an overview of the issues and their impact.'
    );
    body.appendParagraph(introText);
}

function addAnomaliesTable(body, anomalies, options) {
    if (!Array.isArray(anomalies) || anomalies.length === 0) {
        body.appendParagraph('No anomalies found.').setBold(true);
        return;
    }

    const table = body.appendTable();
    const headers = ['Row', 'Amount', 'Date', 'Description', 'Category', 'Email', 'Errors'];
    const headerRow = table.appendTableRow();
    headers.forEach(header => {
        headerRow.appendTableCell(header).setBold(true).setBackgroundColor('#cccccc');
    });

    anomalies.forEach(anomaly => {
        const row = table.appendTableRow();
        const amountStr = (options && options.locale && typeof anomaly.amount === 'number')
            ? Utilities.formatString('%s', anomaly.amount.toLocaleString(options.locale))
            : (anomaly.amount !== undefined ? anomaly.amount.toString() : 'N/A');

        const dateStr = (anomaly.date instanceof Date)
            ? ((options && options.locale)
                ? Utilities.formatDate(anomaly.date, options.locale, 'yyyy-MM-dd')
                : anomaly.date.toLocaleString())
            : anomaly.date || 'N/A';

        row.appendTableCell(anomaly.row ? anomaly.row.toString() : 'N/A');
        row.appendTableCell(amountStr);
        row.appendTableCell(dateStr);
        row.appendTableCell(anomaly.description || 'N/A');
        row.appendTableCell(anomaly.category || 'N/A');
        row.appendTableCell(anomaly.email || 'N/A');
        row.appendTableCell(Array.isArray(anomaly.errors) ? anomaly.errors.join(', ') : 'N/A');
    });
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
}

async function addGeminiAnalysis(body, anomalies, options) {
    try {
        body.appendParagraph('AI Powered Detailed Analysis')
            .setHeading(DocumentApp.ParagraphHeading.HEADING2)
            .setBold(true);

        const analysisPrompt = createGeminiAnalysisPrompt(anomalies, options); // Create specialized prompt
        const analysisResponse = await generateReportAnalysis(analysisPrompt); // Using generateReportAnalysis from GeminiService

        if (analysisResponse) {
            body.appendParagraph(analysisResponse);
        } else {
            body.appendParagraph('AI analysis could not be generated.');
        }

    } catch (error) {
        logError("Error generating Gemini Analysis for Report: " + error.message);
        body.appendParagraph('Error generating AI-powered detailed analysis.');
    }
}

function createGeminiAnalysisPrompt(anomalies, options) {
     const locale = options.locale || 'en-US';
     const formattedAnomalies = anomalies.map(anomaly => {
        const amountStr = (options && options.locale && typeof anomaly.amount === 'number')
            ? Utilities.formatString('%s', anomaly.amount.toLocaleString(locale))
            : (anomaly.amount !== undefined ? anomaly.amount.toString() : 'N/A');

        const dateStr = (anomaly.date instanceof Date)
                ? ((options && options.locale)
                    ? Utilities.formatDate(anomaly.date, options.locale, 'yyyy-MM-dd')
                    : anomaly.date.toLocaleString())
                : anomaly.date || 'N/A';

        return `- Row: ${anomaly.row}, Amount: ${amountStr}, Date: ${dateStr}, Description: ${anomaly.description}, Category: ${anomaly.category}, Email: ${anomaly.email}, Errors: ${anomaly.errors.join(', ')}`;
    }).join('\n');


    return `Generate a detailed and insightful analysis of the following financial anomalies to be included in a transaction error report. Focus on providing actionable insights and potential root causes for these anomalies. Consider the context of each anomaly (amount, date, description, category, email, and errors) to infer underlying issues and suggest next steps for investigation or correction.  The analysis should be concise, professional, and directly address the patterns or individual anomalies listed below. Format the analysis as structured paragraphs suitable for inclusion in a business report.\n\nAnomalies:\n${formattedAnomalies}\n\nAnalysis:`;
}


function addReportFooter(body, doc) {
    const footer = doc.getFooter() || doc.addFooter();
    const reportUrl = doc.getUrl();
    footer.appendParagraph(`Report generated on ${new Date().toLocaleString()}.  URL: ${reportUrl}`)
        .setFontSize(8)
        .setForegroundColor('#777777');
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