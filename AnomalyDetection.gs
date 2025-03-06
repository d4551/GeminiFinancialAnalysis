/**
 * Detects anomalies in a given sheet, leveraging both traditional checks and Gemini AI.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Google Sheet to analyze.
 * @returns {Anomaly[]} An array of anomaly objects.
 */
async function detectAnomalies(sheet) {
    if (!sheet) {
        throw new Error('Sheet is undefined or not provided.');
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length < 2) {
        return [];  // Not enough data.
    }

    const config = getConfig();
    const headers = values[0].map(h => (h || '').toString().toLowerCase().trim());
    const data = values.slice(1);
    const indexes = headers.reduce((acc, header, idx) => { acc[header] = idx; return acc; }, {});

    let anomalies = [];
     //Traditional Checks
    if(config.detectionAlgorithm === 'standard' || config.detectionAlgorithm === 'hybrid'){
      anomalies = performTraditionalChecks(data, indexes, config);
    }

    // AI-Powered Checks with Gemini
    if (config.enableAIDetection && (config.detectionAlgorithm === 'ai' || config.detectionAlgorithm === 'hybrid')) {
        try {
            const aiAnomalies = await analyzeSheetWithGemini(values); // Use Gemini Service
            if (aiAnomalies && aiAnomalies.anomalies && Array.isArray(aiAnomalies.anomalies)) {
              //Adjust row numbers from AI if it did not account for headers
              aiAnomalies.anomalies.forEach(anomaly => {
                if(anomaly.row){
                  anomaly.row = parseInt(anomaly.row) + 1;
                }
              });

               // Merge traditional and AI-detected anomalies, handling duplicates
                anomalies = mergeAnomalies(anomalies, aiAnomalies.anomalies);
            }
        } catch (geminiError) { // Catch errors from Gemini Service
            logError("Error during AI analysis: " + geminiError.message);
            showError("AI Anomaly Detection Failed: " + geminiError.message + ". See logs for details."); // Display user-friendly error
        }
    }

    return anomalies;
}

/**
 * Performs traditional anomaly checks (data validation, duplicates, outliers).
 * @param {Array<Array<any>>} data The sheet data (without headers).
 * @param {{[header: string]: number}} indexes A map of header names to column indexes.
 * @param {object} config
 * @returns {Anomaly[]} An array of anomaly objects.
 */
function performTraditionalChecks(data, indexes, config) {
    const anomalies = [];
    const seenEntries = new Set();
    const amounts = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const errors = [];

        const amountVal = indexes['amount'] !== undefined ? row[indexes['amount']] : null;
        const dateVal = indexes['date'] !== undefined ? row[indexes['date']] : '';
        const descVal = indexes['description'] !== undefined ? row[indexes['description']] : '';
        const catVal = indexes['category'] !== undefined ? row[indexes['category']] : '';
        const emailVal = indexes['email'] !== undefined ? row[indexes['email']] : null;

        let amount, date, description, category, email;

        try {
            amount = amountVal !== null && amountVal !== '' ? parseFloat(amountVal) : NaN;
        } catch (error) {
            logError("Could not Parse amount:" + error);
            amount = NaN
        }

        try {
            date = dateVal ? parseDate(dateVal, config) : null;
        } catch (error) {
            logError("Could not Parse date:" + error)
            date = null;
        }

        description = descVal.toString().trim();
        category = catVal.toString().trim();
        email = emailVal ? emailVal.toString().trim() : '';


        // 1. Mandatory Fields
        config.mandatoryFields.forEach(field => {
            const fieldIndex = indexes[field.toLowerCase()];
            if (fieldIndex !== undefined) {
                const fieldValue = row[fieldIndex];
                if (!fieldValue || fieldValue.toString().trim() === '') {
                    errors.push(`${field} is missing`);
                }
            }
        });

       // 2. Amount
        if (indexes['amount'] !== undefined) {
            if (isNaN(amount)) {
                errors.push('Amount is not a number');
            } else {
                if (!config.amount.allowNegative && amount < 0) {
                    errors.push('Negative amount is not allowed');
                }
                if (amount < config.amount.min || amount > config.amount.max) {
                    errors.push(`Amount out of range (${config.amount.min}-${config.amount.max})`);
                }
                amounts.push(amount);
            }
        }

        // 3. Date
        if (indexes['date'] !== undefined && dateVal.toString().trim() !== "") {
          if (!date) {
              errors.push('Invalid date format');
          } else if (!config.date.allowFuture && date > new Date()) {
              errors.push('Future dates are not allowed');
          }
        }
        else {
            if (indexes['date'] !== undefined && config.mandatoryFields.includes('date')){
                errors.push('Date is missing');
            }
        }

        // 4. Description
        if(indexes['description'] !== undefined){
          if (config.description.required && description === '') {
              errors.push('Description is empty');
          }
        }

        // 5. Category
        if (indexes['category'] !== undefined) {
          if (config.category.required && category !== '' && !config.category.validCategories.includes(category)) {
              errors.push(`Invalid category: ${category}`);
          }
        }

        // 6. Email
        if(indexes['email'] !== undefined){
          if (config.email.required && email !== '') {
            if (!config.email.format.test(email)) {
                errors.push('Invalid email format');
            }
          }
        }

      // 7. Duplicate Check
        if (config.duplicates.check && config.duplicates.uniqueColumns && config.duplicates.uniqueColumns.length > 0) {
          const uniqueKey = config.duplicates.uniqueColumns
              .map(col => {
                  const idx = indexes[col.toLowerCase()];
                  return idx !== undefined ? (row[idx] || '').toString() : '';
              })
              .join('|');

          if (seenEntries.has(uniqueKey)) {
              errors.push('Duplicate entry detected');
          } else {
              seenEntries.add(uniqueKey);
          }
        }

        if (errors.length > 0) {
            anomalies.push({
                row: i + 2,  // +2 for header row and 1-based indexing.
                errors,
                amount: isNaN(amount) ? "N/A" : amount,
                date: date ? date.toISOString().split('T')[0] : "N/A",
                description: description || "N/A",
                category: category || "N/A",
                email: email || "N/A"
            });
        }
    }

    // --- Outlier Detection ---
    if (config.outliers.check && config.outliers.method.toLowerCase() !== 'none' && amounts.length > 0) {
        const outlierAnomalies = detectOutliers(amounts, data, indexes, config);
        anomalies.push(...outlierAnomalies);
    }

    return anomalies;
}

/**
 * Merges traditional and AI-detected anomalies, avoiding duplicates.
 * @param {Anomaly[]} traditionalAnomalies
 * @param {Anomaly[]} aiAnomalies
 * @returns {Anomaly[]} The merged array of anomalies.
 */
function mergeAnomalies(traditionalAnomalies, aiAnomalies) {
    const merged = [...traditionalAnomalies]; // Start with traditional
    const traditionalRows = new Set(traditionalAnomalies.map(a => a.row)); // Keep track of rows

    for (const aiAnomaly of aiAnomalies) {
        if (!traditionalRows.has(aiAnomaly.row)) {
            merged.push(aiAnomaly); // Add if not a duplicate (based on row number)
        }
    }
    return merged;
}

function parseDate(dateStr, config) {
    const trimmedDateStr = dateStr.trim();

    let parsedDate = new Date(trimmedDateStr);
    if (!isNaN(parsedDate.getTime())) {
        return parsedDate;
    }

    for (const pattern of config.date.datePatterns) {
        if (pattern.test(trimmedDateStr)) {
            let year, month, day;

            if (/^(\d{4})-(\d{2})-(\d{2})$/.test(trimmedDateStr)) {
                [year, month, day] = trimmedDateStr.split('-').map(Number);
            } else if (/^(\d{2})\/(\d{2})\/(\d{4})$/.test(trimmedDateStr)) {
                [month, day, year] = trimmedDateStr.split('/').map(Number);
            } else {
                continue;
            }

            parsedDate = new Date(year, month - 1, day);
            if (!isNaN(parsedDate.getTime())) {
                return parsedDate;
            }
        }
    }

    return null;
}

function detectOutliers(amounts, data, indexes, config) {
    const method = config.outliers.method.toLowerCase();
    const anomalies = [];

    switch (method) {
        case 'zscore':
            const threshold = config.outliers.threshold;
            const { mean, stddev } = calculateMeanAndStdDev(amounts);

            amounts.forEach((amt, idx) => {
                if (stddev > 0 && Math.abs(amt - mean) > threshold * stddev) {
                    const anomaly = createOutlierAnomaly(idx, data, indexes, amt, 'Z-score');
                    anomalies.push(anomaly);
                }
            });
            break;

        case 'iqr':
            const iqrFactor = config.outliers.iqrFactor;
            const { q1, q3 } = calculateQuartiles(amounts);
            const iqr = q3 - q1;
            const lowerBound = q1 - iqrFactor * iqr;
            const upperBound = q3 + iqrFactor * iqr;

            amounts.forEach((amt, idx) => {
                if (amt < lowerBound || amt > upperBound) {
                    const anomaly = createOutlierAnomaly(idx, data, indexes, amt, 'IQR');
                    anomalies.push(anomaly);
                }
            });
            break;

        default:
            logMessage(`Outlier detection method "${method}" not recognized or set to none.`);
            break;
    }

    return anomalies;
}

function calculateMeanAndStdDev(numbers) {
    const n = numbers.length;
    if (n === 0) return { mean: 0, stddev: 0 };

    const mean = numbers.reduce((sum, val) => sum + val, 0) / n;
    const variance = numbers.reduce((sum, val) => sum + Math.pow(val - mean, 2), 0) / n;
    const stddev = Math.sqrt(variance);
    return { mean, stddev };
}

function calculateQuartiles(sortedNumbers) {
    const n = sortedNumbers.length;
    if (n === 0) return { q1: 0, q3: 0 };

    const q1Index = Math.floor((n + 1) / 4) - 1;
    const q3Index = Math.floor((3 * (n + 1)) / 4) - 1;

    const q1 = sortedNumbers[q1Index];
    const q3 = sortedNumbers[q3Index];
    return { q1, q3 };
}

function createOutlierAnomaly(idx, data, indexes, amt, method) {
    const rowNumber = idx + 2;
    return {
        row: rowNumber,
        errors: [`Amount is an outlier (${method} method)`],
        amount: amt,
        date: indexes['date'] !== undefined && data[idx][indexes['date']] ? data[idx][indexes['date']] : "N/A",
        description: indexes['description'] !== undefined ? data[idx][indexes['description']] : "N/A",
        category: indexes['category'] !== undefined ? data[idx][indexes['category']] : "N/A",
        email: indexes['email'] !== undefined ? data[idx][indexes['email']] : "N/A"
    };
}

function highlightAnomalies(sheet, anomalies) {
    anomalies.forEach(anomaly => {
        const row = anomaly.row;
        const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
        range.setBackground('red');
        range.setNote(anomaly.errors.join(', '));
    });
}

function createErrorReportSheet(anomalies) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'Error Report';
    let errorSheet = spreadsheet.getSheetByName(sheetName);

    if (errorSheet) {
        spreadsheet.deleteSheet(errorSheet);
    }

    errorSheet = spreadsheet.insertSheet(sheetName);

    const headers = ['Row', 'Amount', 'Date', 'Description', 'Category', 'Email', 'Errors'];
    errorSheet.appendRow(headers).setFrozenRows(1);
    errorSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#4caf50')
        .setFontColor('white');

    anomalies.forEach(anomaly => {
        errorSheet.appendRow([
            anomaly.row,
            anomaly.amount,
            anomaly.date,
            anomaly.description,
            anomaly.category,
            anomaly.email,
            anomaly.errors.join(', ')
        ]);
    });

    errorSheet.appendRow(['Total Anomalies', anomalies.length]).setFontWeight('bold');
    errorSheet.autoResizeColumns(1, headers.length);
}