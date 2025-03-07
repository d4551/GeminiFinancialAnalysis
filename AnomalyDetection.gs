/**
 * Anomaly detection with transaction type support and advanced pattern recognition.
 */

/**
 * Transaction type definitions with specific validation rules
 */
const TRANSACTION_TYPES = {
  'PAYMENT': {
    requiredFields: ['amount', 'date', 'description'],
    validations: {
      amount: { mustBePositive: true },
      description: { minLength: 3 }
    }
  },
  'EXPENSE': {
    requiredFields: ['amount', 'date', 'category', 'description'],
    validations: {
      amount: { mustBeNegative: true },
      category: { mustBeInList: ['Office', 'Travel', 'Utilities', 'Salaries', 'Marketing', 'Other'] }
    }
  },
  'TRANSFER': {
    requiredFields: ['amount', 'date', 'from_account', 'to_account'],
    validations: {
      amount: { mustBePositive: true }
    }
  },
  'REFUND': {
    requiredFields: ['amount', 'date', 'original_transaction_id'],
    validations: {
      amount: { mustBePositive: true }
    }
  }
};

/**
 * Retrieves the user configuration or falls back to defaults.
 */
function getConfig() {
  try {
    // Attempt to load from script properties
    const scriptProps = PropertiesService.getScriptProperties();
    const enableAI = scriptProps.getProperty('ENABLE_AI');
    const config = getDefaultConfig(); 
    if (enableAI !== null) {
      config.enableAIDetection = (enableAI === 'true');
    }
    // Add or override more values if needed
    return config;
  } catch (error) {
    // Fallback to defaults on error
    return getDefaultConfig();
  }
}

/**
 * Detects anomalies in a given sheet, leveraging both traditional checks, 
 * transaction type rules, and Gemini AI.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The Google Sheet to analyze.
 * @param {Object} [configOverrides] Optional config overrides.
 * @returns {Promise<Anomaly[]>} A promise that resolves to an array of anomaly objects.
 */
async function detectAnomalies(sheet, configOverrides = {}) {
    try {
        if (!sheet) {
            throw new Error('Sheet is undefined or not provided.');
        }
        
        const startTime = new Date().getTime();
        logMessage(`Starting anomaly detection on sheet: ${sheet.getName()}`);

        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();

        if (values.length < 2) {
            logMessage('Not enough data for anomaly detection (less than 2 rows)');
            return [];  // Not enough data.
        }

        // Merge default config with any overrides
        const config = { ...getConfig(), ...configOverrides };
        
        // Extract headers and data
        const headers = values[0].map(h => (h || '').toString().toLowerCase().trim());
        const data = values.slice(1);
        const indexes = headers.reduce((acc, header, idx) => { acc[header] = idx; return acc; }, {});

        // Initialize empty anomaly array
        let anomalies = [];
        
        // Traditional Checks
        if (config.detectionAlgorithm === 'standard' || config.detectionAlgorithm === 'hybrid') {
            anomalies = performTraditionalChecks(data, headers, indexes, config);
            logMessage(`Traditional checks found ${anomalies.length} anomalies`);
        }

        // Type-specific validations
        if (indexes['transaction_type'] !== undefined) {
            const typeAnomalies = performTypeSpecificValidations(data, headers, indexes);
            logMessage(`Transaction type validation found ${typeAnomalies.length} anomalies`);
            anomalies = mergeAnomalies(anomalies, typeAnomalies);
        }

        // AI-Powered Checks with enhanced retry logic
        if (config.enableAIDetection && (config.detectionAlgorithm === 'ai' || config.detectionAlgorithm === 'hybrid')) {
            try {
                // Use our enhanced AI function with retry mechanism
                const aiResult = await analyzeWithAI(values, 'Gemini', 3);
                
                if (aiResult && aiResult.anomalies && Array.isArray(aiResult.anomalies)) {
                    // Adjust row numbers from AI if it did not account for headers
                    aiResult.anomalies.forEach(anomaly => {
                        if (anomaly.row) {
                            anomaly.row = parseInt(anomaly.row) + 1;
                        }
                    });
                    
                    logMessage(`AI detection found ${aiResult.anomalies.length} anomalies`);
                    // Merge traditional and AI-detected anomalies
                    anomalies = mergeAnomalies(anomalies, aiResult.anomalies);
                }
            } catch (geminiError) {
                logError(`AI analysis failed: ${geminiError.message}`, 'detectAnomalies.aiCheck');
                // Continue with traditional anomalies only
            }
        }
        
        // Perform pattern-based anomaly detection
        const patternAnomalies = detectPatternAnomalies(data, headers, indexes, config);
        logMessage(`Pattern detection found ${patternAnomalies.length} anomalies`);
        anomalies = mergeAnomalies(anomalies, patternAnomalies);

        // Apply confidence scores
        const scoredAnomalies = applyConfidenceScores(anomalies);
        
        try {
            // Validate the final anomalies array
            validateAnomalies(scoredAnomalies);
        } catch (validationError) {
            logError(`Anomaly validation error: ${validationError.message}`, 'detectAnomalies.validation');
            // Try to clean up invalid anomalies rather than failing completely
            scoredAnomalies = cleanInvalidAnomalies(scoredAnomalies);
        }
        
        const endTime = new Date().getTime();
        logMessage(`Anomaly detection completed in ${(endTime - startTime)/1000} seconds. Found ${scoredAnomalies.length} total anomalies.`);
        
        return scoredAnomalies;
    } catch (error) {
        logError(`Critical error in detectAnomalies: ${error.message}`, 'detectAnomalies');
        notifyUserOfError("Anomaly detection failed: " + error.message);
        throw error;
    }
}

/**
 * Performs transaction type-specific validations based on the transaction type column.
 * @param {Array<Array<any>>} data The sheet data (without headers).
 * @param {Array<string>} headers The header names.
 * @param {{[header: string]: number}} indexes A map of header names to column indexes.
 * @returns {Anomaly[]} An array of anomaly objects.
 */
function performTypeSpecificValidations(data, headers, indexes) {
    const anomalies = [];
    const typeIndex = indexes['transaction_type'];
    
    if (typeIndex === undefined) return anomalies;
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const errors = [];
        const transactionType = row[typeIndex]?.toString().trim().toUpperCase();
        
        if (!transactionType) {
            errors.push('Transaction type is missing');
            anomalies.push(createAnomalyObject(i + 2, errors, row, headers, indexes));
            continue;
        }
        
        // Skip if transaction type is not recognized
        const typeRules = TRANSACTION_TYPES[transactionType];
        if (!typeRules) {
            errors.push(`Unknown transaction type: ${transactionType}`);
            anomalies.push(createAnomalyObject(i + 2, errors, row, headers, indexes));
            continue;
        }
        
        // Check required fields for this transaction type
        typeRules.requiredFields.forEach(field => {
            const fieldIndex = indexes[field.toLowerCase()];
            if (fieldIndex !== undefined) {
                const fieldValue = row[fieldIndex];
                if (!fieldValue || fieldValue.toString().trim() === '') {
                    errors.push(`${field} is required for ${transactionType} transactions`);
                }
            }
        });
        
        // Apply transaction type-specific validations
        if (typeRules.validations) {
            // Amount validations
            if (typeRules.validations.amount) {
                const amountIndex = indexes['amount'];
                if (amountIndex !== undefined) {
                    const amount = parseFloat(row[amountIndex]);
                    
                    if (!isNaN(amount)) {
                        if (typeRules.validations.amount.mustBePositive && amount <= 0) {
                            errors.push(`${transactionType} transactions must have a positive amount`);
                        }
                        if (typeRules.validations.amount.mustBeNegative && amount >= 0) {
                            errors.push(`${transactionType} transactions must have a negative amount`);
                        }
                    }
                }
            }
            
            // Category validations
            if (typeRules.validations.category && typeRules.validations.category.mustBeInList) {
                const categoryIndex = indexes['category'];
                if (categoryIndex !== undefined) {
                    const category = row[categoryIndex]?.toString().trim();
                    const validList = typeRules.validations.category.mustBeInList;
                    
                    if (category && !validList.includes(category)) {
                        errors.push(`Invalid category for ${transactionType}: ${category}. Must be one of: ${validList.join(', ')}`);
                    }
                }
            }
            
            // Description validations
            if (typeRules.validations.description && typeRules.validations.description.minLength) {
                const descIndex = indexes['description'];
                if (descIndex !== undefined) {
                    const description = row[descIndex]?.toString().trim();
                    const minLength = typeRules.validations.description.minLength;
                    
                    if (description && description.length < minLength) {
                        errors.push(`Description for ${transactionType} must be at least ${minLength} characters`);
                    }
                }
            }
        }
        
        if (errors.length > 0) {
            anomalies.push(createAnomalyObject(i + 2, errors, row, headers, indexes));
        }
    }
    
    return anomalies;
}

/**
 * Detects pattern-based anomalies such as round numbers, duplicate values, and suspicious frequencies.
 * @param {Array<Array<any>>} data The sheet data (without headers).
 * @param {Array<string>} headers The header names.
 * @param {{[header: string]: number}} indexes A map of header names to column indexes.
 * @param {object} config Configuration object.
 * @returns {Anomaly[]} An array of anomaly objects.
 */
function detectPatternAnomalies(data, headers, indexes, config) {
    const anomalies = [];
    
    // Check for patterns only if we have enough data
    if (data.length < 5) return anomalies;
    
    // Get all amount values for frequency analysis
    const amountIndex = indexes['amount'];
    const amountValues = [];
    const amountFrequency = {};
    const roundNumberThreshold = config.roundNumberThreshold || 100;
    
    if (amountIndex !== undefined) {
        // Build frequency map
        data.forEach((row, rowIndex) => {
            const amount = parseFloat(row[amountIndex]);
            if (!isNaN(amount)) {
                amountValues.push({ value: amount, rowIndex });
                const key = amount.toString();
                amountFrequency[key] = (amountFrequency[key] || 0) + 1;
            }
        });
        
        // Check for round numbers (potentially suspicious)
        amountValues.forEach(({ value, rowIndex }) => {
            // Check if amount is a round number (divisible by threshold)
            if (Math.abs(value) >= roundNumberThreshold && value % roundNumberThreshold === 0) {
                const errors = [`Suspicious round amount (${value})`];
                anomalies.push(createAnomalyObject(rowIndex + 2, errors, data[rowIndex], headers, indexes));
            }
        });
        
        // Check for suspiciously frequent amounts
        const frequencyThreshold = Math.max(3, Math.ceil(data.length * 0.1)); // At least 3 occurrences or 10% of data
        Object.entries(amountFrequency).forEach(([amount, frequency]) => {
            if (frequency >= frequencyThreshold) {
                // Find all rows with this amount
                const matchingRows = amountValues
                    .filter(item => item.value.toString() === amount)
                    .map(item => item.rowIndex);
                    
                // Add anomaly for each matching row
                matchingRows.forEach(rowIndex => {
                    const errors = [`Suspiciously frequent amount: ${amount} (appears ${frequency} times)`];
                    anomalies.push(createAnomalyObject(rowIndex + 2, errors, data[rowIndex], headers, indexes, 0.7));
                });
            }
        });
    }
    
    // Check for suspicious timing patterns
    const dateIndex = indexes['date'];
    if (dateIndex !== undefined) {
        const dateMap = new Map();
        const weekendTransactions = [];
        
        // Process all dates
        data.forEach((row, rowIndex) => {
            try {
                const dateValue = row[dateIndex];
                let date;
                
                if (dateValue instanceof Date) {
                    date = dateValue;
                } else if (typeof dateValue === 'string') {
                    date = new Date(dateValue);
                }
                
                if (date && !isNaN(date.getTime())) {
                    // Check for weekend transactions (potentially suspicious)
                    const dayOfWeek = date.getDay();
                    if (dayOfWeek === 0 || dayOfWeek === 6) { // Sunday or Saturday
                        weekendTransactions.push(rowIndex);
                    }
                    
                    // Map by date for duplicate detection
                    const dateStr = date.toISOString().split('T')[0];
                    if (!dateMap.has(dateStr)) {
                        dateMap.set(dateStr, []);
                    }
                    dateMap.get(dateStr).push(rowIndex);
                }
            } catch (e) {
                // Skip if date parsing fails
            }
        });
        
        // Flag weekend transactions if configured to do so
        if (config.flagWeekendTransactions) {
            weekendTransactions.forEach(rowIndex => {
                const errors = ['Transaction occurred on a weekend'];
                anomalies.push(createAnomalyObject(rowIndex + 2, errors, data[rowIndex], headers, indexes, 0.5));
            });
        }
        
        // Check for multiple transactions on the same day with same category (if category exists)
        const categoryIndex = indexes['category'];
        if (categoryIndex !== undefined) {
            dateMap.forEach((rowIndexes, dateStr) => {
                if (rowIndexes.length > 1) {
                    // Group by category
                    const categoryCounts = {};
                    rowIndexes.forEach(idx => {
                        const category = data[idx][categoryIndex];
                        if (category) {
                            if (!categoryCounts[category]) {
                                categoryCounts[category] = [];
                            }
                            categoryCounts[category].push(idx);
                        }
                    });
                    
                    // Check for categories with multiple transactions on the same day
                    Object.entries(categoryCounts).forEach(([category, indexes]) => {
                        if (indexes.length > 1) {
                            indexes.forEach(idx => {
                                const errors = [`Multiple ${category} transactions on ${dateStr}`];
                                anomalies.push(createAnomalyObject(idx + 2, errors, data[idx], headers, headers, 0.6));
                            });
                        }
                    });
                }
            });
        }
    }
    
    return anomalies;
}

/**
 * Creates an anomaly object with extracted data from the row.
 * @param {number} rowNumber The 1-based row number.
 * @param {string[]} errors Array of error messages.
 * @param {any[]} row The data row.
 * @param {string[]} headers The header names.
 * @param {{[header: string]: number}} indexes A map of header names to column indexes.
 * @param {number} [confidence=1.0] Confidence score for the anomaly (0-1).
 * @returns {Anomaly} The anomaly object.
 */
function createAnomalyObject(rowNumber, errors, row, headers, indexes, confidence = 1.0) {
    // Extract standard fields
    const standardFields = ['amount', 'date', 'description', 'category', 'email', 'transaction_type'];
    const anomaly = {
        row: rowNumber,
        errors: errors,
        confidence: confidence
    };
    
    // Add all available standard fields
    standardFields.forEach(field => {
        const idx = indexes[field];
        if (idx !== undefined) {
            let value = row[idx];
            
            // Format dates
            if (field === 'date' && value instanceof Date && !isNaN(value.getTime())) {
                value = value.toISOString().split('T')[0];
            }
            
            anomaly[field] = value !== undefined && value !== null ? value : "N/A";
        }
    });
    
    return anomaly;
}

/**
 * Apply confidence scores to anomalies based on their characteristics.
 * @param {Anomaly[]} anomalies Array of anomalies.
 * @returns {Anomaly[]} Array of anomalies with confidence scores.
 */
function applyConfidenceScores(anomalies) {
    return anomalies.map(anomaly => {
        // Start with default confidence if not already set
        let confidence = anomaly.confidence || 1.0;
        
        // Adjust confidence based on number of errors
        if (anomaly.errors && anomaly.errors.length > 1) {
            // More errors = higher confidence it's actually an anomaly
            confidence = Math.min(1.0, confidence + (anomaly.errors.length - 1) * 0.1);
        }
        
        // Adjust confidence based on error types
        if (anomaly.errors) {
            anomaly.errors.forEach(error => {
                // Certain error types have higher confidence
                if (error.includes('missing') || error.includes('required')) {
                    confidence = Math.min(1.0, confidence + 0.1);
                }
                if (error.includes('Invalid') || error.includes('not allowed')) {
                    confidence = Math.min(1.0, confidence + 0.2);
                }
                // Pattern-based anomalies often have lower confidence
                if (error.includes('Suspicious') || error.includes('frequent')) {
                    confidence = Math.max(0.3, confidence - 0.1);
                }
            });
        }
        
        return { ...anomaly, confidence };
    });
}

/**
 * Performs traditional anomaly checks (data validation, duplicates, outliers).
 * @param {Array<Array<any>>} data The sheet data (without headers).
 * @param {Array<string>} headers The header names.
 * @param {{[header: string]: number}} indexes A map of header names to column indexes.
 * @param {object} config
 * @returns {Anomaly[]} An array of anomaly objects.
 */
function performTraditionalChecks(data, headers, indexes, config) {
    const anomalies = [];
    const seenEntries = new Set();
    const amounts = [];

    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const errors = [];

        // Extract values from the row
        const amountVal = indexes['amount'] !== undefined ? row[indexes['amount']] : null;
        const dateVal = indexes['date'] !== undefined ? row[indexes['date']] : '';
        const descVal = indexes['description'] !== undefined ? row[indexes['description']] : '';
        const catVal = indexes['category'] !== undefined ? row[indexes['category']] : '';
        const emailVal = indexes['email'] !== undefined ? row[indexes['email']] : null;

        let amount, date, description, category, email;

        // Parse amount
        try {
            amount = amountVal !== null && amountVal !== '' ? parseFloat(amountVal) : NaN;
        } catch (error) {
            logError("Could not Parse amount:" + error);
            amount = NaN;
        }

        // Parse date
        try {
            date = dateVal ? parseDate(dateVal, config) : null;
        } catch (error) {
            logError("Could not Parse date:" + error);
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
        if (indexes['description'] !== undefined) {
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
        if (indexes['email'] !== undefined) {
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

            if (uniqueKey && seenEntries.has(uniqueKey)) {
                errors.push('Duplicate entry detected');
            } else {
                seenEntries.add(uniqueKey);
            }
        }

        if (errors.length > 0) {
            anomalies.push(createAnomalyObject(i + 2, errors, row, headers, indexes));
        }
    }

    // --- Outlier Detection ---
    if (config.outliers.check && config.outliers.method.toLowerCase() !== 'none' && amounts.length > 0) {
        const outlierAnomalies = detectOutliers(amounts, data, headers, indexes, config);
        anomalies.push(...outlierAnomalies);
    }

    return anomalies;
}

/**
 * Detects outliers in numeric data.
 * @param {number[]} amounts Array of amount values.
 * @param {Array<Array<any>>} data The sheet data.
 * @param {Array<string>} headers The header names.
 * @param {{[header: string]: number}} indexes A map of header names to column indexes.
 * @param {object} config Configuration object.
 * @returns {Anomaly[]} Array of anomalies for outliers.
 */
function detectOutliers(amounts, data, headers, indexes, config) {
    const method = config.outliers.method.toLowerCase();
    const anomalies = [];

    switch (method) {
        case 'zscore':
            const threshold = config.outliers.threshold;
            const { mean, stddev } = calculateMeanAndStdDev(amounts);

            amounts.forEach((amt, idx) => {
                if (stddev > 0 && Math.abs(amt - mean) > threshold * stddev) {
                    const errors = [`Amount (${amt}) is an outlier (Z-score method)`];
                    anomalies.push(createAnomalyObject(idx + 2, errors, data[idx], headers, indexes, 0.8));
                }
            });
            break;

        case 'iqr':
            const iqrFactor = config.outliers.iqrFactor;
            // Sort for quartile calculation
            const sortedAmounts = [...amounts].sort((a, b) => a - b);
            const { q1, q3 } = calculateQuartiles(sortedAmounts);
            const iqr = q3 - q1;
            const lowerBound = q1 - iqrFactor * iqr;
            const upperBound = q3 + iqrFactor * iqr;

            amounts.forEach((amt, idx) => {
                if (amt < lowerBound || amt > upperBound) {
                    const errors = [`Amount (${amt}) is an outlier (IQR method)`];
                    anomalies.push(createAnomalyObject(idx + 2, errors, data[idx], headers, indexes, 0.8));
                }
            });
            break;

        default:
            logMessage(`Outlier detection method "${method}" not recognized or set to none.`);
            break;
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

/**
 * Uses Gemini AI to analyze a dataset for financial anomalies.
 * This is a wrapper around the analyzeSheetWithGemini function from Gemini.gs.
 * 
 * @param {Array<Array<any>>} sheetData The spreadsheet data including headers.
 * @returns {Promise<Object>} Promise that resolves to anomaly detection results.
 */
async function analyzeWithGeminiAI(sheetData) {
    try {
        return await analyzeSheetWithGemini(sheetData);
    } catch (error) {
        logError("Gemini AI Analysis Error: " + error.message);
        throw new Error("Failed to analyze data with Gemini AI: " + error.message);
    }
}

/**
 * Runs focused analysis on a specific subset of data.
 * 
 * @param {Array<Array<any>>} data The data to analyze.
 * @param {string} analysisType The type of analysis to perform.
 * @returns {Promise<string>} The analysis results.
 */
async function runFocusedAnalysis(data, analysisType) {
    try {
        const prompt = `Analyze this financial data with a focus on ${analysisType}. 
        Identify patterns, anomalies, and provide insights specifically related to ${analysisType}.
        Format your response as a clear, concise analysis with bullet points for key findings.
        
        Data: ${JSON.stringify(data.slice(0, 100))}`;
        
        return await generateReportAnalysis(prompt);
    } catch (error) {
        logError("Focused Analysis Error: " + error.message);
        return "Error performing focused analysis: " + error.message;
    }
}

/**
 * Highlights anomalies in the sheet and adds notes from Gemini AI.
 * Using a color scale based on confidence scores.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to highlight.
 * @param {Anomaly[]} anomalies Array of detected anomalies.
 */
async function highlightAndAnnotateAnomalies(sheet, anomalies) {
    if (!anomalies || anomalies.length === 0) return;
    
    try {
        // Group anomalies by row for efficiency
        const rowMap = {};
        anomalies.forEach(anomaly => {
            if (!rowMap[anomaly.row]) rowMap[anomaly.row] = [];
            rowMap[anomaly.row].push(anomaly);
        });
        
        // Process each row with anomalies
        for (const [rowNum, rowAnomalies] of Object.entries(rowMap)) {
            const row = parseInt(rowNum);
            const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
            
            // Calculate maximum confidence score for this row
            const maxConfidence = Math.max(...rowAnomalies.map(a => a.confidence || 1.0));
            
            // Select color based on confidence score
            // High confidence (red) -> medium confidence (orange) -> low confidence (yellow)
            let backgroundColor;
            if (maxConfidence > 0.8) {
                backgroundColor = '#FFCDD2'; // Lighter red
            } else if (maxConfidence > 0.5) {
                backgroundColor = '#FFE0B2'; // Lighter orange
            } else {
                backgroundColor = '#FFF9C4'; // Lighter yellow
            }
            
            // Apply highlighting
            range.setBackground(backgroundColor);
            
            // Generate explanation using Gemini
            const anomalyContext = rowAnomalies.map(a => 
                `${a.errors.join(', ')} (Amount: ${a.amount}, Description: ${a.description}${
                    a.transaction_type ? `, Type: ${a.transaction_type}` : ''
                })`
            ).join('; ');
            
            const prompt = `Briefly explain this financial anomaly in 1-2 sentences: ${anomalyContext}`;
            let explanation = "Anomalies detected: " + rowAnomalies[0].errors.join(', ');
            
            // Try to get AI explanation if enabled
            const config = getConfig();
            if (config.enableAIDetection && config.includeAIExplanations) {
                try {
                    const aiExplanation = await generateReportAnalysis(prompt);
                    if (aiExplanation && aiExplanation.trim()) {
                        explanation = aiExplanation;
                    }
                } catch (aiError) {
                    logError("Error getting AI explanation: " + aiError.message);
                }
            }
            
            // Add note with explanation and confidence score
            range.setNote(`${explanation}\nConfidence: ${(maxConfidence * 100).toFixed(0)}%`);
        }
    } catch (error) {
        logError("Error highlighting anomalies: " + error.message);
    }
}

/**
 * Enhanced AI integration with retry logic for more reliable API calls.
 * @param {Array<Array<any>>} sheetData The spreadsheet data with headers.
 * @param {string} provider The AI provider to use (default: Gemini)
 * @param {number} maxRetries Maximum number of retry attempts
 * @returns {Promise<Object>} Promise that resolves to anomaly detection results.
 */
async function analyzeWithAI(sheetData, provider = 'Gemini', maxRetries = 3) {
    let lastError = null;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            logMessage(`AI analysis attempt ${attempt}/${maxRetries} using ${provider}`);
            
            switch (provider.toLowerCase()) {
                case 'gemini':
                    return await analyzeSheetWithGemini(sheetData);
                case 'custom':
                    // Reserved for future custom AI implementation
                    throw new Error('Custom AI provider not yet implemented');
                default:
                    throw new Error(`Unsupported AI provider: ${provider}`);
            }
        } catch (error) {
            lastError = error;
            logError(`AI analysis attempt ${attempt} failed: ${error.message}`, 'AnomalyDetection.analyzeWithAI');
            
            // Don't wait on the last attempt
            if (attempt < maxRetries) {
                // Exponential backoff with jitter for retry
                const backoffMs = Math.min(1000 * Math.pow(2, attempt) + Math.random() * 1000, 10000);
                logMessage(`Retrying in ${Math.round(backoffMs/1000)} seconds...`);
                await new Promise(resolve => Utilities.sleep(backoffMs));
            }
        }
    }
    
    // If we get here, all retries failed
    logError(`All ${maxRetries} AI analysis attempts failed`, 'AnomalyDetection.analyzeWithAI');
    throw lastError || new Error('AI analysis failed after multiple attempts');
}

/**
 * Enhanced anomaly validation with stronger type checking
 * @param {Array<Object>} anomalies Array of anomalies to validate
 * @returns {boolean} True if valid, throws error if invalid
 */
function validateAnomalies(anomalies) {
    if (!Array.isArray(anomalies)) {
        throw new Error('Anomalies must be an array');
    }
    
    anomalies.forEach((anomaly, index) => {
        if (!anomaly || typeof anomaly !== 'object') {
            throw new Error(`Anomaly at index ${index} is not a valid object`);
        }
        
        // Validate row number is present and valid
        if (anomaly.row === undefined || anomaly.row === null) {
            throw new Error(`Anomaly at index ${index} is missing required 'row' property`);
        }
        
        // Check that errors are present
        if (!anomaly.errors || (!Array.isArray(anomaly.errors) && typeof anomaly.errors !== 'string')) {
            throw new Error(`Anomaly at index ${index} has invalid 'errors' property`);
        }
        
        // Check for amount validity if present
        if (anomaly.amount !== undefined && anomaly.amount !== null && anomaly.amount !== 'N/A') {
            const amountNum = parseFloat(anomaly.amount);
            if (isNaN(amountNum)) {
                throw new Error(`Anomaly at index ${index} has invalid amount format`);
            }
        }
    });
    
    return true;
}

/**
 * Attempts to clean and repair invalid anomaly objects
 * @param {Array<Object>} anomalies Array of potentially invalid anomalies
 * @returns {Array<Object>} Cleaned anomalies array
 */
function cleanInvalidAnomalies(anomalies) {
    if (!Array.isArray(anomalies)) return [];
    
    return anomalies.filter(anomaly => {
        try {
            // Must have at least row and some form of errors
            if (!anomaly || typeof anomaly !== 'object') return false;
            if (anomaly.row === undefined || anomaly.row === null) return false;
            
            // Ensure errors exists in some form
            if (!anomaly.errors) {
                anomaly.errors = ["Unknown error"];
            }
            
            // Convert errors to array if it's a string
            if (typeof anomaly.errors === 'string') {
                anomaly.errors = [anomaly.errors];
            }
            
            // If errors is still not valid, create a default
            if (!Array.isArray(anomaly.errors)) {
                anomaly.errors = ["Error format issue"];
            }
            
            // Ensure confidence is a number between 0-1
            if (anomaly.confidence === undefined || anomaly.confidence === null || 
                isNaN(parseFloat(anomaly.confidence))) {
                anomaly.confidence = 0.5; // Default medium confidence
            } else {
                anomaly.confidence = Math.max(0, Math.min(1, parseFloat(anomaly.confidence)));
            }
            
            return true;
        } catch (e) {
            return false;
        }
    });
}

/**
 * Notifies the user of a critical error in anomaly detection
 * @param {string} errorMessage The error message
 */
function notifyUserOfError(errorMessage) {
    try {
        const ui = SpreadsheetApp.getUi();
        ui.alert('Anomaly Detection Error', errorMessage, ui.ButtonSet.OK);
    } catch (e) {
        // If we can't use UI, log to console
        console.error('ALERT: ' + errorMessage);
    }
}

/**
 * Enhanced error logging with context and optional sheet logging
 * @param {string} errorMessage The error message
 * @param {string} context Additional context about where the error occurred
 * @param {boolean} logToSheet Whether to log to a sheet (if configured)
 */
function logError(errorMessage, context = '', logToSheet = true) {
    const timestamp = new Date().toISOString();
    const fullMessage = `${timestamp} | ERROR | ${context} | ${errorMessage}`;
    
    // Always log to Apps Script logger
    console.error(fullMessage);
    
    // Log to error sheet if configured
    if (logToSheet) {
        try {
            const scriptProperties = PropertiesService.getScriptProperties();
            const errorLogSheetId = scriptProperties.getProperty('ERROR_LOG_SHEET_ID');
            
            if (errorLogSheetId) {
                // Try to open the error log sheet and append the error
                const errorSheet = SpreadsheetApp.openById(errorLogSheetId).getSheetByName('Errors');
                if (errorSheet) {
                    errorSheet.appendRow([timestamp, context, errorMessage]);
                }
            }
        } catch (logError) {
            // Don't throw if logging fails
            console.error(`Failed to log to error sheet: ${logError.message}`);
        }
    }
}

/**
 * Formats numeric data based on locale preferences
 * @param {number} value The value to format
 * @param {string} type Format type ('currency', 'percent', 'number')
 * @param {Object} options Formatting options
 * @returns {string} Formatted value
 */
function formatLocalizedValue(value, type = 'number', options = {}) {
    const {
        locale = getDefaultLocale(),
        currency = getDefaultCurrency(),
        minimumFractionDigits = 2,
        maximumFractionDigits = 2
    } = options;
    
    if (typeof value !== 'number' || isNaN(value)) {
        return value?.toString() || 'N/A';
    }
    
    try {
        switch (type.toLowerCase()) {
            case 'currency':
                return Intl.NumberFormat(locale, {
                    style: 'currency',
                    currency: currency,
                    minimumFractionDigits,
                    maximumFractionDigits
                }).format(value);
                
            case 'percent':
                return Intl.NumberFormat(locale, {
                    style: 'percent',
                    minimumFractionDigits,
                    maximumFractionDigits
                }).format(value);
                
            default:
                return Intl.NumberFormat(locale, {
                    minimumFractionDigits,
                    maximumFractionDigits
                }).format(value);
        }
    } catch (e) {
        // Fall back to basic formatting if Intl isn't available
        return value.toFixed(minimumFractionDigits);
    }
}

/**
 * Formats a date based on locale preferences
 * @param {Date|string} date The date to format
 * @param {string} locale The locale code (e.g., 'en-US')
 * @param {string} format Format string (long, short, or custom)
 * @returns {string} Formatted date
 */
function formatLocalizedDate(date, locale = getDefaultLocale(), format = 'short') {
    if (!date) return 'N/A';
    
    let parsedDate;
    if (typeof date === 'string') {
        parsedDate = new Date(date);
    } else if (date instanceof Date) {
        parsedDate = date;
    } else {
        return 'Invalid Date';
    }
    
    if (isNaN(parsedDate.getTime())) {
        return 'Invalid Date';
    }
    
    try {
        switch (format) {
            case 'long':
                return Utilities.formatDate(parsedDate, locale, 'MMMM dd, yyyy');
            case 'short':
                return Utilities.formatDate(parsedDate, locale, 'yyyy-MM-dd');
            case 'iso':
                return parsedDate.toISOString().split('T')[0];
            default:
                return Utilities.formatDate(parsedDate, locale, format);
        }
    } catch (e) {
        // Fall back to ISO format if formatting fails
        return parsedDate.toISOString().split('T')[0];
    }
}

/**
 * Creates a scheduled trigger for regular anomaly detection
 * @param {string} frequency How often to run ('daily', 'weekly', 'monthly')
 * @param {Object} options Additional scheduling options
 * @returns {string} ID of the created trigger
 */
function setupScheduledAnomalyDetection(frequency = 'weekly', options = {}) {
    try {
        // Delete any existing anomaly detection triggers
        deleteExistingTriggers('runScheduledAnomalyDetection');
        
        let trigger;
        
        switch (frequency.toLowerCase()) {
            case 'daily':
                trigger = ScriptApp.newTrigger('runScheduledAnomalyDetection')
                    .timeBased()
                    .everyDays(1)
                    .atHour(options.hour || 6)
                    .create();
                break;
                
            case 'weekly':
                trigger = ScriptApp.newTrigger('runScheduledAnomalyDetection')
                    .timeBased()
                    .everyWeeks(1)
                    .onWeekDay(options.weekDay || ScriptApp.WeekDay.MONDAY)
                    .atHour(options.hour || 6)
                    .create();
                break;
                
            case 'monthly':
                // For monthly, we need to use everDays(n) where n is approximately a month
                trigger = ScriptApp.newTrigger('runScheduledAnomalyDetection')
                    .timeBased()
                    .everyDays(30)
                    .atHour(options.hour || 6)
                    .create();
                break;
                
            default:
                throw new Error(`Invalid frequency: ${frequency}`);
        }
        
        // Store frequency and notification preferences
        PropertiesService.getScriptProperties().setProperties({
            'ANOMALY_DETECTION_FREQUENCY': frequency,
            'ANOMALY_NOTIFICATION_EMAIL': options.notificationEmail || Session.getActiveUser().getEmail()
        });
        
        return trigger.getUniqueId();
    } catch (error) {
        logError(`Failed to create scheduled trigger: ${error.message}`, 'setupScheduledAnomalyDetection');
        throw error;
    }
}

/**
 * Deletes existing triggers with the specified function name
 * @param {string} functionName Name of the function to look for
 */
function deleteExistingTriggers(functionName) {
    const triggers = ScriptApp.getProjectTriggers();
    
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === functionName) {
            ScriptApp.deleteTrigger(trigger);
        }
    }
}

/**
 * Runs the scheduled anomaly detection and sends notification
 * This function is called by time-based trigger
 */
async function runScheduledAnomalyDetection() {
    try {
        logMessage('Starting scheduled anomaly detection');
        
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getActiveSheet();
        const anomalies = await detectAnomalies(sheet);
        
        if (anomalies.length === 0) {
            logMessage('No anomalies detected in scheduled run');
            return;
        }
        
        // Create error report sheet
        ReportUtils.createErrorReportSheet(anomalies);
        
        // Send notification email
        const notificationEmail = PropertiesService.getScriptProperties().getProperty('ANOMALY_NOTIFICATION_EMAIL');
        if (notificationEmail) {
            const subject = `Anomaly Detection Report - ${new Date().toLocaleDateString()}`;
            const highConfidence = anomalies.filter(a => (a.confidence || 1.0) > 0.8).length;
            const mediumConfidence = anomalies.filter(a => (a.confidence || 1.0) <= 0.8 && (a.confidence || 1.0) > 0.5).length;
            const lowConfidence = anomalies.filter(a => (a.confidence || 1.0) <= 0.5).length;
            
            let body = `Automated anomaly detection completed at ${new Date().toLocaleString()}\n\n`;
            body += `${anomalies.length} potential issues were found:\n`;
            body += `• ${highConfidence} high confidence issues\n`;
            body += `• ${mediumConfidence} medium confidence issues\n`;
            body += `• ${lowConfidence} low confidence issues\n\n`;
            body += `Results are available in the "Error Report" sheet.\n`;
            body += `Spreadsheet URL: ${spreadsheet.getUrl()}`;
            
            GmailApp.sendEmail(notificationEmail, subject, body);
        }
        
        logMessage(`Completed scheduled anomaly detection, found ${anomalies.length} issues`);
    } catch (error) {
        logError(`Scheduled anomaly detection failed: ${error.message}`, 'runScheduledAnomalyDetection');
        
        // Try to send error notification
        try {
            const notificationEmail = PropertiesService.getScriptProperties().getProperty('ANOMALY_NOTIFICATION_EMAIL');
            if (notificationEmail) {
                GmailApp.sendEmail(
                    notificationEmail,
                    'Anomaly Detection Error',
                    `The scheduled anomaly detection failed with error: ${error.message}`
                );
            }
        } catch (emailError) {
            logError(`Failed to send error email: ${emailError.message}`, 'runScheduledAnomalyDetection');
        }
    }
}

/**
 * Deletes existing triggers with the specified function name
 * @param {string} functionName Name of the function to look for
 */
function deleteExistingTriggers(functionName) {
    const triggers = ScriptApp.getProjectTriggers();
    
    for (const trigger of triggers) {
        if (trigger.getHandlerFunction() === functionName) {
            ScriptApp.deleteTrigger(trigger);
        }
    }
}

/**
 * Runs the scheduled anomaly detection and sends notification
 * This function is called by time-based trigger
 */
async function runScheduledAnomalyDetection() {
    try {
        logMessage('Starting scheduled anomaly detection');
        
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = spreadsheet.getActiveSheet();
        const anomalies = await detectAnomalies(sheet);
        
        if (anomalies.length === 0) {
            logMessage('No anomalies detected in scheduled run');
            return;
        }
        
        // Create error report sheet
        createErrorReportSheet(anomalies);
        
        // Send notification email
        const notificationEmail = PropertiesService.getScriptProperties().getProperty('ANOMALY_NOTIFICATION_EMAIL');
        if (notificationEmail) {
            const subject = `Anomaly Detection Report - ${new Date().toLocaleDateString()}`;
            const highConfidence = anomalies.filter(a => (a.confidence || 1.0) > 0.8).length;
            const mediumConfidence = anomalies.filter(a => (a.confidence || 1.0) <= 0.8 && (a.confidence || 1.0) > 0.5).length;
            const lowConfidence = anomalies.filter(a => (a.confidence || 1.0) <= 0.5).length;
            
            let body = `Automated anomaly detection completed at ${new Date().toLocaleString()}\n\n`;
            body += `${anomalies.length} potential issues were found:\n`;
            body += `• ${highConfidence} high confidence issues\n`;
            body += `• ${mediumConfidence} medium confidence issues\n`;
            body += `• ${lowConfidence} low confidence issues\n\n`;
            body += `Results are available in the "Error Report" sheet.\n`;
            body += `Spreadsheet URL: ${spreadsheet.getUrl()}`;
            
            GmailApp.sendEmail(notificationEmail, subject, body);
        }
        
        logMessage(`Completed scheduled anomaly detection, found ${anomalies.length} issues`);
    } catch (error) {
        logError(`Scheduled anomaly detection failed: ${error.message}`, 'runScheduledAnomalyDetection');
        
        // Try to send error notification
        try {
            const notificationEmail = PropertiesService.getScriptProperties().getProperty('ANOMALY_NOTIFICATION_EMAIL');
            if (notificationEmail) {
                GmailApp.sendEmail(
                    notificationEmail,
                    'Anomaly Detection Error',
                    `The scheduled anomaly detection failed with error: ${error.message}`
                );
            }
        } catch (emailError) {
            logError(`Failed to send error email: ${emailError.message}`, 'runScheduledAnomalyDetection');
        }
    }
}

/**
 * Detects and summarizes anomalies in the sheet.
 * @param {number} threshold The threshold for anomaly detection.
 * @returns {Promise<void>}
 */
async function detectAndSummarizeAnomalies(threshold = 3) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const anomalies = await detectAnomalies(sheet, { threshold });
        
        if (anomalies.length === 0) {
            logMessage('No anomalies detected.');
            return;
        }
        
        // Summarize anomalies
        const summary = anomalies.reduce((acc, anomaly) => {
            anomaly.errors.forEach(error => {
                if (!acc[error]) acc[error] = 0;
                acc[error]++;
            });
            return acc;
        }, {});
        
        logMessage('Anomaly Summary:', JSON.stringify(summary, null, 2));
    } catch (error) {
        logError(`Error in detectAndSummarizeAnomalies: ${error.message}`);
    }
}

/**
 * Performs a monthly comparison of anomalies.
 */
function performMonthlyComparison() {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const lastMonth = new Date();
        lastMonth.setMonth(lastMonth.getMonth() - 1);
        
        const anomalies = detectAnomalies(sheet, { dateRange: { start: lastMonth } });
        const currentAnomalies = detectAnomalies(sheet);
        
        const comparison = {
            lastMonth: anomalies.length,
            currentMonth: currentAnomalies.length,
            difference: currentAnomalies.length - anomalies.length
        };
        
        logMessage('Monthly Comparison:', JSON.stringify(comparison, null, 2));
    } catch (error) {
        logError(`Error in performMonthlyComparison: ${error.message}`);
    }
}

/**
 * Runs the monthly comparison of anomalies.
 * @returns {Promise<void>}
 */
async function runMonthlyComparison() {
    try {
        performMonthlyComparison();
    } catch (error) {
        logError(`Error in runMonthlyComparison: ${error.message}`);
    }
}

/**
 * Generates a monthly comparison report.
 * @returns {Promise<void>}
 */
async function generateMonthlyComparisonReport() {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const lastMonth = new Date();
        lastMonth.setMonth(lastMonth.getMonth() - 1);
        
        const anomalies = await detectAnomalies(sheet, { dateRange: { start: lastMonth } });
        const currentAnomalies = await detectAnomalies(sheet);
        
        const report = {
            lastMonth: anomalies.length,
            currentMonth: currentAnomalies.length,
            difference: currentAnomalies.length - anomalies.length
        };
        
        logMessage('Monthly Comparison Report:', JSON.stringify(report, null, 2));
    } catch (error) {
        logError(`Error in generateMonthlyComparisonReport: ${error.message}`);
    }
}

/**
 * Shows a dialog for pattern analysis.
 */
function showPatternAnalysisDialog() {
    const html = HtmlService.createHtmlOutputFromFile('PatternAnalysisDialog')
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(html, 'Pattern Analysis');
}

/**
 * Generates a pattern analysis report.
 * @param {string} analysisType The type of analysis to perform.
 * @param {boolean} includeVisuals Whether to include visuals in the report.
 * @param {boolean} includeAI Whether to include AI analysis in the report.
 * @returns {Promise<void>}
 */
async function generatePatternAnalysisReport(analysisType, includeVisuals, includeAI) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const data = sheet.getDataRange().getValues();
        
        const report = await runFocusedAnalysis(data, analysisType);
        
        if (includeVisuals) {
            // Add visuals to the report
        }
        
        if (includeAI) {
            // Add AI analysis to the report
        }
        
        logMessage('Pattern Analysis Report:', report);
    } catch (error) {
        logError(`Error in generatePatternAnalysisReport: ${error.message}`);
    }
}

/**
 * Analyzes the entire sheet.
 * @returns {Promise<void>}
 */
async function analyzeSheet() {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const data = sheet.getDataRange().getValues();
        
        const analysis = await runFocusedAnalysis(data, 'general');
        logMessage('Sheet Analysis:', analysis);
    } catch (error) {
        logError(`Error in analyzeSheet: ${error.message}`);
    }
}

/**
 * Analyzes the selected data in the sheet.
 */
function analyzeSelectedData() {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        const range = sheet.getActiveRange();
        const data = range.getValues();
        
        analyzeDataSelection(data);
    } catch (error) {
        logError(`Error in analyzeSelectedData: ${error.message}`);
    }
}

/**
 * Analyzes a specific selection of data.
 * @param {Array<Array<any>>} data The data to analyze.
 * @returns {Promise<void>}
 */
async function analyzeDataSelection(data) {
    try {
        const analysis = await runFocusedAnalysis(data, 'selection');
        logMessage('Data Selection Analysis:', analysis);
    } catch (error) {
        logError(`Error in analyzeDataSelection: ${error.message}`);
    }
}

/**
 * Detects column types in the data.
 * @param {Array<Array<any>>} data The data to analyze.
 * @returns {Object} The detected column types.
 */
function detectColumnTypes(data) {
    const columnTypes = {};
    
    if (data.length === 0) return columnTypes;
    
    const headers = data[0];
    const sampleRow = data[1];
    
    headers.forEach((header, index) => {
        const value = sampleRow[index];
        
        if (typeof value === 'number') {
            columnTypes[header] = 'number';
        } else if (typeof value === 'string' && !isNaN(Date.parse(value))) {
            columnTypes[header] = 'date';
        } else {
            columnTypes[header] = 'string';
        }
    });
    
    return columnTypes;
}