/**
 * Configuration management for Gemini Financial Analysis
 * Provides centralized access to all configuration settings
 */

/**
 * Retrieves the Gemini API key from script properties.
 * @returns {string} The Gemini API key.
 */
function getGeminiAPIKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

/**
 * Get the user-selected text model for Gemini
 * @returns {string} Selected model name or default
 */
function getUserSelectedTextModel() {
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_TEXT_MODEL');
  return model || 'gemini-1.5-pro-latest'; // Default to 1.5 Pro
}

/**
 * Get the user-selected vision model for Gemini
 * @returns {string} Selected model name or default
 */
function getUserSelectedVisionModel() {
  const model = PropertiesService.getScriptProperties().getProperty('GEMINI_VISION_MODEL');
  return model || 'gemini-1.5-pro-vision-latest'; // Default to 1.5 Pro Vision
}

/**
 * Get cached Gemini models to avoid frequent API calls
 * @returns {Array} Array of model objects
 */
function getCachedGeminiModels() {
  try {
    const cache = CacheService.getUserCache();
    const cachedModels = cache.get('GEMINI_MODELS');
    
    if (cachedModels) {
      return JSON.parse(cachedModels);
    }
    return null;
  } catch (error) {
    logError(`Error getting cached models: ${error.message}`);
    return null;
  }
}

/**
 * Stores Gemini models in cache to improve performance
 * @param {Array} models Array of model objects to cache
 * @param {number} cacheTime Time in seconds to cache (default: 1 hour)
 * @returns {boolean} True if successful
 */
function cacheGeminiModels(models, cacheTime = 3600) {
  try {
    const cache = CacheService.getUserCache();
    cache.put('GEMINI_MODELS', JSON.stringify(models), cacheTime);
    return true;
  } catch (error) {
    logError(`Error caching models: ${error.message}`);
    return false;
  }
}

/**
 * Retrieves the configured default locale
 * @returns {string} The configured locale
 */
function getDefaultLocale() {
  return PropertiesService.getScriptProperties().getProperty('DEFAULT_LOCALE') || 'en-US';
}

/**
 * Retrieves default currency setting
 * @returns {string} The default currency code
 */
function getDefaultCurrency() {
  return PropertiesService.getScriptProperties().getProperty('DEFAULT_CURRENCY') || 'USD';
}

/**
 * Retrieves QuickBooks Client Id
 * @returns {string} The QuickBooks Client Id.
 */
function getQuickbooksClientId(){
  return PropertiesService.getScriptProperties().getProperty('QB_CLIENT_ID') || '';
}

/**
 * Retrieves QuickBooks Secret Id
 * @returns {string} The QuickBooks Secret Id.
 */
function getQuickbooksClientSecret(){
  return PropertiesService.getScriptProperties().getProperty('QB_CLIENT_SECRET') || '';
}

/**
 * Gets redirect URI for OAuth
 * @returns {string} OAuth redirect URI
 */
function getRedirectUri(){
 return ScriptApp.getService().getUrl();
}

/**
 * Retrieves QuickBooks Environment
 * @returns {string} The QuickBooks Environment.
 */
function getQuickBooksEnvironment(){
  const env = PropertiesService.getScriptProperties().getProperty('QUICKBOOKS_ENV') || 'SANDBOX';
  return env.toUpperCase();
}

/**
 * Checks if AI features are enabled
 * @returns {boolean} Whether AI features are enabled
 */
function isAIEnabled() {
  const setting = PropertiesService.getScriptProperties().getProperty('ENABLE_AI');
  return setting !== 'false'; // Default to enabled if not explicitly disabled
}

/**
 * Retrieves the user configuration, merging with default values.
 * @returns {Object} The configuration object.
 */
function getConfig() {
  try {
    // Attempt to load from script properties
    const scriptProps = PropertiesService.getScriptProperties();
    const userConfigString = scriptProps.getProperty('USER_CONFIG');
    const defaultConfig = getDefaultConfig();
    
    if (userConfigString) {
      try {
        const userConfig = JSON.parse(userConfigString);
        return mergeConfigs(defaultConfig, userConfig);
      } catch (parseError) {
        logError(`Error parsing user config: ${parseError.message}`);
        return defaultConfig;
      }
    } else {
      // If no user config exists, apply any individual settings to the default
      const config = { ...defaultConfig };
      
      // Apply individual settings that might be set outside the USER_CONFIG
      const enableAI = scriptProps.getProperty('ENABLE_AI');
      if (enableAI !== null) {
        config.enableAIDetection = (enableAI === 'true');
      }
      
      // Add report configuration
      config.reportConfig = getReportingDefaults();
      
      return config;
    }
  } catch (error) {
    // Fallback to defaults on error
    logError(`Error getting config: ${error.message}`);
    return getDefaultConfig();
  }
}

/**
 * Get default reporting configuration settings
 * @returns {Object} Default reporting configuration
 */
function getReportingDefaults() {
  return {
    includeAIInsights: isAIEnabled(),
    includeTableOfContents: false,
    includeCharts: true,
    includeConfidenceScores: true,
    formatOptions: {
      locale: getDefaultLocale(),
      currency: getDefaultCurrency(),
      dateFormat: 'yyyy-MM-dd'
    },
    templates: {
      // Standard templates are defined in REPORT_CONFIG in ReportGeneration.gs
    }
  };
}

/**
 * Set Configuration with validation
 * @param {Object} newConfig Configuration object
 * @returns {string} Status message
 */
function setConfig(newConfig){
  try {
    // Validate critical config values
    if (newConfig.outliers && newConfig.outliers.threshold) {
      if (isNaN(newConfig.outliers.threshold) || newConfig.outliers.threshold <= 0) {
        throw new Error("Outlier threshold must be a positive number");
      }
    }
    
    // Sanitize and save
    const sanitizedConfig = sanitizeConfigObject(newConfig);
    PropertiesService.getScriptProperties().setProperty('USER_CONFIG', JSON.stringify(sanitizedConfig));
    logMessage("Configuration Saved");
    return "Configuration Saved Successfully";
  } catch (error) {
    logError("Error Setting Config: " + error);
    throw new Error("Error Setting Config: " + error.message);
  }
}

/**
 * Updates just the reporting configuration section
 * @param {Object} reportConfig Reporting configuration object
 * @returns {string} Status message
 */
function updateReportConfig(reportConfig) {
  try {
    const currentConfig = getConfig();
    currentConfig.reportConfig = {...currentConfig.reportConfig, ...reportConfig};
    return setConfig(currentConfig);
  } catch (error) {
    logError(`Error updating report config: ${error.message}`);
    throw new Error(`Failed to update report configuration: ${error.message}`);
  }
}

/**
 * Sanitizes a configuration object to prevent injection attacks
 * @param {Object} configObj The config object to sanitize
 * @returns {Object} Sanitized config object
 */
function sanitizeConfigObject(configObj) {
  // Deep clone the object to avoid modifying the original
  const sanitized = JSON.parse(JSON.stringify(configObj));
  
  // Helper function for recursive sanitization
  function sanitizeValue(value) {
    if (typeof value === 'string') {
      // Basic sanitization for strings - remove script tags
      return value.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '');
    } else if (typeof value === 'object' && value !== null) {
      // Recursively sanitize objects (including arrays)
      Object.keys(value).forEach(key => {
        value[key] = sanitizeValue(value[key]);
      });
    }
    return value;
  }
  
  return sanitizeValue(sanitized);
}

/**
 * Merges a user-provided configuration with the default configuration.
 * @param {Object} defaultConfig Default configuration
 * @param {Object} userConfig User-provided configuration
 * @returns {Object} Merged configuration
 */
function mergeConfigs(defaultConfig, userConfig) {
  const merged = { ...defaultConfig }; // Create a copy of defaultConfig

  for (const key in userConfig) {
    if (userConfig.hasOwnProperty(key)) {
      if (typeof defaultConfig[key] === 'object' && !Array.isArray(defaultConfig[key]) &&
          typeof userConfig[key] === 'object' && !Array.isArray(userConfig[key])) {
        // Recursive merge for nested objects
        merged[key] = mergeConfigs(defaultConfig[key] || {}, userConfig[key]);
      } else {
        // Override default value with user-provided value
        merged[key] = userConfig[key];
      }
    }
  }
  return merged;
}

/**
 * Checks if QuickBooks integration is properly configured
 * @returns {boolean} Whether QuickBooks is properly configured
 */
function isQuickBooksConfigured() {
  const clientId = getQuickbooksClientId();
  const clientSecret = getQuickbooksClientSecret();
  return Boolean(clientId && clientSecret);
}

/**
 * Retrieves reporting schedule configuration
 * @returns {Object} The scheduled report configuration
 */
function getReportSchedule() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    return {
      frequency: scriptProps.getProperty('SCHEDULED_REPORT_FREQUENCY') || 'none',
      reportType: scriptProps.getProperty('SCHEDULED_REPORT_TYPE') || 'standard',
      email: scriptProps.getProperty('SCHEDULED_REPORT_EMAIL') || '',
      includeAI: scriptProps.getProperty('SCHEDULED_REPORT_INCLUDE_AI') !== 'false'
    };
  } catch (error) {
    logError(`Error getting report schedule: ${error.message}`);
    return {
      frequency: 'none', 
      reportType: 'standard', 
      email: '', 
      includeAI: true
    };
  }
}

/**
 * Sets reporting schedule configuration
 * @param {Object} schedule The schedule configuration
 * @returns {boolean} Success status
 */
function setReportSchedule(schedule) {
  try {
    PropertiesService.getScriptProperties().setProperties({
      'SCHEDULED_REPORT_FREQUENCY': schedule.frequency || 'none',
      'SCHEDULED_REPORT_TYPE': schedule.reportType || 'standard',
      'SCHEDULED_REPORT_EMAIL': schedule.email || '',
      'SCHEDULED_REPORT_INCLUDE_AI': (schedule.includeAI !== false).toString()
    });
    return true;
  } catch (error) {
    logError(`Error setting report schedule: ${error.message}`);
    return false;
  }
}

/**
 * Retrieves anomaly detection schedule configuration
 * @returns {Object} The scheduled anomaly detection configuration
 */
function getAnomalyDetectionSchedule() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    return {
      frequency: scriptProps.getProperty('ANOMALY_DETECTION_FREQUENCY') || 'none',
      notificationEmail: scriptProps.getProperty('ANOMALY_NOTIFICATION_EMAIL') || ''
    };
  } catch (error) {
    logError(`Error getting anomaly detection schedule: ${error.message}`);
    return {frequency: 'none', notificationEmail: ''};
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
      method: 'zscore',
      iqrFactor: 1.5,
      check: true
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
      required: true,
      minLength: 3
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
    includeAIExplanations: true,
    reportConfig: getReportingDefaults()
  };
}

/**
 * Loads data from QuickBooks (stub function)
 * Would normally implement actual API calls
 * @param {string} companyId QuickBooks company ID
 * @param {string} query QuickBooks query
 * @returns {Array<Object>} The retrieved data
 */
async function fetchQuickBooksData(companyId, query) {
  // This is a stub function - real implementation would make API calls
  // In a real implementation, this would use OAuth2 to authenticate with QuickBooks
  // and make the appropriate API calls to retrieve data
  
  // For now, just return mock data
  return [
    ['Date', 'Description', 'Category', 'Amount'],
    ['2023-01-15', 'Office Supplies', 'Expenses', 125.99],
    ['2023-01-28', 'Consulting Services', 'Income', 1500.00],
    ['2023-02-03', 'Software Subscription', 'Expenses', 49.99]
  ];
}

/**
 * Inserts QuickBooks data into a sheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to insert into
 * @param {Array<Array<any>>} data The data to insert
 */
function insertQuickBooksData(sheet, data) {
  if (!sheet || !data || !data.length) return;
  
  try {
    // Clear existing content
    sheet.clear();
    
    // Insert the data
    const range = sheet.getRange(1, 1, data.length, data[0].length);
    range.setValues(data);
    
    // Format the header row
    sheet.getRange(1, 1, 1, data[0].length)
      .setFontWeight('bold')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, data[0].length);
    
    // Format dates and currency if the headers suggest those columns
    const headers = data[0].map(h => String(h).toLowerCase());
    
    headers.forEach((header, index) => {
      if (header.includes('date')) {
        // Format as date
        sheet.getRange(2, index + 1, data.length - 1, 1).setNumberFormat('yyyy-mm-dd');
      }
      
      if (header.includes('amount') || header.includes('price') || header.includes('cost')) {
        // Format as currency
        sheet.getRange(2, index + 1, data.length - 1, 1).setNumberFormat('$#,##0.00');
      }
    });
    
  } catch (error) {
    logError(`Error inserting QuickBooks data: ${error.message}`);
    throw new Error(`Failed to insert data: ${error.message}`);
  }
}

/**
 * Gets the QuickBooks access token (stub function)
 * @returns {string} Access token or empty string
 */
function getQuickBooksAccessToken() {
  return PropertiesService.getScriptProperties().getProperty('QB_ACCESS_TOKEN') || '';
}

/**
 * Creates error report sheet from anomalies data
 * @param {Array<Object>} anomalies The anomalies to display
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The created error sheet
 */
function createErrorReportSheet(anomalies) {
  if (!anomalies || !anomalies.length) {
    return null;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName('Error Report');
    
    // Create or clear the error sheet
    if (errorSheet) {
      errorSheet.clear();
    } else {
      errorSheet = ss.insertSheet('Error Report');
    }
    
    // Create headers based on available fields
    const hasConfidence = anomalies.some(a => a.confidence !== undefined);
    const hasTransactionType = anomalies.some(a => a.transaction_type && a.transaction_type !== 'N/A');
    
    const headers = ['Row', 'Amount', 'Date', 'Description', 'Category'];
    if (hasTransactionType) headers.push('Transaction Type');
    if (hasConfidence) headers.push('Confidence');
    headers.push('Errors');
    
    // Add header row
    errorSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    errorSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold');
    
    // Add data rows
    const rows = anomalies.map(anomaly => {
      const baseRow = [
        anomaly.row || 'N/A',
        anomaly.amount || 'N/A',
        anomaly.date || 'N/A',
        anomaly.description || 'N/A', 
        anomaly.category || 'N/A'
      ];
      
      if (hasTransactionType) baseRow.push(anomaly.transaction_type || 'N/A');
      if (hasConfidence) baseRow.push(`${((anomaly.confidence || 1.0) * 100).toFixed(0)}%`);
      
      baseRow.push(Array.isArray(anomaly.errors) ? anomaly.errors.join(', ') : anomaly.errors || 'N/A');
      
      return baseRow;
    });
    
    if (rows.length > 0) {
      errorSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    // Format date and amount columns
    const dateCol = headers.indexOf('Date') + 1;
    const amountCol = headers.indexOf('Amount') + 1;
    const confidenceCol = headers.indexOf('Confidence') + 1;
    
    if (dateCol > 0) {
      errorSheet.getRange(2, dateCol, rows.length, 1).setNumberFormat('@'); // Format as text to preserve various date formats
    }
    
    if (amountCol > 0) {
      errorSheet.getRange(2, amountCol, rows.length, 1).setNumberFormat('$#,##0.00');
    }
    
    // Auto-resize columns for better visibility
    errorSheet.autoResizeColumns(1, headers.length);
    
    // Set default sort by row number
    errorSheet.sort(1);
    
    // Return the sheet
    return errorSheet;
    
  } catch (error) {
    logError(`Error creating error report sheet: ${error.message}`);
    throw new Error(`Failed to create error report: ${error.message}`);
  }
}

/**
 * @typedef {Object} Config
 * @property {AmountConfig} amount
 * @property {DateConfig} date
 * @property {DescriptionConfig} description
 * @property {CategoryConfig} category
 * @property {EmailConfig} email
 * @property {OutliersConfig} outliers
 * @property {DuplicatesConfig} duplicates
 * @property {string[]} mandatoryFields
 * @property {string} detectionAlgorithm
 * @property {boolean} enableAIDetection
 * @property {boolean} includeAIExplanations
 * @property {Object} reportConfig
 */

/**
 * @typedef {Object} AmountConfig
 * @property {number} min
 * @property {number} max
 * @property {boolean} allowNegative
 */

/**
 * @typedef {Object} DateConfig
 * @property {RegExp[]} datePatterns
 * @property {boolean} allowFuture
 */

/**
 * @typedef {Object} DescriptionConfig
 * @property {boolean} required
 * @property {number} minLength
 */

/**
  * @typedef {Object} CategoryConfig
  * @property {boolean} required
  * @property {string[]} validCategories
  */

/**
 * @typedef {Object} EmailConfig
 * @property {boolean} required
 * @property {RegExp} format
 */

/**
 * @typedef {Object} OutliersConfig
 * @property {boolean} check
 * @property {number} threshold
 * @property {string} method
 * @property {number} iqrFactor
 */

/**
* @typedef {Object} DuplicatesConfig
* @property {boolean} check
* @property {string[]} uniqueColumns
*/