/**
 * Retrieves the Gemini API key from script properties.
 * @returns {string} The Gemini API key.
 */
function getGeminiAPIKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

/**
 * Retrieves the configured default locale
 * @returns string
 */
function getDefaultLocale() {
  return PropertiesService.getScriptProperties().getProperty('DEFAULT_LOCALE') || 'en-US'; // Default to US English
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
 * Gets redirect URI
 * @returns
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
 * Retrieves the user configuration, merging with default values.
 * @returns {Config} The configuration object.
 */
function getConfig() {
  // Store default configuration values within script properties
  const defaultConfig = {
     amount: {
        min: 0,
        max: 10000,
        allowNegative: false
    },
    date: {
        datePatterns: [
          /^(\d{4})-(\d{2})-(\d{2})$/, // yyyy-mm-dd
          /^(\d{2})\/(\d{2})\/(\d{4})$/  // mm/dd/yyyy
        ],
        allowFuture: false
    },
    description: {
        required: true
    },
     category: {
        required: true,
        validCategories: ['Sales', 'Marketing', 'Development', 'HR', 'Operations', 'Other']
    },
    email: {
        required: false,
        format: /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    },
    outliers: {
        check: true,
        threshold: 3,
        method: 'zscore', // or 'iqr'
        iqrFactor: 1.5
    },
    duplicates: {
        check: true,
        uniqueColumns: ['amount', 'date', 'description']
    },
    mandatoryFields: ['amount', 'date', 'description', 'category'],

  };

  try{

    const userConfigString = PropertiesService.getScriptProperties().getProperty('USER_CONFIG');
    if (userConfigString) {

        const userConfig = JSON.parse(userConfigString);
        return mergeConfigs(defaultConfig,userConfig)
    } else {
      // Return the default config in case the user has not set up a USER_CONFIG
       return defaultConfig;
    }

  } catch(error) {
     logError("Error getting config: " + error);
     return defaultConfig; // Fallback

  }
}

/**
 * Set Configuration
 * @param {Config} newConfig
 */
function setConfig(newConfig){

  try {
      PropertiesService.getScriptProperties().setProperty('USER_CONFIG', JSON.stringify(newConfig));
      logMessage("Configuration Saved");
      return "Configuration Saved Successfully"

  } catch (error) {
    logError("Error Setting Config" + error);
    return "Error Setting Config:" + error;
  }

}

/**
 * Merges a user-provided configuration with the default configuration.
 */
function mergeConfigs(defaultConfig, userConfig) {
    const merged = { ...defaultConfig }; // Create a copy of defaultConfig

    for (const key in userConfig) {
        if (userConfig.hasOwnProperty(key)) {
            if (typeof defaultConfig[key] === 'object' && !Array.isArray(defaultConfig[key]) &&
                typeof userConfig[key] === 'object' && !Array.isArray(userConfig[key])) {
                // Recursive merge for nested objects
                merged[key] = { ...defaultConfig[key], ...userConfig[key] };
            } else {
                // Override default value with user-provided value
                merged[key] = userConfig[key];
            }
        }
    }
    return merged;
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