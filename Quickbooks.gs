const QUICKBOOKS_CONFIG = {
    authBaseUrl: 'https://appcenter.intuit.com/connect/oauth2',
    tokenUrl: 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer',
    scope: 'com.intuit.quickbooks.accounting',
    responseType: 'code'
};

/**
 * Returns the configured QuickBooks OAuth2 service.
 * @returns {OAuth2Service} The configured OAuth2 service.
 */
function getQuickBooksService() {
  const clientId = getQuickbooksClientId()
  const clientSecret = getQuickbooksClientSecret();
  const redirectUri = getRedirectUri();
    return OAuth2.createService('QuickBooks')
        .setAuthorizationBaseUrl(QUICKBOOKS_CONFIG.authBaseUrl)
        .setTokenUrl(QUICKBOOKS_CONFIG.tokenUrl)
        .setClientId(clientId)
        .setClientSecret(clientSecret)
        .setCallbackFunction('authCallback')
        .setPropertyStore(PropertiesService.getUserProperties())
        .setScope(QUICKBOOKS_CONFIG.scope)
        .setParam('response_type', QUICKBOOKS_CONFIG.responseType);

}

/**
 * Handles the OAuth2 callback.
 * @param {Object} request The request object.
 * @returns {HtmlOutput} An HTML output indicating success or failure.
 */
function authCallback(request) {
    const service = getQuickBooksService();
    const isAuthorized = service.handleCallback(request);
    if (isAuthorized) {
        return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
        return HtmlService.createHtmlOutput('Denied. You can close this tab');
    }
}

/**
 * Imports data from QuickBooks using a QBO API query.
 * @param {string} companyId The QuickBooks company ID.
 * @param {string} query The QBO API query (e.g., "SELECT * FROM Invoice").
 * @returns {Promise<string>} A success or error message.
 */
async function importDataFromQuickBooks(companyId, query) {
    if (!companyId || !query) {
        return 'Company ID and query are required.';
    }


     // Use a try-catch block for the entire operation, for comprehensive error handling
     try {
        const service = getQuickBooksService();

        if (!service.hasAccess()) {

            const authorizationUrl = service.getAuthorizationUrl();
              return 'Authorization needed.  Open this URL to authorize: ' + authorizationUrl;

        }

        const environment = getQuickBooksEnvironment().toUpperCase();
        const baseUrl = QUICKBOOKS_BASE_URLS[environment] || QUICKBOOKS_BASE_URLS.SANDBOX;
        const url = `${baseUrl}${companyId}/query`;

         const headers = {
            Authorization: 'Bearer ' + service.getAccessToken(),
            Accept: 'application/json',
            'Content-Type': 'application/text'  // For the query payload
        };


       const options = {
            method: 'post',
            headers: headers,
            payload: query, //  Body should contain QBO API Query
            muteHttpExceptions: true // Get the full response even on error codes
        };

        const response = UrlFetchApp.fetch(url, options);
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();


       if (responseCode >= 200 && responseCode < 300) {  // Success
            logMessage("Quickbooks Data Pulled Successfully");
            return 'Data imported successfully! ' + responseText;

        } else {
              let errorMessage = `QuickBooks API Error: ${responseCode} - ${responseText}`;

             // Special handling for 401 (Unauthorized)
             if (responseCode === 401) {
                // Token might have expired, clear and tell user to authorize
                 service.reset();
                errorMessage += "  Your session may have expired.  Please re-authorize.";

             }
            logError(errorMessage);
            return errorMessage;
        }

    } catch (error) {
      logError('Error importing data from QuickBooks: ' + error);
      return 'Error importing data: ' + error.message;

    }
}


/**
 * Resets the QuickBooks OAuth2 service, clearing any stored credentials.
 */
function resetQuickBooksAuth() {
    getQuickBooksService().reset();
    return "QuickBooks authorization has been reset.";
}


/**
 * A map of base URLs by environment for QuickBooks Online.
 */
const QUICKBOOKS_BASE_URLS = {
  SANDBOX: 'https://sandbox-quickbooks.api.intuit.com/v3/company/',
  PRODUCTION: 'https://quickbooks.api.intuit.com/v3/company/'
};

/**
 * Logs the redirect URI for debugging.
 * @returns {string}
 */
function logRedirectUri() {
 const redirectUri = getQuickBooksService().getRedirectUri();
  logMessage("Redirect URI" + redirectUri);
  return redirectUri;
}
