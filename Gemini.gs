/**
 * Gemini AI integration functionality
 * Handles interactions with Google's Gemini API
 */

// Note: WORKSPACE_TOOLS is now defined in Code.gs
// const WORKSPACE_TOOLS = ... 

// Get properties and API key setup
const properties = PropertiesService.getScriptProperties().getProperties();
const geminiBaseEndpoint = "https://generativelanguage.googleapis.com/v1beta";

/**
 * Calls the Gemini API to generate a response.
 * @param {string} prompt The text prompt to send to the API.
 * @param {Array<Array<any>>} data Optional data for context.
 * @returns {Promise<string>} The generated response.
 */
async function generateResponse(prompt, data) {
    try {
        const apiKey = getGeminiAPIKey();
        if (!apiKey) {
            return "Gemini API key is not configured. Please set up your API key in the settings.";
        }

        // Get selected model
        const model = getUserSelectedTextModel();
        
        // Format data into a more descriptive prompt
        let enhancedPrompt = prompt;
        if (data && data.length > 0) {
            const headers = data[0];
            // Limit data to avoid token limits
            let rowsToInclude = Math.min(20, data.length - 1);
            const rows = data.slice(1, rowsToInclude + 1); 
            
            enhancedPrompt += "\n\nHere is some context data (showing first rows):\n";
            enhancedPrompt += `Headers: ${JSON.stringify(headers)}\n`;
            enhancedPrompt += `Data: ${JSON.stringify(rows)}`;
        }

        // Call the Gemini API
        const response = await callGeminiAPI(enhancedPrompt, model, apiKey);
        return response;
    } catch (error) {
        logError(`Error in generateResponse: ${error.message}`);
        return `I encountered an error: ${error.message}. Please try again or check your API key configuration.`;
    }
}

/**
 * Calls Gemini API with tools to determine function to call.
 * @param {string} query User query
 * @param {Array<Object>} tools Available tools
 * @returns {Object} Selected tool and arguments
 */
function callGeminiWithTools(query, tools) {
    try {
        const apiKey = getGeminiAPIKey();
        if (!apiKey) {
            throw new Error("Gemini API key is not configured");
        }

        // Get selected model
        const model = getUserSelectedTextModel();

        // Check if we should use REST API or rule-based approach
        const useApi = PropertiesService.getScriptProperties().getProperty('USE_FUNCTION_CALLING_API');
        
        if (useApi === 'true') {
            // Try to use the function calling API
            try {
                return callGeminiFunctionAPI(query, tools, model, apiKey);
            } catch (apiError) {
                logError(`Function calling API error: ${apiError.message}, falling back to rules`);
                // Fall back to rules if API call fails
            }
        }
        
        // Rule-based approach as fallback
        return processQueryWithRules(query);
    } catch (error) {
        logError(`Error in callGeminiWithTools: ${error.message}`);
        // Return default behavior
        return {
            name: "generateResponse",
            args: {}
        };
    }
}

/**
 * Process query using rule-based approach instead of API
 * @param {string} query The user's query
 * @returns {Object} Selected tool and arguments
 */
function processQueryWithRules(query) {
    let toolName = "generateResponse";  // Default
    let args = {};
    
    const queryLower = query.toLowerCase();

    // Simple keyword matching for tool selection
    if (queryLower.includes('transaction') || queryLower.includes('spending')) {
        toolName = "analyzeTransactions";
        args = {
            period: queryLower.includes('month') ? "monthly" : 
                   queryLower.includes('year') ? "yearly" :
                   queryLower.includes('quarter') ? "quarterly" : "all",
            type: queryLower.includes('trend') ? "trends" :
                  queryLower.includes('categor') ? "categories" : "overview"
        };
    } 
    else if (queryLower.includes('report') || queryLower.includes('summary')) {
        toolName = "generateReport";
        const includeAI = !queryLower.includes('without ai');
        
        if (queryLower.includes('executive')) {
            args = { reportType: "executive", includeAI };
        } else if (queryLower.includes('monthly')) {
            args = { reportType: "monthly", includeAI };
        } else if (queryLower.includes('budget')) {
            args = { reportType: "budget", includeAI };
        } else if (queryLower.includes('anomaly')) {
            args = { reportType: "anomaly", includeAI };
        } else {
            args = { reportType: "standard", includeAI };
        }
    }
    else if (queryLower.includes('anomal') || 
             queryLower.includes('unusual') || 
             queryLower.includes('irregularit') ||
             queryLower.includes('outlier')) {
        toolName = "detectAnomalies";
        // Try to extract threshold if mentioned
        const thresholdMatch = query.match(/threshold\s+(\d+(\.\d+)?)/i);
        args = {
            threshold: thresholdMatch ? parseFloat(thresholdMatch[1]) : 3
        };
    }
    else if ((queryLower.includes('compare') && queryLower.includes('month')) ||
             queryLower.includes('month over month') ||
             queryLower.includes('month-over-month') ||
             queryLower.includes('monthly comparison')) {
        toolName = "monthlyComparison";
    }
    else if (queryLower.includes('categor') || 
             (queryLower.includes('spend') && queryLower.includes('by'))) {
        toolName = "categoryAnalysis";
        const topMatch = query.match(/top\s+(\d+)/i);
        const periodMatch = query.match(/for\s+(this month|last month|this year|last year|q[1-4])/i);
        
        args = {
            topCategories: topMatch ? parseInt(topMatch[1]) : 5,
            period: periodMatch ? periodMatch[1].toLowerCase() : "all"
        };
    }

    return {
        name: toolName,
        args: args
    };
}

/**
 * Calls the Gemini API specifically for function calling
 * @param {string} query User query 
 * @param {Array<Object>} tools Tool definitions
 * @param {string} model The model to use
 * @param {string} apiKey The API key
 * @returns {Object} The selected function and arguments
 */
function callGeminiFunctionAPI(query, tools, model, apiKey) {
    // Convert our tools format to Gemini's expected format
    const formattedTools = [{
        functionDeclarations: tools.map(tool => ({
            name: tool.name,
            description: tool.description,
            parameters: {
                type: "OBJECT",
                properties: Object.entries(tool.parameters).map(([name, desc]) => {
                    const [type, description] = desc.split(' - ');
                    return {
                        name: name,
                        type: type.toUpperCase(),
                        description: description || ''
                    };
                }).reduce((obj, param) => {
                    obj[param.name] = {
                        type: param.type,
                        description: param.description
                    };
                    return obj;
                }, {})
            }
        }))
    }];

    // Create the request
    const endpoint = `${geminiBaseEndpoint}/models/${model}:generateContent?key=${apiKey}`;
    const payload = {
        contents: [{ parts: [{ text: query }] }],
        tools: formattedTools,
        generationConfig: {
            temperature: 0.2,
            topP: 0.8,
            topK: 40
        }
    };

    const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(endpoint, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
        throw new Error(`API returned status ${responseCode}: ${response.getContentText()}`);
    }
    
    const jsonResponse = JSON.parse(response.getContentText());
    
    // Extract the function call from the response
    if (jsonResponse.candidates && 
        jsonResponse.candidates[0] && 
        jsonResponse.candidates[0].content && 
        jsonResponse.candidates[0].content.parts && 
        jsonResponse.candidates[0].content.parts[0] && 
        jsonResponse.candidates[0].content.parts[0].functionCall) {
        
        const functionCall = jsonResponse.candidates[0].content.parts[0].functionCall;
        return {
            name: functionCall.name,
            args: functionCall.args
        };
    }
    
    // If no function call was made, fall back to default
    return {
        name: "generateResponse",
        args: {}
    };
}

/**
 * Calls the Gemini API for report analysis.
 * @param {string} prompt The prompt to send.
 * @returns {Promise<string>} The generated analysis.
 */
async function generateReportAnalysis(prompt) {
    try {
        const apiKey = getGeminiAPIKey();
        if (!apiKey) {
            return "Gemini API key is not configured. Please set up your API key in the settings.";
        }
        
        // Get selected model
        const model = getUserSelectedTextModel();

        // Add more focused instructions for financial reports
        const enhancedPrompt = `As a financial analysis expert, ${prompt}

Please be professional, objective, and concise in your analysis. Focus on key insights that would be relevant for financial decision-makers. Use clear, simple language and avoid unnecessary jargon.`;

        // Call the Gemini API
        const response = await callGeminiAPI(enhancedPrompt, model, apiKey);
        return response;
    } catch (error) {
        logError(`Error in generateReportAnalysis: ${error.message}`);
        return `Error generating analysis: ${error.message}`;
    }
}

/**
 * Lists available Gemini models
 * @returns {Array<Object>} List of available models
 */
async function listGeminiModels() {
    try {
        const apiKey = getGeminiAPIKey();
        if (!apiKey) {
            return [];
        }
        
        const endpoint = `${geminiBaseEndpoint}/models?key=${apiKey}`;
        
        try {
            const response = UrlFetchApp.fetch(endpoint, {
                method: 'get',
                muteHttpExceptions: true
            });
            
            const responseCode = response.getResponseCode();
            
            if (responseCode !== 200) {
                logError(`Error fetching models: HTTP ${responseCode}`);
                return getDefaultModels();
            }
            
            const jsonResponse = JSON.parse(response.getContentText());
            
            if (jsonResponse && jsonResponse.models) {
                // Filter for only Gemini models
                const geminiModels = jsonResponse.models.filter(model => 
                    model.name && model.name.includes('gemini')
                );
                
                // Cache the results
                cacheModels(geminiModels);
                
                return geminiModels;
            }
        } catch (fetchError) {
            logError(`Error fetching models from API: ${fetchError.message}`);
        }
        
        // Fall back to default models if API call fails
        return getDefaultModels();
    } catch (error) {
        logError(`Error listing Gemini models: ${error.message}`);
        return getDefaultModels();
    }
}

/**
 * Caches the model list for future use
 * @param {Array<Object>} models The models to cache
 */
function cacheModels(models) {
    if (!models || models.length === 0) return;
    
    try {
        const cache = CacheService.getUserCache();
        cache.put('GEMINI_MODELS', JSON.stringify(models), 3600); // Cache for 1 hour
    } catch (error) {
        logError(`Error caching models: ${error.message}`);
    }
}

/**
 * Gets default models when API isn't available
 * @returns {Array<Object>} Array of default model objects
 */
function getDefaultModels() {
    return [
        { name: 'projects/your-project/models/gemini-1.5-pro-latest' },
        { name: 'projects/your-project/models/gemini-1.0-pro-latest' },
        { name: 'projects/your-project/models/gemini-1.5-pro-vision-latest' },
        { name: 'projects/your-project/models/gemini-1.0-pro-vision-latest' }
    ];
}

/**
 * Makes the actual API call to Gemini.
 * @param {string} prompt The prompt to send.
 * @param {string} model The model to use.
 * @param {string} apiKey The Gemini API key.
 * @returns {Promise<string>} The API response.
 */
async function callGeminiAPI(prompt, model, apiKey) {
    try {
        const endpoint = `${geminiBaseEndpoint}/models/${model}:generateContent?key=${apiKey}`;
        
        const payload = {
            contents: [{
                parts: [{
                    text: prompt
                }]
            }],
            generationConfig: {
                temperature: 0.3,
                topP: 0.8,
                topK: 40,
                maxOutputTokens: 8192
            }
        };
        
        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(endpoint, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
            const errorText = response.getContentText();
            logError(`Gemini API Error (${responseCode}): ${errorText}`);
            
            // If rate limited or model not found, try with default model
            if (responseCode === 429 || responseCode === 404) {
                logMessage("Attempting with default model due to error");
                return fallbackToDefaultModel(prompt, apiKey);
            }
            
            throw new Error(`API returned status ${responseCode}`);
        }
        
        const jsonResponse = JSON.parse(response.getContentText());
        
        if (jsonResponse.candidates && 
            jsonResponse.candidates[0] && 
            jsonResponse.candidates[0].content && 
            jsonResponse.candidates[0].content.parts && 
            jsonResponse.candidates[0].content.parts[0] && 
            jsonResponse.candidates[0].content.parts[0].text) {
            
            return jsonResponse.candidates[0].content.parts[0].text;
        } else {
            // If model returns no content
            if (jsonResponse.promptFeedback && jsonResponse.promptFeedback.blockReason) {
                return `The AI model couldn't process this request: ${jsonResponse.promptFeedback.blockReason}`;
            }
            
            // Otherwise generic error
            throw new Error("Unable to generate content from Gemini");
        }
    } catch (error) {
        logError(`Error in callGeminiAPI: ${error.message}`);
        
        // Check if error is likely due to API key issues
        if (error.message.includes("API key") || error.message.includes("403")) {
            return "There appears to be an issue with your Gemini API key. Please check your configuration.";
        }
        
        throw new Error(`Failed to call Gemini API: ${error.message}`);
    }
}

/**
 * Fallback to default model if specified model has issues
 * @param {string} prompt The prompt to send
 * @param {string} apiKey The API key to use
 * @returns {string} The generated response
 */
function fallbackToDefaultModel(prompt, apiKey) {
    const defaultModel = "gemini-1.0-pro"; // Most stable fallback
    
    try {
        logMessage(`Falling back to ${defaultModel} due to issues with primary model`);
        
        const endpoint = `${geminiBaseEndpoint}/models/${defaultModel}:generateContent?key=${apiKey}`;
        
        const payload = {
            contents: [{
                parts: [{
                    text: prompt
                }]
            }],
            generationConfig: {
                temperature: 0.2,
                maxOutputTokens: 4096
            }
        };
        
        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(endpoint, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
            throw new Error(`Fallback API also returned status ${responseCode}`);
        }
        
        const jsonResponse = JSON.parse(response.getContentText());
        
        if (jsonResponse.candidates && 
            jsonResponse.candidates[0] && 
            jsonResponse.candidates[0].content && 
            jsonResponse.candidates[0].content.parts && 
            jsonResponse.candidates[0].content.parts[0] && 
            jsonResponse.candidates[0].content.parts[0].text) {
            
            return jsonResponse.candidates[0].content.parts[0].text + "\n\n(Note: Response generated using fallback model)";
        } else {
            throw new Error("Fallback model unable to generate content");
        }
    } catch (error) {
        logError(`Fallback model error: ${error.message}`);
        return "Unable to generate response with the AI models. Please check your API key configuration and try again later.";
    }
}

/**
 * Analyzes a sheet with Gemini AI
 * @param {Array<Array<any>>} sheetData Sheet data to analyze
 * @returns {Object} AI analysis results
 */
async function analyzeSheetWithGemini(sheetData) {
    try {
        const apiKey = getGeminiAPIKey();
        if (!apiKey) {
            return {
                anomalies: [],
                insights: "Gemini API key is not configured. Please set up your API key in the settings."
            };
        }
        
        // Get selected model
        const model = getUserSelectedTextModel();
        
        // Create headers and sample rows for analysis
        const headers = sheetData[0];
        
        // Limit the amount of data sent to the API
        const MAX_ROWS = 100;
        const sampleRows = sheetData.slice(1, Math.min(1 + MAX_ROWS, sheetData.length));
        
        const prompt = `Analyze this financial data:
        
Headers: ${JSON.stringify(headers)}
Sample data (${sampleRows.length} rows): ${JSON.stringify(sampleRows)}

1. Identify any anomalies or issues in the data (missing values, outliers, inconsistencies)
2. Provide insights about:
   - Spending patterns
   - Unusual transactions
   - Areas of concern
   - Distribution of categories (if present)
   - Temporal patterns (if dates are present)
3. Suggest specific actions to address the issues

Format your analysis to be clear, concise and actionable. Focus on financial insights that would be valuable for decision-making.`;
        
        // Call the Gemini API
        const response = await callGeminiAPI(prompt, model, apiKey);
        
        // We'll parse the response to extract structured anomalies if needed
        let anomalies = [];
        
        // Try to extract anomalies if response looks like it contains them
        if (response.includes("[") && response.includes("]") && response.includes("row")) {
            try {
                // Look for JSON-like patterns in the response
                const anomalyMatch = response.match(/\[[\s\S]*?\]/);
                if (anomalyMatch) {
                    const potentialJson = anomalyMatch[0];
                    anomalies = JSON.parse(potentialJson);
                }
            } catch (parseError) {
                // Parsing failed, but we can still use the text response
                logError(`Error parsing anomalies from response: ${parseError.message}`);
            }
        }
        
        return {
            anomalies: anomalies,
            insights: response
        };
    } catch (error) {
        logError(`Error in analyzeSheetWithGemini: ${error.message}`);
        return {
            anomalies: [],
            insights: `Error analyzing sheet: ${error.message}`
        };
    }
}

/**
 * Get information about available models and current selections for UI display
 * @returns {Object} Model information
 */
function getGeminiModelInfo() {
    try {
        // Get current model selections
        const currentTextModel = getUserSelectedTextModel();
        const currentVisionModel = getUserSelectedVisionModel();
        
        // Get available models from cache or API
        let availableModels = getCachedGeminiModels();
        
        // If no models in cache, use defaults
        if (!availableModels || availableModels.length === 0) {
            availableModels = getDefaultModels();
            
            // Try to refresh in background
            refreshModelsInBackground();
        }
        
        return {
            availableModels,
            currentTextModel,
            currentVisionModel
        };
    } catch (error) {
        logError(`Error getting Gemini model info: ${error.message}`);
        throw new Error(`Failed to get model information: ${error.message}`);
    }
}

/**
 * Asynchronously refreshes the model list in the background
 */
function refreshModelsInBackground() {
    // This kicks off the model list update without waiting for it
    listGeminiModels();
}

/**
 * Analyzes categories with Gemini AI
 * @returns {Promise<string>} Analysis of categories
 */
async function analyzeCategoriesWithGemini() {
    const sheetData = getSheetData();
    
    const prompt = `Analyze the spending categories in this financial data: 
    
${JSON.stringify(sheetData.slice(0, 50))}

Identify the top categories by total spend, show their percentage of total expenses, 
and provide insights about spending patterns. If you detect any unusual category allocations, 
please highlight them.

Format your response with proper sections and bullet points for clarity.`;
    
    return await generateResponse(prompt, null);
}

/**
 * Performs an advanced image analysis using Gemini Vision models
 * @param {Blob} imageBlob The image blob to analyze
 * @param {string} prompt The prompt for analysis
 * @returns {string} Analysis results
 */
async function analyzeImageWithGemini(imageBlob, prompt = null) {
    try {
        const apiKey = getGeminiAPIKey();
        if (!apiKey) {
            return "Gemini API key is not configured. Please set up your API key in the settings.";
        }
        
        // Get selected vision model
        const model = getUserSelectedVisionModel();
        
        // Base64 encode the image
        const imageData = Utilities.base64Encode(imageBlob.getBytes());
        
        // Create a default prompt if none provided
        const analysisPrompt = prompt || "Analyze this financial chart or document. Explain what it shows, identify key trends or data points, and provide any relevant financial insights.";
        
        const endpoint = `${geminiBaseEndpoint}/models/${model}:generateContent?key=${apiKey}`;
        
        const payload = {
            contents: [{
                parts: [
                    { text: analysisPrompt },
                    {
                        inline_data: {
                            mime_type: imageBlob.getContentType() || "image/png",
                            data: imageData
                        }
                    }
                ]
            }],
            generationConfig: {
                temperature: 0.2,
                topP: 0.8,
                topK: 40,
                maxOutputTokens: 8192
            }
        };
        
        const options = {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(endpoint, options);
        const responseCode = response.getResponseCode();
        
        if (responseCode !== 200) {
            const errorText = response.getContentText();
            logError(`Gemini Vision API Error (${responseCode}): ${errorText}`);
            throw new Error(`API returned status ${responseCode}`);
        }
        
        const jsonResponse = JSON.parse(response.getContentText());
        
        if (jsonResponse.candidates && 
            jsonResponse.candidates[0] && 
            jsonResponse.candidates[0].content && 
            jsonResponse.candidates[0].content.parts && 
            jsonResponse.candidates[0].content.parts[0] && 
            jsonResponse.candidates[0].content.parts[0].text) {
            
            return jsonResponse.candidates[0].content.parts[0].text;
        } else {
            if (jsonResponse.promptFeedback && jsonResponse.promptFeedback.blockReason) {
                return `The AI model couldn't process this image: ${jsonResponse.promptFeedback.blockReason}`;
            }
            
            throw new Error("Unable to analyze image with Gemini");
        }
    } catch (error) {
        logError(`Error in analyzeImageWithGemini: ${error.message}`);
        return `Error analyzing image: ${error.message}`;
    }
}

/**
 * Runs a focused analysis on the provided data.
 * @param {Array<Object>} data The data to analyze.
 * @param {string} analysisType The type of analysis to perform.
 * @param {Object} dateRange Optional date range for filtering.
 * @returns {Promise<string>} The analysis results.
 */
async function runFocusedAnalysis(data, analysisType, dateRange) {
    try {
        // Filter by date if dateRange is provided
        if (dateRange && dateRange.start && dateRange.end) {
            data = data.filter(row => {
                const rowDate = new Date(row.date);
                return rowDate >= dateRange.start && rowDate <= dateRange.end;
            });
        }

        if (data.length === 0) {
            return "No data found for the specified date range.";
        }

        // Define analysis prompt based on analysis type
        let prompt = "";
        switch (analysisType.toLowerCase()) {
            case "spending_trends":
                prompt = `Analyze the following financial data and identify spending trends over time. 
                Focus on patterns, increases/decreases, and seasonal variations if present.
                ${getDateRangeText(dateRange)}`;
                break;

            case "category_breakdown":
                prompt = `Analyze the following financial data and provide a breakdown by category. 
                Identify top spending categories, calculate percentages of total spending, and highlight any
                categories with unusual spending patterns.
                ${getDateRangeText(dateRange)}`;
                break;

            case "anomaly_detection":
                prompt = `Analyze the following financial data and identify any anomalies or outliers. 
                Focus on transactions that deviate significantly from normal patterns, 
                unusual spending amounts, or unexpected transactions.
                ${getDateRangeText(dateRange)}`;
                break;

            case "budget_comparison":
                prompt = `Compare the actual spending in the following financial data against typical budgets.
                Identify areas of overspending and underspending, and suggest potential budget adjustments.
                ${getDateRangeText(dateRange)}`;
                break;

            case "monthly_trends":
                prompt = `Analyze the following financial data and compare spending across different months.
                Identify month-over-month changes, consistent patterns, and noteworthy variations.
                ${getDateRangeText(dateRange)}`;
                break;

            default:
                prompt = `Analyze the following financial data and provide insights.
                ${getDateRangeText(dateRange)}`;
        }

        // Prepare data for analysis - convert to 2D array format for generateResponse
        const headers = Object.keys(data[0]);
        const rows = data.map(row => Object.values(row));
        const formattedData = [headers, ...rows];

        // Call Gemini API with the prompt and data
        return await generateResponse(prompt, formattedData);
    } catch (error) {
        logError(`Error in runFocusedAnalysis: ${error.message}`);
        return `Error performing ${analysisType} analysis: ${error.message}`;
    }
}

/**
 * Helper function to format date range text for prompts
 * @param {Object} dateRange The date range object
 * @returns {string} Formatted date range text
 */
function getDateRangeText(dateRange) {
    if (!dateRange || !dateRange.start || !dateRange.end) {
        return "Analysis covers all available data.";
    }
    
    const startDate = new Date(dateRange.start).toLocaleDateString();
    const endDate = new Date(dateRange.end).toLocaleDateString();
    return `Analysis covers data from ${startDate} to ${endDate}.`;
}