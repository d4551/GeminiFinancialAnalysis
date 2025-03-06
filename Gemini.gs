/**
 * Gemini Interaction Service.
 */
class GeminiService {
    constructor() {
        this.config = this.getGeminiConfig();
        this.genAI = new GeminiApp(this.config);
    }

    /**
     * Generates a text response from Gemini based on the user query and sheet data.
     * @param {string} userQuery The user's query.
     * @param {Array<Array<any>>} sheetData The spreadsheet data.
     * @returns {Promise<string>} The generated text response.
     */
    async generateResponse(userQuery, sheetData) {
        if (!userQuery) {
            return "Please enter a query.";
        }

        try {
            const model = this.genAI.getGenerativeModel({ model: await this.getGeminiModel() });

            const chat = model.startChat({
                history: this.createChatHistory(sheetData),
                generationConfig: {
                    maxOutputTokens: 2000,
                    temperature: 0.7,
                },
            });

            const result = await chat.sendMessage(userQuery);
            const response = result.response;
            return response.text();
        } catch (error) {
            this.logError(`Gemini Chat Error: ${error.message}`);
            return `Gemini Chat Error: ${error.message}. Please check logs for details.`;
        }
    }

    /**
     * Analyzes sheet data for anomalies using Gemini AI.
     * @param {Array<Array<any>>} sheetData The spreadsheet data.
     * @returns {Promise<object>} The JSON response containing detected anomalies.
     */
    async analyzeSheetWithGemini(sheetData) {
        try {
            const modelName = await this.getGeminiModel();
            const model = this.genAI.getGenerativeModel({ model: modelName });
            const prompt = this.createAnomalyPrompt(sheetData);
            const result = await model.generateContent(prompt);
            const response = result.response;
            const text = response.text();

            try {
                return JSON.parse(text);
            } catch (e) {
                return this.extractAndValidateJson(text);
            }
        } catch (error) {
            this.logError(`Gemini Anomaly Analysis Error: ${error.message}`);
            throw new Error(`AI Anomaly Analysis Failed: ${error.message}. Please check logs for details.`);
        }
    }

    /**
     * Generates detailed analysis for reports using Gemini AI.
     * @param {string} analysisPrompt The prompt for detailed analysis.
     * @returns {Promise<string>} The detailed analysis text response.
     */
    async generateReportAnalysis(analysisPrompt) {
        if (!analysisPrompt) {
            return "No analysis prompt provided.";
        }

        try {
            const model = this.genAI.getGenerativeModel({ model: await this.getGeminiModel() });
            const result = await model.generateContent(analysisPrompt);
            const response = result.response;
            return response.text();
        } catch (error) {
            this.logError(`Gemini Report Analysis Error: ${error.message}`);
            return `AI Report Analysis Error: ${error.message}. Please check logs for details.`;
        }
    }


    /**
     * Creates a prompt for Gemini to analyze spreadsheet data for anomalies.
     * @param {Array<Array<any>>} sheetData The spreadsheet data.
     * @returns {string} The anomaly analysis prompt.
     */
    createAnomalyPrompt(sheetData) {
        const MAX_LENGTH = 10000;
        const truncatedSheetData = JSON.stringify(sheetData).substring(0, MAX_LENGTH);

        return `You are an intelligent assistant specialized in accounting and financial analysis.  Your task is to identify anomalies in transaction data, such as unusual amounts, dates, or descriptions.

Analyze the provided spreadsheet data and respond in a structured JSON format *without* any additional content or explanation. Do NOT include markdown notation.

Format example:
{"anomalies":[{"row":1,"amount":5000,"date":"2024-01-01","description":"Software Purchase","category":"Sales","email":"","errors":["Large Amount"]},...]}

Data:
${truncatedSheetData}`;
    }

    /**
     * Extracts and validates JSON from a Gemini response string.
     * @param {string} text The Gemini response text.
     * @returns {object} The parsed JSON object.
     * @throws {Error} If no valid JSON is found or JSON structure is invalid.
     */
    extractAndValidateJson(text) {
        try {
            const jsonMatch = text.match(/{\s*"anomalies"\s*:\s*\[[\s\S]*?}\s*]/);
            if (!jsonMatch) {
                throw new Error("No valid JSON found in Gemini response.");
            }
            const jsonString = jsonMatch[0];
            const parsed = JSON.parse(jsonString);

            if (!parsed.anomalies || !Array.isArray(parsed.anomalies)) {
                throw new Error("Invalid JSON structure.  'anomalies' array is missing.");
            }
            return parsed;

        } catch (parseError) {
            this.logError("Error parsing Gemini JSON response: " + parseError.message + "\nResponse: " + text);
            throw new Error("Failed to parse a valid JSON response from Gemini.");
        }
    }

    /**
     * Creates chat history for Gemini conversation.
     * @param {Array<Array<any>>} sheetData The spreadsheet data.
     * @returns {Array<object>} The chat history array.
     */
    createChatHistory(sheetData) {
        const MAX_LENGTH = 10000;
        const truncatedSheetData = JSON.stringify(sheetData).substring(0, MAX_LENGTH);

        return [
            {
                role: "user",
                parts: [{ text: "You are an intelligent assistant that can provide information about spreadsheet data. I will give you data, and then I will ask you questions about it." }],
            },
            {
                role: "model",
                parts: [{ text: "Understood. I'm ready to assist with your spreadsheet data." }],
            },
            {
                role: "user",
                parts: [{ text: "Here is the sheet data:\n" + truncatedSheetData }],
            },
            {
                role: "model",
                parts: [{ text: "Data received. I have the spreadsheet context. Ask me anything." }],
            }
        ];
    }

    /**
     * Gets the Gemini model name to be used.
     * Prioritizes user-selected model, then script property, then default model.
     * @returns {Promise<string>} The Gemini model name.
     */
    async getGeminiModel() {
        let userModel = PropertiesService.getUserProperties().getProperty('USER_SELECTED_MODEL');
        if (userModel) {
            return userModel
        }

        const scriptModel = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL');
        if (scriptModel) {
            return scriptModel;
        }

        return "gemini-1.5-pro-002";
    }

    /**
     * Sets the user-selected Gemini model in user properties.
     * @param {string} modelName The Gemini model name.
     * @returns {string} Success or error message.
     */
    setUserSelectedModel(modelName) {
        try {
            PropertiesService.getUserProperties().setProperty('USER_SELECTED_MODEL', modelName);
            return "Gemini Model Set Successfully";
        } catch (error) {
            this.logError("Error Setting Gemini Model in User Property:" + error);
            return "Error Setting Gemini Model";
        }
    }

    /**
     * Retrieves the Gemini API Key from script properties.
     * @returns {string} The Gemini API key.
     */
    getGeminiAPIKey() {
        return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
    }

    /**
     * Retrieves and constructs the Gemini configuration.
     * Uses Service Account Key if available, otherwise falls back to API Key.
     * @returns {object|string} Gemini configuration object or API key string.
     */
    getGeminiConfig() {
        const apiKey = this.getGeminiAPIKey();
        const serviceAccountKey = this.getServiceAccountKey();

        if (serviceAccountKey) {
            try {
                const parsedCredentials = JSON.parse(serviceAccountKey);
                return {
                    region: YOUR_PROJECT_LOCATION, // Replace with your project location - should likely be a config property too for users.
                    ...parsedCredentials
                };
            } catch (error) {
                this.logError("Error Using Service Account Key:" + error);
                if (apiKey) {
                    return apiKey;
                } else {
                    this.logError("No Valid Gemini Configuration Found.");
                    return "";
                }
            }
        } else {
            if (apiKey) {
                return apiKey;
            } else {
                this.logError("No Valid Gemini Configuration Found.");
                return "";
            }
        }
    }

    /**
     * Sets the Service Account Key in script properties.
     * @param {string} key The Service Account Key in JSON format.
     * @returns {string} Success or error message.
     */
    setServiceAccountKey(key) {
        try {
            PropertiesService.getScriptProperties().setProperty('SERVICE_ACCOUNT_KEY', key);
            return "Service Account Key Set Successfully";
        } catch (error) {
            this.logError("Error Setting Service Account Key:" + error);
            return "Error Setting Service Account Key";
        }
    }

    /**
     * Retrieves the Service Account Key from script properties.
     * @returns {string} The Service Account Key in JSON format.
     */
    getServiceAccountKey() {
        return PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_KEY') || '';
    }

    /**
     * Logs an error message with timestamp.
     * @param {string} message The error message.
     */
    logError(message) {
        Logger.log(`[GeminiService ERROR] ${new Date().toISOString()} - ${message}`);
    }
}

/**
 * Global instance of GeminiService to be used across the application.
 */
const geminiServiceInstance = new GeminiService();

/**
 * Generates a text response from Gemini based on the user query and sheet data.
 * @param {string} userQuery The user's query.
 * @param {Array<Array<any>>} sheetData The spreadsheet data.
 * @returns {Promise<string>} The generated text response.
 */
async function generateResponse(userQuery, sheetData) {
    return geminiServiceInstance.generateResponse(userQuery, sheetData);
}

/**
 * Analyzes sheet data for anomalies using Gemini AI.
 * @param {Array<Array<any>>} sheetData The spreadsheet data.
 * @returns {Promise<object>} The JSON response containing detected anomalies.
 */
async function analyzeSheetWithGemini(sheetData) {
    return geminiServiceInstance.analyzeSheetWithGemini(sheetData);
}

/**
 * Generates detailed analysis for reports using Gemini AI.
 * @param {string} analysisPrompt The prompt for detailed analysis.
 * @returns {Promise<string>} The detailed analysis text response.
 */
async function generateReportAnalysis(analysisPrompt) {
    return geminiServiceInstance.generateReportAnalysis(analysisPrompt);
}


/**
 * Sets the user-selected Gemini model in user properties.
 * @param {string} modelName The Gemini model name.
 * @returns {string} Success or error message.
 */
function setUserSelectedModel(modelName) {
    return geminiServiceInstance.setUserSelectedModel(modelName);
}