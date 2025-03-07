# Gemini Financial Analysis

> **NOTICE:** This project is undergoing active refactoring to centralize functions and improve design patterns. Some functionality may be in the process of being consolidated across files. Please refer to the code comments for the most up-to-date implementation details.

An open source Google Apps Script application that leverages Google's Gemini AI for advanced financial analysis and anomaly detection in Google Sheets.

![GitHub License](https://img.shields.io/github/license/d4551/GeminiFinancialAnalysis)
![Open Source](https://img.shields.io/badge/Open%20Source-Yes-brightgreen)
![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-V8-blue)
![Gemini AI](https://img.shields.io/badge/Gemini%20AI-1.5%20Pro-orange)

## Overview

Gemini Financial Analysis is an open source tool built to empower financial professionals, small businesses, and individuals to leverage the power of AI for financial data analysis. By combining Google Sheets' familiar interface with Google's Gemini AI capabilities, this application democratizes advanced financial analysis techniques.

## Features

- 🤖 **AI-powered Analysis:** Leverage Google's Gemini large language models for financial insights
- 📊 **Anomaly Detection:** Automatically identify outliers, missing data, duplicates, and patterns
- 📝 **Report Generation:** Create professional financial reports with executive summaries and visualizations
- 💬 **Interactive Chat:** Query your financial data in natural language
- 📈 **Pattern Analysis:** Analyze frequency, temporal, value distribution and category patterns
- 🔄 **QuickBooks Integration:** Import and analyze your QuickBooks financial data
- 📱 **User-friendly Interface:** Easy-to-use menu system within Google Sheets

## Screenshots

<img width="1725" alt="Screenshot 2025-03-08 at 03 06 50" src="https://github.com/user-attachments/assets/d7267f7d-59be-4f11-8f55-e06761c314ee" />
<img width="1722" alt="Screenshot 2025-03-08 at 02 58 39" src="https://github.com/user-attachments/assets/93f9838b-0d27-49e8-8a55-c5e9d734c322" />
<img width="1715" alt="Screenshot 2025-03-08 at 02 59 26" src="https://github.com/user-attachments/assets/01c6422d-2493-45b6-8a72-92c5217f8b4d" />
<img width="1716" alt="Screenshot 2025-03-08 at 03 00 02" src="https://github.com/user-attachments/assets/96acb2ab-26ba-4902-81ce-4b2eedfe0022" />
<img width="1711" alt="Screenshot 2025-03-08 at 03 00 49" src="https://github.com/user-attachments/assets/92ed7d09-de0d-4020-a8b1-c290d855d4ca" />


## Installation

### Option 1: Use the Sample Sheet (Recommended)

1. Open the [Sample Sheet](https://docs.google.com/spreadsheets/d/1r2QB5vk8tI2tC2yiaYJSKqd_AmQGdCl-e7fEQRm8ods/edit?usp=sharing)
2. Make a copy to your Google Drive
3. The script is already embedded and ready to use

### Option 2: Manual Installation

1. Create a new Google Sheet or open an existing one
2. Open Script Editor (Extensions > Apps Script)
3. Copy each file from this repository into your Apps Script project
4. Set up required script properties (see below)
5. Save and refresh your spreadsheet

## Required Script Properties

Set these properties in the Apps Script project settings (Project Settings > Script Properties):

| Property | Description | Example |
|----------|-------------|---------|
| DEFAULT_CURRENCY | Default currency for formatting | `USD` |
| DEFAULT_LOCALE | Default locale for formatting | `en-US` |
| ENABLE_AI | Enable AI features | `true` |
| GEMINI_API_KEY | Your Google Gemini API key | `abc123...` |
| IS_INITIALIZED | Indicates if the script is initialized | `true` |
| QB_CLIENT_ID | QuickBooks API Client ID (optional) | `0` |
| QB_CLIENT_SECRET | QuickBooks API Client Secret (optional) | `#VALUE#` |
| QUICKBOOKS_ENV | QuickBooks environment (optional) | `SANDBOX` |
| USER_CONFIG | User configuration settings | `{"reportConfig":{"includeTableOfContents":"fa` |
| WORKSPACE_TOOLS | Workspace tools configuration | `#VALUE#` |
| YOUR_PROJECT_LOCATION | Your project location | `#VALUE#` |

## Getting Started

1. After installation, refresh your Google Sheet
2. You'll see a new "Gemini Financial AI" menu in the top menu bar
3. Start by running "Analyze Sheet" on your financial data
4. Explore the other menu options to access all features

## Demos and Tutorials

Coming Soon!

## Usage Guide

### Menu Options

- **Analyze Sheet**: Run anomaly detection on current sheet
- **Open Chat Assistant**: Launch AI chat interface for natural language queries
- **Reports**:
  - Generate Standard Report: Create a comprehensive analysis report
  - Generate Executive Summary: Create a concise summary for decision makers
  - Email Report: Send reports directly to stakeholders *NOTE: Requires Gmail authorisation*
- **Data Analysis**:
  - Analyze Selected Data: Focus analysis on selected cells only
  - Monthly Comparison: Compare financial data month-by-month
  - Transaction Pattern Analysis: Identify patterns in transactions
- **Integrations**:
  - QuickBooks Configuration: Set up QuickBooks API access
  - Import From QuickBooks: Import financial data from QuickBooks
- **Settings**:
  - Configuration: Set locale, currency and general preferences
  - Gemini AI Settings: Configure AI model behavior
  - Set Gemini Models: Select specific Gemini AI models to use, pulls from the latest Gemini models.

### Analysis Types

The application provides several types of financial analysis:

1. **Anomaly Detection**
   - Missing data detection
   - Invalid value detection
   - Duplicate entry detection
   - Statistical outlier detection (Z-score and IQR methods)
   - Pattern-based anomalies
   - AI-enhanced contextual analysis

2. **Pattern Analysis**
   - Frequency analysis: Identify recurring transaction patterns
   - Temporal patterns: Discover time-based trends (daily, weekly, monthly)
   - Value distribution: Analyze transaction amount distributions
   - Category analysis: Understand spending across categories

3. **Report Generation**
   - Standard detailed reports
   - Executive summaries for management
   - Custom reports with configurable sections
   - Visual charts and metrics
   - AI-generated insights and recommendations

## Configuration Options

Configure the application through the Settings menu:

1. **General Settings**
   - Default locale and currency
   - AI feature toggles
   - Data validation parameters

2. **Gemini AI Settings**
   - Model selection (switch between different Gemini models)
   - Analysis parameters
   - Response formatting preferences

3. **QuickBooks Integration**
   - API credentials configuration
   - Environment selection
   - Data synchronization options

## Project Structure

```
GeminiFinancialAnalysis/
├── AnomalyDetection.gs
├── Code.gs              # Main application code and menu handling
├── Configuration.gs     # Configuration management and preferences
├── Gemini.gs
├── HTMLUtils.gs
├── LICENSE
├── Quickbooks.gs
├── README.md            # Documentation
├── ReportGeneration.gs  # Report creation and formatting logic
├── ReportUtils.gs
├── Style.html           # CSS styles for UI components
├── UI_ConfigDialog.html
├── UI_EmailReportDialog.html
├── UI_GeminiModelSelection.html # Gemini model selection UI
├── UI_Main.html         # Main UI template for chat interface
├── UI_PatternAnalysisDialog.html
├── UI_QuickbooksConfig.html # QuickBooks integration UI
├── UI_QuickBooksImportDialog.html
├── UI_ReportMenu.html   # Report generation menu interface
└── utils.gs             # Utility functions and helpers
```

## Contributing

This is an open source project and contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Commit your changes (`git commit -m 'Add some amazing feature'`)
5. Push to the branch (`git push origin feature/amazing-feature`)
6. Open a Pull Request

Please make sure to update tests and documentation as appropriate.

## Development Guidelines

### Adding New Features

1.  Implement new functionalities by creating well-documented functions within the relevant `.gs` files. Ensure each function has a clear purpose and adheres to existing coding standards.
2.  Modify the `Code.gs` file to integrate new features into the application's menu structure. This ensures users can easily access and utilize the added functionalities through the Google Sheets interface.
3.  Develop or update UI components using HTML files to provide user-friendly interfaces for new features. Ensure that the UI elements are responsive and consistent with the application's overall design.
4.  Incorporate new configuration settings within the `Configuration.gs` file to allow users to customize and control the behavior of new features.
5.  Add or modify utility functions in `utils.gs` to support the implementation of new features, ensuring that these functions are reusable and efficient.
6.  Provide comprehensive documentation for all new features, including usage instructions, configuration options, and troubleshooting tips. Update the `README.md` file to reflect these changes.

### Best Practices

-   Employ TypeScript-style JSDoc comments to ensure comprehensive and maintainable code documentation.
-   Adhere to Google Apps Script best practices to optimize performance and reliability.
-   Implement robust error handling mechanisms to provide informative feedback to users and prevent unexpected issues.
-   Log significant operations and events to facilitate debugging and monitoring of the application's behavior.
-   Maintain consistent code formatting and naming conventions to enhance code readability and collaboration.
-   Thoroughly test new features and modifications on various datasets to ensure compatibility and accuracy.
  
## Dependencies

- Google Apps Script
- Google Sheets API
- Google Gemini AI API
- Google Charts Service (for visualizations)
- Google Drive API (for report generation)
- QuickBooks API (optional)

## Version History

- 1.1.0: Added pattern analysis and QuickBooks integration
- 1.0.0: Initial release with core anomaly detection and reporting

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support and Contact

- Create an issue in the GitHub repository
- Contact on LinkedIn via https://www.linkedin.com/in/stracos .

## Acknowledgments

- Google Gemini AI team for providing the latest up to date models and a free API tier for usage.
- GitHub user 'mhawksey' for providing excellent reference repo guidance https://github.com/mhawksey/GeminiApp .
