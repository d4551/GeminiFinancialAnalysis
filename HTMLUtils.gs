/**
 * Comprehensive utilities for HTML generation and handling in Google Apps Script
 */

/**
 * Correctly includes and evaluates an HTML file with proper template processing.
 * @param {string} filename The name of the file to include without the extension
 * @returns {string} The evaluated content of the HTML file
 */
function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename)
      .evaluate()
      .getContent();
  } catch (error) {
    logError(`Error including file "${filename}": ${error.message}`);
    return `<!-- Error loading ${filename}: ${error.message} -->`;
  }
}

/**
 * Builds a template from an HTML file and injects CSS from Style.html.
 * @param {string} filename The file to include
 * @returns {GoogleAppsScript.HTML.HtmlTemplate} A processed template
 */
function buildStyledTemplateFromFile(filename) {
  try {
    const template = HtmlService.createTemplateFromFile(filename);
    // Add the shared style
    template.styles = include('Style');
    // Add standard script utilities if they exist
    try {
      template.scriptUtils = include('ScriptUtils');
    } catch (e) {
      // ScriptUtils is optional
    }
    return template;
  } catch (error) {
    logError(`Error building template from "${filename}": ${error.message}`);
    throw error;
  }
}

/**
 * Creates a standard HTML page with common elements.
 * @param {string} title Page title
 * @param {string} content Main content HTML
 * @param {Object} options Additional options
 * @param {boolean} [options.includeSidebar=false] Whether to include sidebar-specific styling
 * @param {boolean} [options.includeUtils=true] Whether to include script utilities
 * @param {Object} [options.metadata={}] Additional metadata key-value pairs
 * @returns {string} The complete HTML content
 */
function createStandardPage(title, content, options = {}) {
  const {
    includeSidebar = false,
    includeUtils = true,
    metadata = {}
  } = options;
  
  // Build metadata tags
  const metaTags = Object.entries(metadata)
    .map(([name, content]) => `<meta name="${sanitizeAttribute(name)}" content="${sanitizeAttribute(content)}">`)
    .join('\n  ');

  // Add viewport meta tag if not provided
  const viewportMeta = metadata.viewport ? '' : 
    '<meta name="viewport" content="width=device-width, initial-scale=1">';
  
  try {  
    const styleContent = include("Style");
    let scriptUtils = '';
    
    if (includeUtils) {
      try {
        scriptUtils = `<script>\n${include("ScriptUtils")}\n</script>`;
      } catch (e) {
        // Script utils are optional
      }
    }
    
    return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>${styleContent}</style>
  ${viewportMeta}
  ${metaTags}
  <title>${sanitizeHTML(title)}</title>
  ${includeSidebar ? '<style>.container { max-width: 100%; }</style>' : ''}
  ${scriptUtils}
</head>
<body>
  <div class="container">
    <h3>${sanitizeHTML(title)}</h3>
    ${content}
  </div>
</body>
</html>`;
  } catch (error) {
    logError(`Error creating standard page: ${error.message}`);
    return `<html><body><h3>Error creating page</h3><p>${sanitizeHTML(error.message)}</p></body></html>`;
  }
}

/**
 * Creates a standard dialog with improved styling.
 * @param {string} title Dialog title
 * @param {string} content Main content HTML
 * @param {number} width Dialog width in pixels
 * @param {number} height Dialog height in pixels
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output for the dialog
 */
function createStandardDialog(title, content, width = 400, height = 300) {
  try {
    const htmlContent = createStandardPage(title, content);
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setWidth(width)
      .setHeight(height)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    return htmlOutput;
  } catch (error) {
    logError(`Error creating standard dialog: ${error.message}`);
    return HtmlService.createHtmlOutput(`<p>Error creating dialog: ${sanitizeHTML(error.message)}</p>`)
      .setWidth(width)
      .setHeight(height);
  }
}

/**
 * Creates a standard sidebar with optimized UI.
 * @param {string} title Sidebar title
 * @param {string} content Main content HTML
 * @param {number} width Optional width (defaults to 300)
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output for the sidebar
 */
function createStandardSidebar(title, content, width = 300) {
  try {
    const htmlContent = createStandardPage(title, content, { includeSidebar: true });
    const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
      .setTitle(title)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    if (width) {
      htmlOutput.setWidth(width);
    }
    
    return htmlOutput;
  } catch (error) {
    logError(`Error creating standard sidebar: ${error.message}`);
    return HtmlService.createHtmlOutput(`<p>Error creating sidebar: ${sanitizeHTML(error.message)}</p>`)
      .setTitle('Error')
      .setWidth(width);
  }
}

/**
 * Creates a template-based HTML dialog with full evaluation and styling.
 * @param {string} filename The HTML template filename without extension
 * @param {Object} templateVars Variables to pass to the template
 * @param {Object} dialogOptions Dialog configuration (width, height, title)
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The evaluated HTML dialog
 */
function createTemplatedDialog(filename, templateVars = {}, dialogOptions = {}) {
  try {
    const { width = 400, height = 350, title = filename } = dialogOptions;
    
    // Create and prepare the template with variables
    const template = HtmlService.createTemplateFromFile(filename);
    
    // Add each template variable
    Object.entries(templateVars).forEach(([key, value]) => {
      template[key] = value;
    });
    
    // Always include Style.html content
    template.styles = include('Style');
    
    const htmlOutput = template.evaluate()
      .setWidth(width)
      .setHeight(height)
      .setTitle(title)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    return htmlOutput;
  } catch (error) {
    logError(`Error creating templated dialog from "${filename}": ${error.message}`);
    return createStandardDialog(
      'Error',
      `<p class="error">Failed to load dialog template: ${sanitizeHTML(error.message)}</p>`,
      400,
      200
    );
  }
}

/**
 * Sanitizes a string for safe HTML display
 * @param {string} str The string to sanitize
 * @returns {string} The sanitized string
 */
function sanitizeHTML(str) {
  if (str === undefined || str === null) return '';
  return String(str).replace(/[&<>"']/g, function(c) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[c];
  });
}

/**
 * Sanitizes attribute values for HTML attributes
 * @param {string} value The attribute value to sanitize
 * @returns {string} Sanitized attribute value
 */
function sanitizeAttribute(value) {
  if (value === undefined || value === null) return '';
  return String(value).replace(/["'&<>]/g, function(c) {
    return {
      '"': '&quot;',
      "'": '&#39;',
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;'
    }[c];
  });
}

/**
 * Validates if a string is a properly formatted Google Document URL
 * @param {string} url The URL to validate
 * @returns {boolean} Whether the URL is valid
 */
function isValidDocUrl(url) {
  if (!url) return false;
  
  // More robust pattern matching for Google Docs URLs
  const docPattern = /^https:\/\/docs\.google\.com\/document\/d\/([-\w]{25,})\/(edit|preview|copy|view)(\?.*)?$/;
  
  // Check if the URL matches the pattern
  if (docPattern.test(url)) {
    return true;
  }
  
  // If not a standard URL format, check if it might be just a document ID
  const idOnlyPattern = /^([-\w]{25,})$/;
  if (idOnlyPattern.test(url)) {
    try {
      // Try to access the document to confirm it's real
      DocumentApp.openById(url);
      return true;
    } catch (e) {
      return false;
    }
  }
  
  return false;
}

/**
 * Creates a card-style section for UI elements
 * @param {string} title Section title
 * @param {string} content HTML content
 * @param {Object} options Card styling options
 * @returns {string} HTML for the card
 */
function createCard(title, content, options = {}) {
  const { 
    cssClass = '',
    titleTag = 'h3',
    includeIcon = false,
    icon = 'description'
  } = options;
  
  const iconHtml = includeIcon ? 
    `<span class="material-icon">${sanitizeHTML(icon)}</span>` : '';
  
  return `
    <div class="card ${cssClass}">
      <${titleTag}>${iconHtml}${sanitizeHTML(title)}</${titleTag}>
      <div class="card-content">
        ${content}
      </div>
    </div>
  `;
}

/**
 * Creates a button HTML element with improved styling
 * @param {string} label Button label
 * @param {string} onclick JavaScript onclick code
 * @param {Object} options Button configuration options
 * @returns {string} HTML for the button
 */
function createButton(label, onclick, options = {}) {
  const { 
    type = 'primary',
    disabled = false,
    cssClass = '',
    title = '',
    id = ''
  } = options;
  
  const classes = [type !== 'primary' ? type : '', cssClass].filter(Boolean).join(' ');
  const classAttr = classes ? ` class="${classes}"` : '';
  const disabledAttr = disabled ? ' disabled="disabled"' : '';
  const titleAttr = title ? ` title="${sanitizeAttribute(title)}"` : '';
  const idAttr = id ? ` id="${sanitizeAttribute(id)}"` : '';
  
  return `<button${classAttr}${disabledAttr}${titleAttr}${idAttr} onclick="${sanitizeAttribute(onclick)}">${sanitizeHTML(label)}</button>`;
}

/**
 * Creates a full HTML form with labels and inputs
 * @param {Array<Object>} fields Array of field definitions
 * @param {string} submitButtonLabel Submit button label
 * @param {string} submitFunction JavaScript function to call on submit
 * @returns {string} Complete form HTML
 */
function createForm(fields, submitButtonLabel, submitFunction) {
  const formFields = fields.map(field => {
    const {
      type = 'text',
      name,
      label,
      value = '',
      placeholder = '',
      required = false,
      options = []
    } = field;
    
    const nameAttr = ` name="${sanitizeAttribute(name)}" id="${sanitizeAttribute(name)}"`;
    const requiredAttr = required ? ' required="required"' : '';
    const placeholderAttr = placeholder ? ` placeholder="${sanitizeAttribute(placeholder)}"` : '';
    const valueAttr = value ? ` value="${sanitizeAttribute(value)}"` : '';
    
    let inputHtml;
    if (type === 'select') {
      const optionsHtml = options.map(opt => {
        const optValue = typeof opt === 'object' ? opt.value : opt;
        const optLabel = typeof opt === 'object' ? opt.label : opt;
        const selected = optValue === value ? ' selected="selected"' : '';
        return `<option value="${sanitizeAttribute(optValue)}"${selected}>${sanitizeHTML(optLabel)}</option>`;
      }).join('\n');
      
      inputHtml = `<select${nameAttr}${requiredAttr}>${optionsHtml}</select>`;
    } else if (type === 'textarea') {
      inputHtml = `<textarea${nameAttr}${requiredAttr}${placeholderAttr} rows="4">${sanitizeHTML(value)}</textarea>`;
    } else if (type === 'checkbox') {
      const checkedAttr = value ? ' checked="checked"' : '';
      inputHtml = `
        <label class="checkbox-label">
          <input type="${type}"${nameAttr}${requiredAttr}${checkedAttr}>
          ${sanitizeHTML(label)}
        </label>
      `;
      // Return early as we've already included the label
      return `<div class="form-group checkbox-group">${inputHtml}</div>`;
    } else {
      inputHtml = `<input type="${type}"${nameAttr}${requiredAttr}${placeholderAttr}${valueAttr}>`;
    }
    
    return `
      <div class="form-group">
        <label for="${sanitizeAttribute(name)}">${sanitizeHTML(label)}</label>
        ${inputHtml}
      </div>
    `;
  }).join('\n');
  
  return `
    <form id="dynamicForm" onsubmit="event.preventDefault(); ${submitFunction}">
      ${formFields}
      <div class="form-actions">
        <button type="submit" class="primary">${sanitizeHTML(submitButtonLabel)}</button>
      </div>
    </form>
  `;
}

/**
 * Creates styled HTML output from a file with error handling.
 * @param {string} filename The name of the HTML file to include without the extension
 * @param {Object} [templateVars={}] Optional variables to pass to the template
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output with styles applied
 */
function createStyledHtmlOutput(filename, templateVars = {}) {
  try {
    const template = HtmlService.createTemplateFromFile(filename);
    
    // Add template variables
    Object.entries(templateVars).forEach(([key, value]) => {
      template[key] = value;
    });
    
    // Add styles
    template.styles = include('Style');
    
    return template
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    logError(`Error creating styled HTML from "${filename}": ${error.message}`);
    return HtmlService.createHtmlOutput(
      `<div class="error">Failed to load UI: ${sanitizeHTML(error.message)}</div>`
    );
  }
}

/**
 * Creates a notification toast message that displays temporarily
 * @param {string} message The message to display
 * @param {string} type The notification type (success, error, warning, info)
 * @returns {string} HTML for the notification
 */
function createNotification(message, type = 'info') {
  const id = `notification-${new Date().getTime()}`;
  return `
    <div id="${id}" class="notification ${sanitizeAttribute(type)}">
      ${sanitizeHTML(message)}
      <span class="notification-close" onclick="this.parentElement.style.display='none';">&times;</span>
      <script>
        setTimeout(function() {
          var element = document.getElementById('${id}');
          if (element) {
            element.style.opacity = '0';
            setTimeout(function() { element.style.display = 'none'; }, 500);
          }
        }, 5000);
      </script>
    </div>
  `;
}

/**
 * Utility functions for HTML templates
 */

/**
 * Generate hour options for dropdowns
 * @returns {string} HTML option tags for hours
 */
function generateHourOptions() {
  let options = '';
  for (let i = 0; i < 24; i++) {
    const hourLabel = i < 10 ? '0' + i + ':00' : i + ':00';
    const selected = i === 6 ? 'selected' : ''; // Default to 6 AM
    options += `<option value="${i}" ${selected}>${hourLabel}</option>`;
  }
  return options;
}

/**
 * Include another HTML file in a template
 * @param {string} filename The name of the file to include
 * @returns {string} The contents of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Creates a templated dialog with consistent styling
 * @param {string} templateName The name of the HTML template file
 * @param {Object} data Data to pass to the template
 * @param {Object} options Options for the dialog
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output
 */
function createTemplatedDialog(templateName, data = {}, options = {}) {
  try {
    // Create a template with our data
    const template = HtmlService.createTemplateFromFile(templateName);
    
    // Add all data properties to the template
    Object.keys(data).forEach(key => {
      template[key] = data[key];
    });
    
    // Evaluate the template
    let htmlOutput = template.evaluate();
    
    // Set width and height if provided
    if (options.width) {
      htmlOutput.setWidth(options.width);
    }
    
    if (options.height) {
      htmlOutput.setHeight(options.height);
    }
    
    // Set title if provided
    if (options.title) {
      htmlOutput.setTitle(options.title);
    }
    
    return htmlOutput;
  } catch (error) {
    logError(`Error creating dialog from template ${templateName}: ${error.message}`);
    
    // Create a simple error dialog instead
    return HtmlService.createHtmlOutput(`
      <div style="color: red; padding: 20px;">
        <h3>Error</h3>
        <p>Failed to create dialog: ${sanitizeHTML(error.message)}</p>
      </div>
    `);
  }
}

/**
 * Creates a standard dialog with title and content
 * @param {string} title The dialog title
 * @param {string} content The HTML content
 * @param {number} width Width in pixels
 * @param {number} height Height in pixels
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output
 */
function createStandardDialog(title, content, width = 400, height = 300) {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body {
          font-family: Arial, sans-serif;
          margin: 0;
          padding: 20px;
          color: #333;
          line-height: 1.4;
        }
        h1, h2, h3 {
          color: #4285f4;
        }
        a {
          color: #4285f4;
          text-decoration: none;
        }
        a:hover {
          text-decoration: underline;
        }
      </style>
    </head>
    <body>
      ${content}
    </body>
    </html>
  `;
  
  return HtmlService.createHtmlOutput(htmlContent)
    .setWidth(width)
    .setHeight(height)
    .setTitle(title);
}

/**
 * Sanitizes HTML to prevent XSS attacks
 * @param {string} input The input to sanitize
 * @returns {string} Sanitized output
 */
function sanitizeHTML(input) {
  if (!input) return '';
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Gets sheet data for analysis
 * @returns {Array<Array<any>>} The sheet data
 */
function getSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getDataRange().getValues();
}

/**
 * Validates and normalizes a document URL to ensure it's properly formatted
 * @param {string} url The document URL to validate
 * @returns {string|null} Normalized URL or null if invalid
 */
function validateAndNormalizeDocUrl(url) {
  if (!url) return null;
  
  // Check for valid Google Docs URL format
  const docIdRegex = /[-\w]{25,}/;
  const match = url.match(docIdRegex);
  
  if (!match) return null;
  
  const docId = match[0];
  return `https://docs.google.com/document/d/${docId}/edit`;
}

/**
 * Appends content to a Google Doc report
 * @param {string} docUrl The URL of the document
 * @param {string} content The content to append
 * @returns {boolean} True if successful, false otherwise
 */
function appendToReport(docUrl, content) {
  try {
    // Extract the document ID from the URL
    const idMatch = docUrl.match(/[-\w]{25,}/);
    if (!idMatch) {
      return false;
    }
    
    const doc = DocumentApp.openById(idMatch[0]);
    const body = doc.getBody();
    
    // Add a section header for appended content
    body.appendParagraph("Additional Insights")
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    // Split content into paragraphs and append each
    const paragraphs = content.split("\n\n");
    paragraphs.forEach(paragraph => {
      if (paragraph.trim()) {
        body.appendParagraph(paragraph);
      }
    });
    
    // Save the changes
    doc.saveAndClose();
    return true;
  } catch (error) {
    logError(`Error appending to report: ${error.message}`);
    return false;
  }
}

/**
 * Helper method for logging messages
 * @param {string} message The message to log
 * @param {string} context Optional context information
 */
function logMessage(message, context = '') {
  if (context) {
    console.log(`${context}: ${message}`);
  } else {
    console.log(message);
  }
}
