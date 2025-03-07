/**
 * Advanced utilities for report generation and visualization
 * Supporting the enhanced reporting capabilities
 */

/**
 * Inserts a dynamic table of contents at the beginning of the document
 * @param {GoogleAppsScript.Document.Document} doc The document to insert the ToC into
 * @param {Object} options Optional configuration for the ToC
 */
function insertDynamicTableOfContents(doc, options = {}) {
  try {
    const {
      title = 'Table of Contents',
      heading = DocumentApp.ParagraphHeading.HEADING1,
      style = DocumentApp.TableOfContentsStyle.BLUE_LINKS
    } = options;
    
    const body = doc.getBody();
    
    // Insert heading and ToC at the beginning
    body.insertParagraph(0, title)
        .setHeading(heading);
    
    // Insert the actual ToC
    body.insertParagraph(1, '')
        .appendTableOfContents(style);
    
    logMessage('Table of contents inserted successfully');
  } catch (error) {
    logError(`Failed to insert table of contents: ${error.message}`, 'insertDynamicTableOfContents');
    // Continue without ToC rather than failing the whole report
  }
}

/**
 * Creates an interactive chart with enhanced visualizations
 * @param {Array<Object>} anomalies Anomalies to visualize
 * @param {string} chartType Type of chart to create ('bar', 'pie', 'column')
 * @param {Object} options Chart configuration options
 * @returns {Blob} The chart as an image blob
 */
function createInteractiveChart(anomalies, chartType = 'bar', options = {}) {
  try {
    // Create temporary spreadsheet for chart generation
    const ss = SpreadsheetApp.create('TempChartData_' + new Date().getTime());
    const sheet = ss.getActiveSheet();
    
    // Get data based on analysis type
    let chartData;
    switch (options.dataType || 'errors') {
      case 'errors':
        chartData = breakdownErrors(anomalies);
        sheet.appendRow(['Error Type', 'Frequency']);
        break;
      case 'categories':
        chartData = breakdownCategories(anomalies);
        sheet.appendRow(['Category', 'Count']);
        break;
      case 'confidence':
        chartData = {
          'High Confidence': anomalies.filter(a => (a.confidence || 1.0) > 0.8).length,
          'Medium Confidence': anomalies.filter(a => {
            const conf = a.confidence || 1.0;
            return conf <= 0.8 && conf > 0.5;
          }).length,
          'Low Confidence': anomalies.filter(a => (a.confidence || 1.0) <= 0.5).length
        };
        sheet.appendRow(['Confidence Level', 'Count']);
        break;
      default:
        throw new Error(`Unsupported data type: ${options.dataType}`);
    }
    
    // Populate data rows
    Object.entries(chartData).forEach(([key, value]) => {
      sheet.appendRow([key, value]);
    });
    
    // Create chart in sheets with the appropriate type
    let chartBuilder;
    
    switch (chartType.toLowerCase()) {
      case 'pie':
        chartBuilder = sheet.newChart()
          .setChartType(Charts.ChartType.PIE)
          .addRange(sheet.getRange(1, 1, sheet.getLastRow(), 2));
        break;
      case 'column':
        chartBuilder = sheet.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(sheet.getRange(1, 1, sheet.getLastRow(), 2));
        break;
      case 'bar':
      default:
        chartBuilder = sheet.newChart()
          .setChartType(Charts.ChartType.BAR)
          .addRange(sheet.getRange(1, 1, sheet.getLastRow(), 2));
        break;
    }
    
    // Add chart options
    chartBuilder
      .setOption('title', options.title || `${options.dataType || 'Error'} Distribution`)
      .setOption('legend', { position: options.legendPosition || 'right' })
      .setOption('width', options.width || 600)
      .setOption('height', options.height || 400)
      .setOption('colors', options.colors || ['#4285F4', '#34A853', '#FBBC05', '#EA4335'])
      .setPosition(5, 5, 0, 0);
    
    // Build and insert the chart
    sheet.insertChart(chartBuilder.build());
    
    // Export chart as image blob
    const charts = sheet.getCharts();
    const chartBlob = charts[0].getAs('image/png');
    
    // Clean up the temporary spreadsheet
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    
    return chartBlob;
  } catch (error) {
    logError(`Failed to create interactive chart: ${error.message}`, 'createInteractiveChart');
    // Return null instead of throwing, so report creation can continue
    return null;
  }
}

/**
 * Creates a log sheet for capturing detailed application logs
 * @param {string} sheetName Name of the log sheet
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The created or existing log sheet
 */
function setupLoggingSheet(sheetName = 'SystemLogs') {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    let logSpreadsheetId = scriptProperties.getProperty('LOG_SPREADSHEET_ID');
    let spreadsheet;
    
    // Create new log spreadsheet if none exists
    if (!logSpreadsheetId) {
      spreadsheet = SpreadsheetApp.create('GeminiFinancialAI Logs');
      logSpreadsheetId = spreadsheet.getId();
      scriptProperties.setProperty('LOG_SPREADSHEET_ID', logSpreadsheetId);
    } else {
      try {
        spreadsheet = SpreadsheetApp.openById(logSpreadsheetId);
      } catch (e) {
        // If the spreadsheet was deleted, create a new one
        spreadsheet = SpreadsheetApp.create('GeminiFinancialAI Logs');
        logSpreadsheetId = spreadsheet.getId();
        scriptProperties.setProperty('LOG_SPREADSHEET_ID', logSpreadsheetId);
      }
    }
    
    // Get or create the specific log sheet
    let logSheet;
    try {
      logSheet = spreadsheet.getSheetByName(sheetName);
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(sheetName);
        // Add headers to the new sheet
        logSheet.appendRow(['Timestamp', 'Level', 'Context', 'Message']);
        // Format headers
        logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
      }
    } catch (e) {
      // If there's an issue with the sheet, create a new one
      logSheet = spreadsheet.insertSheet(sheetName);
      logSheet.appendRow(['Timestamp', 'Level', 'Context', 'Message']);
      logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
    }
    
    return logSheet;
  } catch (error) {
    console.error(`Failed to setup logging sheet: ${error.message}`);
    // Return null instead of throwing to prevent cascading failures
    return null;
  }
}

/**
 * Appends a log entry to the logging sheet
 * @param {string} level Log level (INFO, WARNING, ERROR, DEBUG)
 * @param {string} message The message to log
 * @param {string} context Additional context information
 */
function appendToLogSheet(level, message, context = '') {
  try {
    const logSheet = setupLoggingSheet();
    if (!logSheet) return;
    
    const timestamp = new Date();
    logSheet.appendRow([timestamp, level.toUpperCase(), context, message]);
    
    // Apply conditional formatting based on log level
    const lastRow = logSheet.getLastRow();
    const cell = logSheet.getRange(lastRow, 2); // Level column
    
    switch (level.toUpperCase()) {
      case 'ERROR':
        cell.setBackground('#F4C7C3'); // Light red
        break;
      case 'WARNING':
        cell.setBackground('#FCE8B2'); // Light yellow
        break;
      case 'INFO':
        cell.setBackground('#CFE8FC'); // Light blue
        break;
    }
    
    // Keep the log sheet to a reasonable size (e.g., last 1000 entries)
    const maxRows = 1000;
    if (lastRow > maxRows + 1) { // +1 for header
      logSheet.deleteRows(2, lastRow - maxRows - 1);
    }
  } catch (error) {
    // Just log to console if sheet logging fails
    console.error(`Log sheet error: ${error.message}`);
  }
}

/**
 * Creates an advanced visualization for anomaly patterns
 * @param {Array<Object>} anomalies The anomalies to visualize
 * @param {string} visualizationType The type of visualization to create
 * @param {Object} options Additional visualization options
 * @returns {Blob} Image blob of the visualization
 */
function createAdvancedVisualization(anomalies, visualizationType, options = {}) {
  try {
    if (!anomalies || anomalies.length === 0) {
      return null;
    }
    
    switch (visualizationType.toLowerCase()) {
      case 'heatmap':
        return createHeatmapVisualization(anomalies, options);
      case 'timeline':
        return createTimelineVisualization(anomalies, options);
      case 'network':
        return createNetworkVisualization(anomalies, options);
      case 'distribution':
        return createDistributionVisualization(anomalies, options);
      default:
        // Default to bar chart
        return createInteractiveChart(anomalies, 'bar', options);
    }
  } catch (error) {
    logError(`Failed to create advanced visualization: ${error.message}`, 'createAdvancedVisualization');
    return null;
  }
}

/**
 * Creates a heatmap visualization of anomalies by date and category
 * @param {Array<Object>} anomalies The anomalies to visualize
 * @param {Object} options Visualization options
 * @returns {Blob} Image blob of the heatmap
 */
function createHeatmapVisualization(anomalies, options = {}) {
  try {
    // Create a temporary spreadsheet for the heatmap
    const ss = SpreadsheetApp.create('TempHeatmap_' + new Date().getTime());
    const sheet = ss.getActiveSheet();
    
    // Group anomalies by date and category
    const dateCategories = {};
    
    // Extract all unique dates and categories
    const dates = new Set();
    const categories = new Set();
    
    anomalies.forEach(anomaly => {
      if (anomaly.date) {
        let dateStr;
        if (anomaly.date instanceof Date) {
          dateStr = anomaly.date.toISOString().split('T')[0];
        } else {
          try {
            // Try to parse as date
            const date = new Date(anomaly.date);
            if (!isNaN(date.getTime())) {
              dateStr = date.toISOString().split('T')[0];
            } else {
              dateStr = String(anomaly.date);
            }
          } catch (e) {
            dateStr = String(anomaly.date);
          }
        }
        
        const category = anomaly.category || 'Uncategorized';
        dates.add(dateStr);
        categories.add(category);
        
        if (!dateCategories[dateStr]) {
          dateCategories[dateStr] = {};
        }
        
        if (!dateCategories[dateStr][category]) {
          dateCategories[dateStr][category] = 0;
        }
        
        dateCategories[dateStr][category]++;
      }
    });
    
    // Convert to arrays and sort
    const sortedDates = [...dates].sort();
    const sortedCategories = [...categories].sort();
    
    // Create the heatmap data
    // Headers: add empty cell for row headers, then dates
    const headers = [''];
    headers.push(...sortedDates);
    sheet.appendRow(headers);
    
    // Add data rows
    sortedCategories.forEach(category => {
      const row = [category];
      
      sortedDates.forEach(date => {
        const count = dateCategories[date]?.[category] || 0;
        row.push(count);
      });
      
      sheet.appendRow(row);
    });
    
    // Apply conditional formatting for the heatmap effect
    const dataRange = sheet.getRange(2, 2, sortedCategories.length, sortedDates.length);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#FF4081', SpreadsheetApp.InterpolationType.NUMBER, '=MAX(B2:Z999)')
      .setGradientMidpointWithValue('#FFCDD2', SpreadsheetApp.InterpolationType.NUMBER, '=AVERAGE(B2:Z999)')
      .setGradientMinpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, 0)
      .setRanges([dataRange])
      .build();
      
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
    
    // Create a chart based on the heatmap data
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.TABLE)
      .addRange(sheet.getDataRange())
      .setPosition(sortedCategories.length + 5, 1, 0, 0)
      .setOption('allowHtml', true)
      .setOption('width', 600)
      .setOption('height', 400);
      
    sheet.insertChart(chartBuilder.build());
    
    // Export the visualization
    Utilities.sleep(1000); // Give time for the chart to render
    const charts = sheet.getCharts();
    const chartBlob = charts[0].getAs('image/png');
    
    // Clean up the temporary spreadsheet
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    
    return chartBlob;
  } catch (error) {
    logError(`Failed to create heatmap visualization: ${error.message}`, 'createHeatmapVisualization');
    return null;
  }
}

/**
 * Creates a timeline visualization of anomalies
 * @param {Array<Object>} anomalies The anomalies to visualize
 * @param {Object} options Visualization options
 * @returns {Blob} Image blob of the timeline
 */
function createTimelineVisualization(anomalies, options = {}) {
  try {
    const ss = SpreadsheetApp.create('TempTimeline_' + new Date().getTime());
    const sheet = ss.getActiveSheet();
    
    // Set up the headers
    sheet.appendRow(['Date', 'Amount', 'Category', 'Confidence']);
    
    // Filter anomalies to only those with dates
    const validAnomalies = anomalies.filter(a => a.date);
    
    // Sort by date
    validAnomalies.sort((a, b) => {
      const dateA = a.date instanceof Date ? a.date : new Date(a.date);
      const dateB = b.date instanceof Date ? b.date : new Date(b.date);
      return dateA - dateB;
    });
    
    // Add data
    validAnomalies.forEach(anomaly => {
      let dateValue;
      if (anomaly.date instanceof Date) {
        dateValue = anomaly.date;
      } else {
        try {
          dateValue = new Date(anomaly.date);
          if (isNaN(dateValue.getTime())) {
            dateValue = anomaly.date; // Use as is if not a valid date
          }
        } catch (e) {
          dateValue = anomaly.date;
        }
      }
      
      const amount = typeof anomaly.amount === 'number' ? 
        anomaly.amount : 
        (parseFloat(anomaly.amount) || 0);
      
      const category = anomaly.category || 'Uncategorized';
      const confidence = (anomaly.confidence || 1.0) * 100;
      
      sheet.appendRow([dateValue, amount, category, confidence]);
    });
    
    // Format the date column
    sheet.getRange(2, 1, validAnomalies.length, 1).setNumberFormat('yyyy-mm-dd');
    
    // Format the amount column
    sheet.getRange(2, 2, validAnomalies.length, 1).setNumberFormat('#,##0.00');
    
    // Create scatter chart for timeline visualization
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.SCATTER)
      .addRange(sheet.getRange(1, 1, validAnomalies.length + 1, 3))
      .setPosition(5, 5, 0, 0)
      .setOption('title', 'Anomaly Timeline')
      .setOption('hAxis.title', 'Date')
      .setOption('vAxis.title', 'Amount')
      .setOption('legend', { position: 'right' })
      .setOption('bubble.textStyle', { color: 'black' })
      .setOption('width', 800)
      .setOption('height', 500);
      
    sheet.insertChart(chartBuilder.build());
    
    // Wait for chart to render
    Utilities.sleep(1000);
    
    // Get the chart
    const charts = sheet.getCharts();
    const chartBlob = charts[0].getAs('image/png');
    
    // Clean up
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    
    return chartBlob;
  } catch (error) {
    logError(`Failed to create timeline visualization: ${error.message}`, 'createTimelineVisualization');
    return null;
  }
}

/**
 * Creates a network visualization showing relationships between anomalies
 * @param {Array<Object>} anomalies The anomalies to visualize
 * @param {Object} options Visualization options
 * @returns {Blob} Image blob of the network diagram
 */
function createNetworkVisualization(anomalies, options = {}) {
  try {
    const ss = SpreadsheetApp.create('TempNetwork_' + new Date().getTime());
    const sheet = ss.getActiveSheet();
    
    // Create sheets for nodes and edges
    const nodeSheet = ss.insertSheet('Nodes');
    const edgeSheet = ss.insertSheet('Edges');
    
    // Set up node headers
    nodeSheet.appendRow(['ID', 'Label', 'Category', 'Size', 'Color']);
    
    // Set up edge headers
    edgeSheet.appendRow(['Source', 'Target', 'Weight', 'Description']);
    
    // Create nodes for categories
    const categories = new Set();
    anomalies.forEach(a => categories.add(a.category || 'Uncategorized'));
    
    // Add category nodes
    let nodeId = 1;
    const categoryIds = {};
    categories.forEach(category => {
      const count = anomalies.filter(a => (a.category || 'Uncategorized') === category).length;
      nodeSheet.appendRow([nodeId, category, 'Category', 10 + (count * 2), '#4285F4']);
      categoryIds[category] = nodeId++;
    });
    
    // Add transaction description nodes
    const descriptions = new Map();
    anomalies.forEach(a => {
      const desc = a.description || 'No Description';
      if (!descriptions.has(desc)) {
        descriptions.set(desc, nodeId++);
        nodeSheet.appendRow([descriptions.get(desc), desc.substring(0, 30), 'Transaction', 5, '#34A853']);
      }
    });
    
    // Create edges between categories and descriptions
    anomalies.forEach(a => {
      const category = a.category || 'Uncategorized';
      const desc = a.description || 'No Description';
      
      if (categoryIds[category] && descriptions.has(desc)) {
        edgeSheet.appendRow([
          categoryIds[category], 
          descriptions.get(desc),
          1,
          `${category} → ${desc}`
        ]);
      }
    });
    
    // Create a network diagram representation using a scatter plot
    // (This is an approximation as Google Sheets doesn't have native network diagrams)
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.SCATTER)
      .addRange(nodeSheet.getRange(1, 1, nodeSheet.getLastRow(), 5))
      .setPosition(5, 5, 0, 0)
      .setOption('title', 'Anomaly Relationship Network')
      .setOption('width', 800)
      .setOption('height', 600)
      .setOption('bubble.textStyle', { color: 'black' })
      .setOption('hAxis.title', '')
      .setOption('vAxis.title', '');
      
    sheet.insertChart(chartBuilder.build());
    
    // Wait for chart to render
    Utilities.sleep(1000);
    
    // Get the chart
    const charts = sheet.getCharts();
    const chartBlob = charts[0].getAs('image/png');
    
    // Clean up
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    
    return chartBlob;
  } catch (error) {
    logError(`Failed to create network visualization: ${error.message}`, 'createNetworkVisualization');
    return null;
  }
}

/**
 * Creates a distribution visualization showing the distribution of anomaly values
 * @param {Array<Object>} anomalies The anomalies to visualize
 * @param {Object} options Visualization options
 * @returns {Blob} Image blob of the distribution chart
 */
function createDistributionVisualization(anomalies, options = {}) {
  try {
    const ss = SpreadsheetApp.create('TempDistribution_' + new Date().getTime());
    const sheet = ss.getActiveSheet();
    
    // Extract numeric amounts
    const amounts = anomalies
      .map(a => typeof a.amount === 'number' ? a.amount : parseFloat(a.amount))
      .filter(a => !isNaN(a));
      
    if (amounts.length < 2) {
      return null;
    }
    
    // Calculate histogram bins
    const min = Math.min(...amounts);
    const max = Math.max(...amounts);
    const range = max - min;
    
    // Determine number of bins (Sturges' formula: k = log2(n) + 1)
    const binCount = Math.max(5, Math.min(20, Math.ceil(Math.log2(amounts.length) + 1)));
    const binWidth = range / binCount;
    
    // Create bin boundaries
    const bins = [];
    for (let i = 0; i < binCount; i++) {
      const lowerBound = min + (i * binWidth);
      const upperBound = lowerBound + binWidth;
      bins.push({ 
        lowerBound, 
        upperBound, 
        label: `${lowerBound.toFixed(2)} - ${upperBound.toFixed(2)}`,
        count: 0 
      });
    }
    
    // Count values in each bin
    amounts.forEach(amount => {
      const binIndex = Math.min(
        binCount - 1, 
        Math.floor((amount - min) / binWidth)
      );
      if (bins[binIndex]) {
        bins[binIndex].count++;
      }
    });
    
    // Add headers
    sheet.appendRow(['Range', 'Count']);
    
    // Add bin data
    bins.forEach(bin => {
      sheet.appendRow([bin.label, bin.count]);
    });
    
    // Create a column chart for the distribution
    const chartBuilder = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getDataRange())
      .setPosition(5, 5, 0, 0)
      .setOption('title', 'Distribution of Anomaly Amounts')
      .setOption('hAxis.title', 'Amount Range')
      .setOption('vAxis.title', 'Frequency')
      .setOption('legend', { position: 'none' })
      .setOption('colors', ['#4285F4'])
      .setOption('width', 800)
      .setOption('height', 500);
      
    sheet.insertChart(chartBuilder.build());
    
    // Wait for chart to render
    Utilities.sleep(1000);
    
    // Get the chart
    const charts = sheet.getCharts();
    const chartBlob = charts[0].getAs('image/png');
    
    // Clean up
    DriveApp.getFileById(ss.getId()).setTrashed(true);
    
    return chartBlob;
  } catch (error) {
    logError(`Failed to create distribution visualization: ${error.message}`, 'createDistributionVisualization');
    return null;
  }
}

/**
 * Inserts a report chart into a Google Doc with proper formatting and caption
 * @param {GoogleAppsScript.Document.Document} doc The document to insert into
 * @param {Blob} chartBlob The chart image blob
 * @param {string} caption Optional caption for the chart
 * @returns {GoogleAppsScript.Document.InlineImage} The inserted image
 */
function insertReportChart(doc, chartBlob, caption = '') {
  if (!chartBlob) return null;
  
  try {
    const body = doc.getBody();
    const image = body.appendImage(chartBlob);
    
    // Set max width to fit the page
    const width = Math.min(image.getWidth(), 500);
    const ratio = width / image.getWidth();
    const height = Math.round(image.getHeight() * ratio);
    
    image.setWidth(width);
    image.setHeight(height);
    
    // Add caption if provided
    if (caption) {
      const captionParagraph = body.appendParagraph(caption);
      captionParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      captionParagraph.setItalic(true);
      captionParagraph.setFontSize(9);
    }
    
    return image;
  } catch (error) {
    logError(`Failed to insert chart: ${error.message}`, 'insertReportChart');
    return null;
  }
}

/**
 * Analyzes the distribution of data values and returns statistical insights
 * @param {Array<number>} values Array of numeric values to analyze
 * @returns {Object} Statistical analysis results
 */
function analyzeDistribution(values) {
  if (!values || values.length === 0) {
    return { count: 0 };
  }
  
  try {
    const validValues = values.filter(v => typeof v === 'number' && !isNaN(v));
    const count = validValues.length;
    
    if (count === 0) {
      return { count: 0 };
    }
    
    // Sort values for percentile calculations
    const sortedValues = [...validValues].sort((a, b) => a - b);
    
    // Basic statistics
    const sum = validValues.reduce((acc, val) => acc + val, 0);
    const mean = sum / count;
    const min = sortedValues[0];
    const max = sortedValues[count - 1];
    
    // Variance and standard deviation
    const variance = validValues.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / count;
    const stdDev = Math.sqrt(variance);
    
    // Quartiles and IQR
    const q1Index = Math.floor(count * 0.25);
    const q2Index = Math.floor(count * 0.5);
    const q3Index = Math.floor(count * 0.75);
    
    const q1 = sortedValues[q1Index];
    const median = sortedValues[q2Index];
    const q3 = sortedValues[q3Index];
    const iqr = q3 - q1;
    
    // Skewness (measure of asymmetry)
    const skewness = validValues.reduce((acc, val) => {
      return acc + Math.pow((val - mean) / stdDev, 3);
    }, 0) / count;
    
    // Kurtosis (measure of "tailedness")
    const kurtosis = validValues.reduce((acc, val) => {
      return acc + Math.pow((val - mean) / stdDev, 4);
    }, 0) / count - 3; // Excess kurtosis (normal distribution = 0)
    
    return {
      count,
      min,
      max,
      range: max - min,
      sum,
      mean,
      median,
      variance,
      stdDev,
      q1,
      q3,
      iqr,
      skewness,
      kurtosis,
      // Count of values in distribution tails (potential outliers)
      lowerOutliers: validValues.filter(v => v < q1 - 1.5 * iqr).length,
      upperOutliers: validValues.filter(v => v > q3 + 1.5 * iqr).length
    };
  } catch (error) {
    logError(`Error analyzing distribution: ${error.message}`, 'analyzeDistribution');
    return { 
      count: values.length,
      error: error.message
    };
  }
}

/**
 * Generates a report metadata section with version, date, and execution info
 * @param {GoogleAppsScript.Document.Document} doc The document to add metadata to
 * @param {Object} options Additional metadata options
 */
function addReportMetadata(doc, options = {}) {
  try {
    const body = doc.getBody();
    
    // Add a metadata section at the end of the document
    const metadataSectionTitle = body.appendParagraph('Report Metadata');
    metadataSectionTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    const metadataTable = body.appendTable([
      ['Generated On', new Date().toLocaleString()],
      ['Report Version', options.version || APP_VERSION || '1.0.0'],
      ['Generated By', options.author || Session.getActiveUser().getEmail() || 'System'],
      ['Execution Time', options.executionTime ? `${options.executionTime.toFixed(2)}s` : 'N/A'],
      ['Analysis Method', options.analysisMethod || 'Standard']
    ]);
    
    // Style the metadata table
    metadataTable.setAttributes({
      [DocumentApp.Attribute.BORDER_WIDTH]: 0,
      [DocumentApp.Attribute.FONT_SIZE]: 9
    });
    
    // First column bold
    const numRows = metadataTable.getNumRows();
    for (let i = 0; i < numRows; i++) {
      metadataTable.getCell(i, 0).setBold(true);
    }
    
    // Add notices and disclaimers if any
    if (options.notices?.length > 0) {
      const noticesParagraph = body.appendParagraph('\nNotices:');
      noticesParagraph.setFontSize(9).setBold(true);
      
      options.notices.forEach(notice => {
        body.appendParagraph(`• ${notice}`).setFontSize(9);
      });
    }
  } catch (error) {
    logError(`Error adding report metadata: ${error.message}`, 'addReportMetadata');
  }
}

/**
 * Extracts text summary from a particular section of a document
 * @param {GoogleAppsScript.Document.Document} doc The document to extract from
 * @param {string} sectionHeading The heading text to find
 * @param {number} maxLength Maximum number of characters to extract
 * @returns {string} Extracted summary text
 */
function extractSectionSummary(doc, sectionHeading, maxLength = 500) {
  try {
    const body = doc.getBody();
    const searchResult = body.findText(sectionHeading);
    
    if (!searchResult) {
      return '';
    }
    
    // Get the paragraph containing the heading
    const headingParagraph = searchResult.getElement().getParent();
    let currentElement = headingParagraph.getNextSibling();
    let summary = '';
    
    // Collect text until we hit another heading or reach maxLength
    while (currentElement && 
           summary.length < maxLength && 
           currentElement.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      
      // Check if we've hit another heading
      if (currentElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
        const paragraphHeading = currentElement.asParagraph().getHeading();
        if (paragraphHeading !== DocumentApp.ParagraphHeading.NORMAL) {
          break;
        }
      }
      
      // Add text content
      if (currentElement.getType() === DocumentApp.ElementType.PARAGRAPH) {
        summary += currentElement.asParagraph().getText() + ' ';
      } else if (currentElement.getType() === DocumentApp.ElementType.LIST_ITEM) {
        summary += '• ' + currentElement.asListItem().getText() + ' ';
      }
      
      currentElement = currentElement.getNextSibling();
    }
    
    // Truncate to maxLength and add ellipsis if needed
    if (summary.length > maxLength) {
      summary = summary.substring(0, maxLength - 3) + '...';
    }
    
    return summary.trim();
  } catch (error) {
    logError(`Error extracting section summary: ${error.message}`, 'extractSectionSummary');
    return '';
  }
}

/**
 * Creates a report executive summary with meaningful insights
 * @param {Array<Object>} anomalies The anomalies to analyze
 * @param {Object} data Additional data context
 * @returns {Promise<string>} Executive summary text
 */
async function generateExecutiveSummary(anomalies, data = null) {
  try {
    // Extract key information for the summary
    const totalAnomalies = anomalies.length;
    const highConfidence = anomalies.filter(a => (a.confidence || 1.0) > 0.8).length;
    const categories = {};
    const amounts = [];
    
    // Collect category and amount information
    anomalies.forEach(a => {
      if (a.category) {
        categories[a.category] = (categories[a.category] || 0) + 1;
      }
      
      if (a.amount !== undefined && a.amount !== null && a.amount !== 'N/A') {
        const amount = parseFloat(a.amount);
        if (!isNaN(amount)) {
          amounts.push(amount);
        }
      }
    });
    
    // Extract top categories
    const topCategories = Object.entries(categories)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 3)
      .map(([name, count]) => `${name} (${count})`);
      
    // Calculate financial impact
    let totalAmount = 0;
    if (amounts.length > 0) {
      totalAmount = amounts.reduce((sum, amt) => sum + amt, 0);
    }
    
    // Create prompt for Gemini
    const prompt = `Generate a concise executive summary (250-350 words) for a financial anomaly report with the following information:
    
    Total anomalies detected: ${totalAnomalies}
    High confidence issues: ${highConfidence}
    Top affected categories: ${topCategories.join(', ')}
    Financial impact: ${totalAmount}
    
    The summary should:
    1. Start with a professional opening statement about the analysis performed
    2. Highlight key findings and their business impact
    3. Summarize the risk profile of the detected anomalies
    4. Offer 2-3 specific recommendations for executives
    5. Include a brief conclusion
    
    Use professional but clear business language suitable for executives. Format with paragraph breaks.`;
    
    // Generate the summary using Gemini
    const summary = await generateReportAnalysis(prompt);
    return summary;
  } catch (error) {
    logError(`Error generating executive summary: ${error.message}`, 'generateExecutiveSummary');
    return "Error generating executive summary. Please check the detailed analysis section for more information.";
  }
}

/**
 * Generates a PDF version of a report document
 * @param {string} documentId The ID of the Google Doc to convert
 * @returns {string} URL to the PDF version
 */
function generateReportPDF(documentId) {
  try {
    const doc = DocumentApp.openById(documentId);
    const docFile = DriveApp.getFileById(documentId);
    
    // Save any pending changes
    doc.saveAndClose();
    
    // Use advanced Drive API to export as PDF
    const pdf = Drive.Files.export(documentId, 'application/pdf');
    
    // Create a blob from the PDF content
    const pdfBlob = Utilities.newBlob(pdf.getBytes(), 'application/pdf', docFile.getName() + '.pdf');
    
    // Create the PDF file in Drive
    const folder = docFile.getParents().next();
    const pdfFile = folder.createFile(pdfBlob);
    
    // Make it accessible to the user
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return pdfFile.getUrl();
  } catch (error) {
    logError(`Error generating PDF: ${error.message}`, 'generateReportPDF');
    throw new Error(`Failed to generate PDF: ${error.message}`);
  }
}

/**
 * Recursively cleans up temporary files used for report generation
 * @param {string} reportId The report ID to look for in filenames
 */
function cleanupTemporaryFiles(reportId) {
  try {
    if (!reportId) return;
    
    // Search for files related to this report ID
    const files = DriveApp.searchFiles(
      `title contains 'Temp' and title contains '${reportId}'`
    );
    
    // Delete all matching files
    while (files.hasNext()) {
      const file = files.next();
      try {
        file.setTrashed(true);
      } catch (fileError) {
        logError(`Error deleting temporary file ${file.getName()}: ${fileError.message}`, 'cleanupTemporaryFiles');
      }
    }
  } catch (error) {
    logError(`Error cleaning up temporary files: ${error.message}`, 'cleanupTemporaryFiles');
    // Continue without throwing - cleanup errors shouldn't affect report generation
  }
}

/**
 * Retrieves schedule configuration for anomaly detection
 * @returns {Object} The current schedule configuration
 */
function getScheduleConfig() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const frequency = scriptProperties.getProperty('ANOMALY_DETECTION_FREQUENCY');
    
    if (!frequency || frequency === 'none') {
      return { frequency: 'none' };
    }
    
    return {
      frequency,
      notificationEmail: scriptProperties.getProperty('ANOMALY_NOTIFICATION_EMAIL'),
      weekDay: scriptProperties.getProperty('ANOMALY_DETECTION_WEEKDAY'),
      hour: scriptProperties.getProperty('ANOMALY_DETECTION_HOUR')
    };
  } catch (error) {
    logError(`Error getting schedule config: ${error.message}`, 'getScheduleConfig');
    return { frequency: 'none' };
  }
}

/**
 * Validates a Google Docs document URL and extracts the document ID
 * @param {string} docUrl The URL to validate and parse
 * @returns {string|null} Document ID or null if invalid
 */
function extractDocumentId(docUrl) {
  if (!docUrl) return null;
  
  try {
    // Match both edit and view URLs
    const match = docUrl.match(/\/document\/d\/([a-zA-Z0-9_-]+)(\/|$)/);
    return match ? match[1] : null;
  } catch (error) {
    logError(`Error extracting document ID: ${error.message}`, 'extractDocumentId');
    return null;
  }
}

/**
 * Creates an email with the report attached
 * @param {string} recipient Email recipient
 * @param {string} reportUrl Report URL
 * @param {string} customMessage Custom message for email body
 * @param {Object} options Additional email options
 * @returns {boolean} Whether the email was sent successfully
 */
function sendReportEmail(recipient, reportUrl, customMessage = '', options = {}) {
  try {
    const documentId = extractDocumentId(reportUrl);
    if (!documentId) {
      throw new Error('Invalid report URL');
    }
    
    // Get report document info
    const doc = DocumentApp.openById(documentId);
    const docTitle = doc.getName();
    
    // Set up email parameters
    const subject = options.subject || `Financial Report: ${docTitle}`;
    let body = customMessage || `Please find attached your financial report.`;
    
    // Add link to the document
    body += `\n\nView the report online: ${reportUrl}`;
    
    // Add footer
    body += `\n\nThis report was generated by Gemini Financial AI on ${new Date().toLocaleString()}.`;
    
    // Try to create PDF attachment if requested
    let pdfUrl = null;
    if (options.includePDF) {
      try {
        pdfUrl = generateReportPDF(documentId);
        body += `\n\nPDF Version: ${pdfUrl}`;
      } catch (pdfError) {
        logError(`Error creating PDF: ${pdfError.message}`, 'sendReportEmail');
        // Continue without PDF
      }
    }
    
    // Send the email
    GmailApp.sendEmail(recipient, subject, body, {
      name: 'Gemini Financial AI',
      replyTo: Session.getActiveUser().getEmail()
    });
    
    return true;
  } catch (error) {
    logError(`Error sending report email: ${error.message}`, 'sendReportEmail');
    return false;
  }
}

/**
 * Creates a formatted financial narrative for the report
 * @param {Object} statistics Financial statistics to include
 * @param {string} currency Currency code for formatting
 * @param {string} locale Locale for formatting
 * @returns {string} Formatted narrative
 */
function createFinancialNarrative(statistics, currency = 'USD', locale = 'en-US') {
  if (!statistics) return '';
  
  try {
    const formatter = new Intl.NumberFormat(locale, {
      style: 'currency',
      currency: currency
    });
    
    let narrative = '## Financial Overview\n\n';
    
    if (statistics.totalAmount !== undefined) {
      narrative += `The total amount involved is ${formatter.format(statistics.totalAmount)}. `;
    }
    
    if (statistics.averageAmount !== undefined) {
      narrative += `On average, transactions amount to ${formatter.format(statistics.averageAmount)}. `;
    }
    
    if (statistics.minAmount !== undefined && statistics.maxAmount !== undefined) {
      narrative += `Values range from ${formatter.format(statistics.minAmount)} to ${formatter.format(statistics.maxAmount)}. `;
    }
    
    if (statistics.outlierCount !== undefined && statistics.outlierCount > 0) {
      narrative += `\n\n**Key Finding:** ${statistics.outlierCount} transaction${statistics.outlierCount !== 1 ? 's' : ''} `;
      narrative += `${statistics.outlierCount !== 1 ? 'are' : 'is'} identified as statistical outliers. `;
      
      if (statistics.outlierImpact !== undefined) {
        narrative += `These outliers represent a total financial impact of ${formatter.format(statistics.outlierImpact)}. `;
      }
    }
    
    return narrative;
  } catch (error) {
    logError(`Error creating financial narrative: ${error.message}`, 'createFinancialNarrative');
    return '';
  }
}

/**
 * Validates and normalizes a Google Docs URL for consistent handling
 * @param {string} docUrl The URL to validate and normalize
 * @returns {string|null} Normalized document URL or null if invalid
 */
function validateAndNormalizeDocUrl(docUrl) {
  if (!docUrl) return null;
  
  try {
    // Extract document ID using regex pattern
    const docIdMatch = docUrl.match(/[-\w]{25,}|docs\.google\.com\/document\/d\/([-\w]{25,})/);
    
    // If we found a valid document ID
    if (docIdMatch && docIdMatch[1]) {
      const docId = docIdMatch[1];
      
      // Check if it's a valid doc by attempting to open it
      try {
        DocumentApp.openById(docId);
        // If no error, return a properly formatted URL
        return `https://docs.google.com/document/d/${docId}/edit`;
      } catch (docError) {
        logError(`Invalid document ID: ${docError.message}`);
        return null;
      }
    } else if (docIdMatch) {
      // If we have a match but not in the capture group, it might be the ID directly
      const possibleId = docIdMatch[0];
      try {
        DocumentApp.openById(possibleId);
        return `https://docs.google.com/document/d/${possibleId}/edit`;
      } catch (docError) {
        logError(`Invalid document ID format: ${docError.message}`);
        return null;
      }
    }
    
    return null;
  } catch (error) {
    logError(`Error validating document URL: ${error.message}`);
    return null;
  }
}

/**
 * Append content to an existing document
 * @param {string} docUrl URL of the document
 * @param {string} content Content to append
 * @returns {boolean} Success status
 */
function appendToReport(docUrl, content) {
  try {
    // Validate the URL first
    const validDocUrl = validateAndNormalizeDocUrl(docUrl);
    
    if (!validDocUrl) {
      throw new Error("Invalid document URL format");
    }
    
    // Extract document ID from the valid URL
    const docId = validDocUrl.match(/[-\w]{25,}/)[0];
    const doc = DocumentApp.openById(docId);
    
    if (!doc) {
      throw new Error("Could not open document");
    }
    
    const body = doc.getBody();
    
    // Add a separator
    body.appendParagraph('───────────────────────────────────')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
        
    // Add a heading for the appended content
    body.appendParagraph('Additional Analysis')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
        
    // Append the content
    const contentParagraph = body.appendParagraph(content);
    
    // Save the document
    doc.saveAndClose();
    
    return true;
  } catch (error) {
    logError(`Error appending to report: ${error.message}`);
    throw new Error(`Failed to append to report: ${error.message}`);
  }
}

/**
 * Validates and normalizes a Google Document URL
 * @param {string} url The URL to validate
 * @returns {string|null} The normalized URL or null if invalid
 */
function validateAndNormalizeDocUrl(url) {
  if (!url) return null;
  
  try {
    // Extract the document ID from various URL formats
    let docId = null;
    
    // Full URL format: https://docs.google.com/document/d/DOCUMENT_ID/edit...
    const fullUrlMatch = url.match(/^https:\/\/docs\.google\.com\/document\/d\/([-\w]{25,})/);
    if (fullUrlMatch && fullUrlMatch[1]) {
      docId = fullUrlMatch[1];
    }
    
    // Handle sharing URL format: https://docs.google.com/document/d/e/DOCUMENT_ID/pub
    const sharingUrlMatch = url.match(/^https:\/\/docs\.google\.com\/document\/d\/e\/([-\w]{25,})/);
    if (sharingUrlMatch && sharingUrlMatch[1]) {
      docId = sharingUrlMatch[1];
    }
    
    // Check if it's already just a document ID
    if (!docId && /^[-\w]{25,}$/.test(url)) {
      docId = url;
    }
    
    if (docId) {
      // Validate that the document actually exists
      try {
        DocumentApp.openById(docId);
        // Return normalized URL
        return `https://docs.google.com/document/d/${docId}/edit`;
      } catch (e) {
        logError(`Invalid document ID: ${docId} - ${e.message}`);
        return null;
      }
    }
    
    logError(`URL format not recognized: ${url}`);
    return null;
  } catch (error) {
    logError(`Error validating document URL: ${error.message}`);
    return null;
  }
}

/**
 * Safely appends content to a report document with enhanced error handling
 * @param {string} docUrl The URL of the document
 * @param {string} content The content to append
 * @returns {boolean} True if successful, false otherwise
 */
function appendToReport(docUrl, content) {
  try {
    let docId;
    
    // Handle different URL formats
    if (docUrl.includes('/document/d/')) {
      // Extract ID from full URL
      const match = docUrl.match(/\/document\/d\/([-\w]{25,})/);
      if (match && match[1]) {
        docId = match[1];
      }
    } else if (/^[-\w]{25,}$/.test(docUrl)) {
      // It's already just an ID
      docId = docUrl;
    }
    
    if (!docId) {
      logError(`Could not extract document ID from URL: ${docUrl}`);
      return false;
    }
    
    // Open the document and append content
    const doc = DocumentApp.openById(docId);
    if (!doc) {
      logError(`Could not open document with ID: ${docId}`);
      return false;
    }
    
    const body = doc.getBody();
    body.appendParagraph('\n'); // Add spacing
    body.appendParagraph('AI-Generated Pattern Analysis')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
    
    // Split content by paragraphs and add each one
    const paragraphs = content.split('\n\n');
    paragraphs.forEach(paragraph => {
      if (paragraph.trim()) {
        body.appendParagraph(paragraph);
      }
    });
    
    return true;
  } catch (error) {
    logError(`Error appending to report: ${error.message}`);
    return false;
  }
}

/**
 * Formats a date string based on locale preferences
 * @param {Date} date The date to format
 * @param {string} locale The locale code
 * @param {string} format The format type ('short', 'long', 'full', etc.)
 * @returns {string} The formatted date string
 */
function formatLocalizedDate(date, locale, format) {
  if (!date || !(date instanceof Date)) return 'N/A';
  
  try {
    switch(format) {
      case 'short':
        return Utilities.formatDate(date, locale, 'MM/dd/yyyy');
      case 'long':
        return Utilities.formatDate(date, locale, 'MMMM dd, yyyy');
      case 'full':
        return Utilities.formatDate(date, locale, 'EEEE, MMMM dd, yyyy');
      case 'iso':
        return Utilities.formatDate(date, locale, 'yyyy-MM-dd');
      default:
        return Utilities.formatDate(date, locale, 'MM/dd/yyyy');
    }
  } catch(e) {
    logError(`Error formatting date: ${e.message}`);
    return date.toISOString().split('T')[0]; // Fallback to ISO format
  }
}

/**
 * Adds a dynamic Table of Contents to the document
 * @param {GoogleAppsScript.Document.Document} doc The document to add the TOC to
 * @returns {boolean} True if successful
 */
function insertDynamicTableOfContents(doc) {
  try {
    // Extract all headings from the document
    const body = doc.getBody();
    const paragraphs = body.getParagraphs();
    const headings = [];
    
    // Find all headings and their levels
    paragraphs.forEach(para => {
      const heading = para.getHeading();
      if (heading !== DocumentApp.ParagraphHeading.NORMAL &&
          heading !== DocumentApp.ParagraphHeading.TITLE) {
        headings.push({
          text: para.getText(),
          level: getHeadingLevel(heading)
        });
      }
    });
    
    // Insert TOC at the beginning of the document
    if (headings.length > 0) {
      // Create a TOC section
      body.insertParagraph(0, '')
          .setHeading(DocumentApp.ParagraphHeading.NORMAL);
      
      body.insertParagraph(0, 'Table of Contents')
          .setHeading(DocumentApp.ParagraphHeading.HEADING1);
      
      // Add TOC entries with indentation based on heading level
      let index = 2; // Start after the TOC title and blank line
      headings.forEach(heading => {
        const indent = '  '.repeat(heading.level - 1);
        body.insertParagraph(index++, `${indent}• ${heading.text}`);
      });
      
      // Add spacing after TOC
      body.insertParagraph(index, '')
          .setHeading(DocumentApp.ParagraphHeading.NORMAL);
      
      return true;
    }
    
    return false;
  } catch (error) {
    logError(`Error creating table of contents: ${error.message}`);
    return false;
  }
}

/**
 * Gets the numeric level for a heading enum
 * @param {DocumentApp.ParagraphHeading} heading The heading enum
 * @returns {number} The heading level (1-6)
 */
function getHeadingLevel(heading) {
  switch(heading) {
    case DocumentApp.ParagraphHeading.HEADING1:
      return 1;
    case DocumentApp.ParagraphHeading.HEADING2:
      return 2;
    case DocumentApp.ParagraphHeading.HEADING3:
      return 3;
    case DocumentApp.ParagraphHeading.HEADING4:
      return 4;
    case DocumentApp.ParagraphHeading.HEADING5:
      return 5;
    case DocumentApp.ParagraphHeading.HEADING6:
      return 6;
    default:
      return 1;
  }
}

/**
 * Creates a templated HTML dialog with standard styling
 * @param {string} templateName Name of the HTML template file
 * @param {Object} data Data to pass to the template
 * @param {Object} options Dialog options
 * @returns {GoogleAppsScript.HTML.HtmlOutput} The HTML output
 */
function createTemplatedDialog(templateName, data = {}, options = {}) {
  // Load the template file
  let template = HtmlService.createTemplateFromFile(templateName);
  
  // Add data to the template
  Object.assign(template, data);
  
  // Evaluate the template
  let htmlOutput = template.evaluate();
  
  // Apply options
  if (options.width) htmlOutput.setWidth(options.width);
  if (options.height) htmlOutput.setHeight(options.height);
  if (options.title) htmlOutput.setTitle(options.title);
  
  return htmlOutput;
}

/**
 * Sanitizes HTML content to prevent XSS
 * @param {string} content The HTML content to sanitize
 * @returns {string} Sanitized HTML
 */
function sanitizeHTML(content) {
  if (!content) return '';
  
  return content
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"//g, '&quot;')
    .replace(/'/g, '&#039;');
}

/**
 * Applies standard formatting to a document
 * @param {GoogleAppsScript.Document.Document} doc The document to format
 * @param {Object} options Formatting options
 */
function applyStandardFormatting(doc, options = {}) {
  try {
    const body = doc.getBody();
    
    // Set default font and size
    body.setFontFamily(options.fontFamily || 'Arial');
    body.setFontSize(options.fontSize || 11);
    
    // Format headings and paragraphs
    const paragraphs = body.getParagraphs();
    paragraphs.forEach(para => {
      const heading = para.getHeading();
      
      // Format based on heading type
      switch (heading) {
        case DocumentApp.ParagraphHeading.TITLE:
          para.setFontSize(18).setFontFamily('Arial').setBold(true);
          break;
        case DocumentApp.ParagraphHeading.HEADING1:
          para.setFontSize(16).setFontFamily('Arial').setBold(true)
              .setForegroundColor('#1155cc');
          break;
        case DocumentApp.ParagraphHeading.HEADING2:
          para.setFontSize(14).setFontFamily('Arial').setBold(true)
              .setForegroundColor('#1c4587');
          break;
        case DocumentApp.ParagraphHeading.HEADING3:
          para.setFontSize(12).setFontFamily('Arial').setBold(true);
          break;
        default:
          // Regular paragraphs
          if (options.lineSpacing) {
            para.setLineSpacing(options.lineSpacing);
          }
          break;
      }
    });
    
    // Apply page margins if specified
    if (options.margins) {
      const margins = options.margins;
      body.setMarginTop(margins.top || 72);     // 1 inch in points
      body.setMarginBottom(margins.bottom || 72);
      body.setMarginLeft(margins.left || 72);
      body.setMarginRight(margins.right || 72);
    }
    
    return true;
  } catch (error) {
    logError(`Error applying document formatting: ${error.message}`);
    return false;
  }
}

/**
 * Computes comprehensive numeric statistics for a specified field in anomalies
 * @param {Array<Object>} anomalies Array of anomaly objects to analyze
 * @param {string} field The field name to compute statistics for
 * @returns {Object} Object containing computed statistics
 */
function computeNumericStats(anomalies, field) {
  try {
    if (!anomalies || anomalies.length === 0) {
      return { count: 0, error: 'No anomalies to analyze' };
    }

    // Extract numeric values from the specified field
    const values = anomalies
      .map(a => {
        const value = a[field];
        if (typeof value === 'number') return value;
        if (typeof value === 'string') {
          const parsed = parseFloat(value);
          return isNaN(parsed) ? null : parsed;
        }
        return null;
      })
      .filter(val => val !== null && !isNaN(val));

    if (values.length === 0) {
      return { count: 0, error: `No valid numeric values found in field '${field}'` };
    }

    // Sort values for percentile calculations
    const sortedValues = [...values].sort((a, b) => a - b);
    const count = values.length;
    
    // Basic statistics
    const sum = values.reduce((acc, val) => acc + val, 0);
    const mean = sum / count;
    const min = sortedValues[0];
    const max = sortedValues[count - 1];
    
    // Median calculation
    const midIndex = Math.floor(count / 2);
    const median = count % 2 === 0 
      ? (sortedValues[midIndex - 1] + sortedValues[midIndex]) / 2 
      : sortedValues[midIndex];
    
    // Quartile calculations
    const q1Index = Math.floor(count * 0.25);
    const q3Index = Math.floor(count * 0.75);
    const q1 = sortedValues[q1Index];
    const q3 = sortedValues[q3Index];
    const iqr = q3 - q1;
    
    // Standard deviation
    const variance = values.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / count;
    const stdDev = Math.sqrt(variance);
    
    // Identify outliers using 1.5 * IQR rule
    const lowerBound = q1 - (1.5 * iqr);
    const upperBound = q3 + (1.5 * iqr);
    const outliers = values.filter(v => v < lowerBound || v > upperBound);
    
    return {
      count,
      sum,
      min,
      max,
      range: max - min,
      mean,
      median,
      q1,
      q3,
      iqr,
      stdDev,
      variance,
      outlierCount: outliers.length,
      outlierPercent: (outliers.length / count) * 100,
      outliers: outliers,
      lowerBound,
      upperBound
    };
  } catch (error) {
    logError(`Error computing numeric stats: ${error.message}`, 'computeNumericStats');
    return { error: `Failed to compute statistics: ${error.message}` };
  }
}

/**
 * Breaks down anomalies by error type and counts occurrences
 * @param {Array<Object>} anomalies The anomalies to analyze
 * @returns {Object} Object mapping error types to counts
 */
function breakdownErrors(anomalies) {
  try {
    if (!anomalies || anomalies.length === 0) {
      return { 'No Data': 0 };
    }
    
    const errorTypes = {};
    
    // Count occurrences of each error type
    anomalies.forEach(anomaly => {
      let errorType = anomaly.errorType || anomaly.type || 'Unspecified';
      
      // If there's an error message but no type, use a substring of the message
      if (errorType === 'Unspecified' && anomaly.message) {
        errorType = anomaly.message.substring(0, 30) + (anomaly.message.length > 30 ? '...' : '');
      }
      
      // Increment the count for this error type
      errorTypes[errorType] = (errorTypes[errorType] || 0) + 1;
    });
    
    // If we didn't find any error types, provide a fallback
    if (Object.keys(errorTypes).length === 0) {
      return { 'Unclassified Errors': anomalies.length };
    }
    
    return errorTypes;
  } catch (error) {
    logError(`Error breaking down errors: ${error.message}`, 'breakdownErrors');
    return { 'Error Processing Data': anomalies?.length || 0 };
  }
}

/**
 * Breaks down anomalies by category and counts occurrences
 * @param {Array<Object>} anomalies The anomalies to analyze
 * @returns {Object} Object mapping categories to counts
 */
function breakdownCategories(anomalies) {
  try {
    if (!anomalies || anomalies.length === 0) {
      return { 'No Data': 0 };
    }
    
    const categories = {};
    
    // Count occurrences of each category
    anomalies.forEach(anomaly => {
      const category = anomaly.category || 'Uncategorized';
      categories[category] = (categories[category] || 0) + 1;
    });
    
    return categories;
  } catch (error) {
    logError(`Error breaking down categories: ${error.message}`, 'breakdownCategories');
    return { 'Error Processing Categories': anomalies?.length || 0 };
  }
}

/**
 * Sends a generated report by email with anomaly summary
 * @param {string} recipient Email address to send the report to
 * @param {string} reportUrl URL to the generated report document
 * @param {Array<Object>} anomalies The anomalies included in the report
 * @param {string} customMessage Optional custom message to include in the email
 * @returns {Promise<boolean>} Promise resolving to true if email was sent successfully
 */
async function sendReportEmail(recipient, reportUrl, anomalies, customMessage = '') {
  try {
    if (!recipient || !reportUrl) {
      throw new Error('Missing required parameters: recipient or reportUrl');
    }
    
    // Extract document ID from URL
    const documentId = extractDocumentId(reportUrl);
    
    if (!documentId) {
      throw new Error('Invalid report URL format');
    }
    
    // Get document details
    let docTitle, docDate;
    try {
      const doc = DocumentApp.openById(documentId);
      docTitle = doc.getName();
      docDate = new Date();
    } catch (docError) {
      logError(`Error accessing document: ${docError.message}`, 'sendReportEmail');
      docTitle = 'Financial Analysis Report';
      docDate = new Date();
    }
    
    // Create email subject
    const subject = `Financial Analysis Report: ${docTitle}`;
    
    // Create summary of anomalies
    let anomalySummary = '';
    if (anomalies && anomalies.length > 0) {
      // Group by priority/severity if available
      const byPriority = {};
      anomalies.forEach(a => {
        const priority = a.priority || a.severity || 'medium';
        byPriority[priority] = (byPriority[priority] || 0) + 1;
      });
      
      anomalySummary = `\n\nSummary of Findings:\n`;
      anomalySummary += `- Total anomalies detected: ${anomalies.length}\n`;
      
      // Add breakdown by priority
      Object.entries(byPriority).forEach(([priority, count]) => {
        anomalySummary += `- ${capitalize(priority)} priority issues: ${count}\n`;
      });
      
      // Add financial impact if available
      const amountStats = computeNumericStats(anomalies, 'amount');
      if (amountStats && !amountStats.error) {
        anomalySummary += `- Total financial impact: $${amountStats.sum.toFixed(2)}\n`;
      }
    }
    
    // Build email body
    let emailBody = `Dear User,\n\n`;
    emailBody += `Your financial analysis report "${docTitle}" is now available.\n`;
    
    // Add custom message if provided
    if (customMessage) {
      emailBody += `\n${customMessage}\n`;
    }
    
    // Add anomaly summary if we have anomalies
    if (anomalySummary) {
      emailBody += anomalySummary;
    }
    
    // Add link to the report
    emailBody += `\n\nView the full report here: ${reportUrl}\n\n`;
    
    // Add footer
    emailBody += `This report was generated on ${docDate.toLocaleString()} by Gemini Financial AI.\n`;
    emailBody += `Please do not reply to this automated email.\n`;
    
    // Send the email
    GmailApp.sendEmail(recipient, subject, emailBody, {
      name: 'Gemini Financial AI',
      replyTo: 'no-reply@example.com'
    });
    
    // Log the successful email
    logMessage(`Report email sent to ${recipient} for document ${docTitle}`, 'sendReportEmail');
    
    return true;
  } catch (error) {
    logError(`Failed to send report email: ${error.message}`, 'sendReportEmail');
    return false;
  }
}

/**
 * Helper function to capitalize the first letter of a string
 * @param {string} str The string to capitalize
 * @returns {string} The capitalized string
 */
function capitalize(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1).toLowerCase();
}

/**
 * Logs a message for tracking purposes
 * @param {string} message The message to log
 * @param {string} context The function or context where the message originated
 */
function logMessage(message, context = '') {
  console.log(`[INFO] [${context}] ${message}`);
  try {
    appendToLogSheet('INFO', message, context);
  } catch (e) {
    // Fallback to just console logging if sheet logging fails
  }
}

/**
 * Logs an error message
 * @param {string} message The error message to log
 * @param {string} context The function or context where the error occurred
 */
function logError(message, context = '') {
  console.error(`[ERROR] [${context}] ${message}`);
  try {
    appendToLogSheet('ERROR', message, context);
  } catch (e) {
    // Already failed logging, just continue
  }
}