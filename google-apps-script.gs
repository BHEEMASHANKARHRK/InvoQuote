/**
 * Google Apps Script Code for Professional Quotation System
 * 
 * Instructions:
 * 1. Create a new Google Apps Script project at script.google.com
 * 2. Replace the default code with this script
 * 3. Update the SHEET_NAME and SPREADSHEET_ID variables below
 * 4. Deploy as a web app with execute permissions set to "Anyone"
 * 5. Copy the web app URL and update it in script.js
 */

// Configuration - Update these values
const SHEET_NAME = 'Quotations'; // Name of the sheet tab
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // Optional: specify spreadsheet ID

/**
 * Handle POST requests to store quotation data
 */
function doPost(e) {
  try {
    // Parse the incoming data
    const data = JSON.parse(e.postData.contents);
    
    // Get or create the spreadsheet and sheet
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, SHEET_NAME);
    
    // Ensure headers exist
    setupHeaders(sheet);
    
    // Check for duplicates
    const isDuplicate = checkDuplicateEmail(sheet, data.clientEmail);
    if (isDuplicate) {
      return ContentService
        .createTextOutput(JSON.stringify({ 
          success: false, 
          error: 'Duplicate email found. A quotation for this client already exists.' 
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Add the quotation data
    addQuotationRow(sheet, data);
    
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Quotation saved successfully' }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error saving quotation:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle GET requests (for duplicate checking and testing)
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'checkDuplicate') {
      const email = e.parameter.email;
      const spreadsheet = getOrCreateSpreadsheet();
      const sheet = getOrCreateSheet(spreadsheet, SHEET_NAME);
      const isDuplicate = checkDuplicateEmail(sheet, email);
      
      return ContentService
        .createTextOutput(JSON.stringify({ isDuplicate }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Default response
    return ContentService
      .createTextOutput(JSON.stringify({ 
        status: 'Professional Quotation System API is running',
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error('Error in GET request:', error);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Get or create spreadsheet
 */
function getOrCreateSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== 'YOUR_SPREADSHEET_ID_HERE') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } else {
    // Create a new spreadsheet in Drive
    const spreadsheetName = 'Professional Quotations - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return SpreadsheetApp.create(spreadsheetName);
  }
}

/**
 * Get or create sheet
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    
    // Delete default sheet if this is a new spreadsheet
    const sheets = spreadsheet.getSheets();
    if (sheets.length > 1) {
      const defaultSheet = sheets.find(s => s.getName() === 'Sheet1');
      if (defaultSheet && defaultSheet.getName() !== sheetName) {
        spreadsheet.deleteSheet(defaultSheet);
      }
    }
  }
  
  return sheet;
}

/**
 * Setup column headers
 */
function setupHeaders(sheet) {
  const headers = [
    'Timestamp',
    'Quotation No',
    'Quotation Date',
    'Valid Till Date',
    'Company Name',
    'Company Address',
    'Client Name',
    'Client Email',
    'Client Phone',
    'Client Location',
    'Items Count',
    'Subtotal',
    'Total CGST',
    'Total SGST',
    'Grand Total',
    'Terms & Conditions',
    'Items Details'
  ];
  
  // Check if headers already exist
  const range = sheet.getRange(1, 1, 1, headers.length);
  const existingHeaders = range.getValues()[0];
  
  if (existingHeaders.every(cell => cell === '')) {
    // Set headers
    range.setValues([headers]);
    
    // Format headers
    range.setBackground('#3b82f6');
    range.setFontColor('#ffffff');
    range.setFontWeight('bold');
    range.setFontSize(11);
    range.setWrap(true);
    
    // Auto-resize columns
    sheet.autoResizeColumns(1, headers.length);
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
}

/**
 * Check for duplicate email
 */
function checkDuplicateEmail(sheet, email) {
  if (!email) return false;
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false; // No data rows
  
  // Get all email addresses (column H - Client Email)
  const emailRange = sheet.getRange(2, 8, lastRow - 1, 1);
  const emails = emailRange.getValues().flat();
  
  return emails.some(existingEmail => 
    existingEmail && existingEmail.toString().toLowerCase() === email.toLowerCase()
  );
}

/**
 * Add quotation row to sheet
 */
function addQuotationRow(sheet, data) {
  const timestamp = new Date(data.timestamp || new Date());
  const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  // Format items details for storage
  const itemsDetails = data.items.map(item => 
    `${item.description} | Qty: ${item.quantity} | Rate: ₹${item.rate} | GST: ${item.gstRate}% | Total: ₹${item.total.toFixed(2)}`
  ).join('\n');
  
  const rowData = [
    formattedTimestamp,
    data.quotationNo || '',
    data.quotationDate || '',
    data.validTillDate || '',
    data.companyName || '',
    data.companyAddress || '',
    data.clientName || '',
    data.clientEmail || '',
    data.clientPhone || '',
    data.clientLocation || '',
    data.items.length,
    data.subtotal || 0,
    data.totalCGST || 0,
    data.totalSGST || 0,
    data.grandTotal || 0,
    data.termsConditions || '',
    itemsDetails
  ];
  
  // Find the next empty row
  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;
  
  // Insert the data
  const range = sheet.getRange(nextRow, 1, 1, rowData.length);
  range.setValues([rowData]);
  
  // Format the new row
  formatDataRow(sheet, nextRow, rowData.length);
  
  // Log the action
  console.log(`Quotation added to row ${nextRow}:`, data.quotationNo);
}

/**
 * Format data rows
 */
function formatDataRow(sheet, row, numColumns) {
  const range = sheet.getRange(row, 1, 1, numColumns);
  
  // Alternate row coloring
  if (row % 2 === 0) {
    range.setBackground('#f8f9fa');
  } else {
    range.setBackground('#ffffff');
  }
  
  // Set border
  range.setBorder(true, true, true, true, false, false);
  
  // Set font
  range.setFontSize(10);
  range.setFontFamily('Arial');
  
  // Format currency columns (L, M, N, O - Subtotal, CGST, SGST, Grand Total)
  const currencyColumns = [12, 13, 14, 15];
  currencyColumns.forEach(col => {
    const cellRange = sheet.getRange(row, col);
    cellRange.setNumberFormat('₹#,##0.00');
  });
  
  // Wrap text for long content columns
  const wrapColumns = [6, 16, 17]; // Company Address, Terms, Items Details
  wrapColumns.forEach(col => {
    if (col <= numColumns) {
      const cellRange = sheet.getRange(row, col);
      cellRange.setWrap(true);
    }
  });
}

/**
 * Test function to verify setup
 */
function testSetup() {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, SHEET_NAME);
    setupHeaders(sheet);
    
    // Add test quotation
    const testData = {
      timestamp: new Date().toISOString(),
      quotationNo: 'TEST001',
      quotationDate: '2024-01-15',
      validTillDate: '2024-02-15',
      companyName: 'Test Company Ltd',
      companyAddress: 'Test Address\nTest City, Test State',
      clientName: 'Test Client',
      clientEmail: 'test@example.com',
      clientPhone: '+91 98765 43210',
      clientLocation: 'Test Location',
      items: [
        {
          description: 'Test Item 1',
          gstRate: 18,
          quantity: 1,
          rate: 1000,
          amount: 1000,
          cgst: 90,
          sgst: 90,
          total: 1180
        }
      ],
      subtotal: 1000,
      totalCGST: 90,
      totalSGST: 90,
      grandTotal: 1180,
      termsConditions: 'Test terms and conditions'
    };
    
    addQuotationRow(sheet, testData);
    
    console.log('Setup test completed successfully!');
    console.log('Spreadsheet URL:', spreadsheet.getUrl());
    
    return {
      success: true,
      spreadsheetUrl: spreadsheet.getUrl(),
      message: 'Setup completed successfully'
    };
    
  } catch (error) {
    console.error('Setup test failed:', error);
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * Get all quotations (for future use)
 */
function getAllQuotations() {
  try {
    const spreadsheet = getOrCreateSpreadsheet();
    const sheet = getOrCreateSheet(spreadsheet, SHEET_NAME);
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
    const values = range.getValues();
    
    return values.map(row => ({
      timestamp: row[0],
      quotationNo: row[1],
      quotationDate: row[2],
      validTillDate: row[3],
      companyName: row[4],
      companyAddress: row[5],
      clientName: row[6],
      clientEmail: row[7],
      clientPhone: row[8],
      clientLocation: row[9],
      itemsCount: row[10],
      subtotal: row[11],
      totalCGST: row[12],
      totalSGST: row[13],
      grandTotal: row[14],
      termsConditions: row[15],
      itemsDetails: row[16]
    }));
    
  } catch (error) {
    console.error('Error getting quotations:', error);
    return [];
  }
}