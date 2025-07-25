// --- Global Constants for Sheet Names ---
const DASHBOARD_SHEET_NAME = 'Dashboard';
const ROOM_TYPES_SHEET_NAME = 'RoomTypes';
const ROOMS_SHEET_NAME = 'Rooms';
const RESIDENTS_SHEET_NAME = 'Residents';
const BOOKINGS_SHEET_NAME = 'Bookings';
const PAYMENTS_SHEET_NAME = 'Payments';
const README_SHEET_NAME = 'Readme';

// --- Column Indices (0-based) ---
const RT_TYPE_ID_COL = 0, RT_DESCRIPTION_COL = 1, RT_CAPACITY_COL = 2, RT_SEMESTER_RATE_1_COL = 3, RT_SEMESTER_RATE_2_COL = 4, RT_SEMESTER_RATE_3_COL = 5;
const ROOM_ID_COL = 0, ROOM_TYPE_ID_COL = 1, ROOM_CAPACITY_COL = 2, ROOM_OCCUPIED_COL = 3, ROOM_STATUS_COL = 4;
const RES_ID_COL = 0, RES_NAME_COL = 1, RES_CONTACT_COL = 2, RES_EMAIL_COL = 3, RES_STATUS_COL = 4, RES_LAST_PAYMENT_COL = 5, RES_NOTES_COL = 6;
const BOOK_ID_COL = 0, BOOK_RESIDENT_ID_COL = 1, BOOK_ROOM_ID_COL = 2, BOOK_START_DATE_COL = 3, BOOK_END_DATE_COL = 4, BOOK_TOTAL_RATE_COL = 5, BOOK_PAID_AMOUNT_COL = 6, BOOK_STATUS_COL = 7, BOOK_NOTES_COL = 8;
const PAY_ID_COL = 0, PAY_BOOKING_ID_COL = 1, PAY_RESIDENT_ID_COL = 2, PAY_AMOUNT_COL = 3, PAY_DATE_COL = 4, PAY_METHOD_COL = 5, PAY_STATUS_COL = 6;

/**
 * @OnlyCurrentDoc
 */

// --- CORE UI FUNCTIONS ---

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('sidebar');
  template.context = 'WebApp';
  const htmlOutput = template.evaluate();
  htmlOutput.setTitle('Jeffston Hostels App');
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return htmlOutput;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Hostel Actions')
    .addItem('Open Dashboard UI (In-Sheet)', 'showSidebarUI')
    .addSeparator()
    .addItem('⚡ Initial Setup / Reset All Sheets ⚡', 'initialHostelSetup')
    .addToUi();
}

function showSidebarUI() {
  const template = HtmlService.createTemplateFromFile('sidebar');
  template.context = 'Sidebar';
  const html = template.evaluate().setTitle('Jeffston Hostels Dashboard');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getModalHtml(formType, recordId) {
  const template = HtmlService.createTemplateFromFile('modal');
  template.formTypeToLoad = formType; // <-- This must be set!
  template.recordId = recordId || '';
  template.context = 'WebApp';
  return template.evaluate().getContent();
}

function showModalFormInSheet(formType, recordId) {
  const template = HtmlService.createTemplateFromFile('modal');
  template.formTypeToLoad = formType;
  template.recordId = recordId || null;
  template.context = 'Sidebar';
  const html = template.evaluate().setHeight(600).setWidth(500);
  SpreadsheetApp.getUi().showModalDialog(html, `${formType} Entry`);
}

function navigateToSheet(sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(sheetName);
    if (targetSheet) ss.setActiveSheet(targetSheet);
  } catch (e) {
    console.error("Error navigating to sheet:", e.message);
  }
}

// --- OPTIMIZED HELPER FUNCTIONS ---

function findRowIndexById(sheet, id, idColumn) {
    if (!sheet || sheet.getLastRow() < 2) return -1;
    const ids = sheet.getRange(2, idColumn + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const index = ids.findIndex(cellId => cellId == id);
    return (index !== -1) ? index + 2 : -1;
}

function getNextId(sheet, prefix, idColumn) {
    const lastRow = sheet.getLastRow();
    const lastId = (lastRow > 1) ? String(sheet.getRange(lastRow, idColumn + 1).getValue()) : `${prefix}000`;
    const nextIdNum = (parseInt(lastId.replace(prefix, '')) || 0) + 1;
    return `${prefix}${String(nextIdNum).padStart(3, '0')}`;
}


// --- DATA RETRIEVAL FUNCTIONS ---
function getResidentIDs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESIDENTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, RES_ID_COL + 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String).filter(id => id);
}
function getRoomIDs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ROOMS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, ROOM_ID_COL + 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String).filter(id => id);
}
function getBookingIDs() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, BOOK_ID_COL + 1, sheet.getLastRow() - 1, 1).getValues().flat().map(String).filter(id => id);
}
function getSemesterRatesForRoom(roomId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
  const roomTypesSheet = ss.getSheetByName(ROOM_TYPES_SHEET_NAME);
  if (!roomsSheet || !roomTypesSheet || roomsSheet.getLastRow() < 2 || roomTypesSheet.getLastRow() < 2) return [];
  const roomData = roomsSheet.getDataRange().getValues();
  const roomTypeId = roomData.find(r => r[ROOM_ID_COL] == roomId)?.[ROOM_TYPE_ID_COL];
  if (!roomTypeId) return [];
  const roomTypesData = roomTypesSheet.getDataRange().getValues();
  const typeRow = roomTypesData.find(rt => rt[RT_TYPE_ID_COL] == roomTypeId);
  if (!typeRow) return [];
  const rates = [typeRow[RT_SEMESTER_RATE_1_COL], typeRow[RT_SEMESTER_RATE_2_COL], typeRow[RT_SEMESTER_RATE_3_COL]];
  return rates.filter(rate => rate && !isNaN(parseFloat(rate))).map(parseFloat).sort((a,b) => a-b);
}

function getBookingDetails(bookingId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return null;
    const data = sheet.getDataRange().getValues();
    const bookingRow = data.find(row => row[BOOK_ID_COL] == bookingId);
    if (!bookingRow) return null;
    
    return {
      bookingId: bookingRow[BOOK_ID_COL],
      residentId: bookingRow[BOOK_RESIDENT_ID_COL],
      roomId: bookingRow[BOOK_ROOM_ID_COL],
      startDate: new Date(bookingRow[BOOK_START_DATE_COL]).toISOString().split('T')[0],
      endDate: new Date(bookingRow[BOOK_END_DATE_COL]).toISOString().split('T')[0],
      totalRate: bookingRow[BOOK_TOTAL_RATE_COL],
      paidAmount: bookingRow[BOOK_PAID_AMOUNT_COL],
      status: bookingRow[BOOK_STATUS_COL],
      notes: bookingRow[BOOK_NOTES_COL]
    };
  } catch (e) {
    console.error(`getBookingDetails Error: ${e.message}`);
    return null;
  }
}

/**
 * NEW: Fetches details for a single resident.
 */
function getResidentDetails(residentId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESIDENTS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) return null;
    const data = sheet.getDataRange().getValues();
    const residentRow = data.find(row => row[RES_ID_COL] == residentId);
    if (!residentRow) return null;
    
    return {
      residentId: residentRow[RES_ID_COL],
      name: residentRow[RES_NAME_COL],
      contact: residentRow[RES_CONTACT_COL],
      email: residentRow[RES_EMAIL_COL],
      status: residentRow[RES_STATUS_COL],
      notes: residentRow[RES_NOTES_COL]
    };
  } catch (e) {
    console.error(`getResidentDetails Error: ${e.message}`);
    return null;
  }
}

// --- DATA SUBMISSION FUNCTIONS ---

function addBooking(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet) throw new Error("Bookings sheet not found");

    // Sanitize inputs
    const sanitizedData = {
      residentId: sanitizeInput(formData.residentId),
      roomId: sanitizeInput(formData.roomId),
      startDate: new Date(formData.startDate),
      endDate: new Date(formData.endDate),
      totalRate: parseFloat(formData.totalRate) || 0,
      paidAmount: parseFloat(formData.paidAmount) || 0,
      status: sanitizeInput(formData.status) || 'Pending',
      notes: sanitizeInput(formData.notes) || ''
    };

    // Generate new ID
    const newId = getNextId(sheet, "BKG", BOOK_ID_COL);

    // Create new row
    const newRow = [
      newId,
      sanitizedData.residentId,
      sanitizedData.roomId,
      sanitizedData.startDate,
      sanitizedData.endDate,
      sanitizedData.totalRate,
      sanitizedData.paidAmount,
      sanitizedData.status,
      sanitizedData.notes
    ];

    // Append to sheet
    sheet.appendRow(newRow);

    // Update dashboard
    updateDashboardData();

    return makeResponse('success', `Booking ${newId} created successfully!`, { bookingId: newId });
  } catch (e) {
    logError('addBooking', e);
    return makeResponse('error', e.message);
  }
}
function addResident(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESIDENTS_SHEET_NAME);
    if (!sheet) throw new Error("Residents sheet not found");

    // Sanitize inputs
    const sanitizedData = {
      name: sanitizeInput(formData.name),
      contact: sanitizeInput(formData.contact),
      email: sanitizeInput(formData.email),
      status: sanitizeInput(formData.status),
      notes: sanitizeInput(formData.notes)
    };

    // Generate new ID
    const newId = getNextId(sheet, "RES", RES_ID_COL);

    // Create new row
    const newRow = [
      newId,
      sanitizedData.name,
      sanitizedData.contact,
      sanitizedData.email,
      sanitizedData.status,
      new Date(), // Last updated
      sanitizedData.notes
    ];

    // Append to sheet
    sheet.appendRow(newRow);

    // Update dashboard
    updateDashboardData();

    return makeResponse('success', `Resident ${sanitizedData.name} added successfully!`, { residentId: newId });
  } catch (e) {
    logError('addResident', e);
    return makeResponse('error', e.message);
  }
}
function addPayment(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAYMENTS_SHEET_NAME);
    if (!sheet) throw new Error("Payments sheet not found");

    // Sanitize inputs
    const sanitizedData = {
      bookingId: sanitizeInput(formData.bookingId),
      residentId: sanitizeInput(formData.residentId),
      amount: parseFloat(formData.amount) || 0,
      date: new Date(formData.date),
      method: sanitizeInput(formData.method) || 'Cash',
      status: sanitizeInput(formData.status) || 'Paid'
    };

    // Generate new ID
    const newId = getNextId(sheet, "PAY", PAY_ID_COL);

    // Create new row
    const newRow = [
      newId,
      sanitizedData.bookingId,
      sanitizedData.residentId,
      sanitizedData.amount,
      sanitizedData.date,
      sanitizedData.method,
      sanitizedData.status
    ];

    // Append to sheet
    sheet.appendRow(newRow);

    // Update dashboard
    updateDashboardData();

    return makeResponse('success', `Payment ${newId} recorded successfully!`, { paymentId: newId });
  } catch (e) {
    logError('addPayment', e);
    return makeResponse('error', e.message);
  }
}
function updateBooking(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet) throw new Error("Bookings sheet not found");

    const rowIndex = findRowIndexById(sheet, formData.bookingId, BOOK_ID_COL);
    if (rowIndex === -1) throw new Error("Booking not found");

    // Sanitize inputs
    const sanitizedData = {
      residentId: sanitizeInput(formData.residentId),
      roomId: sanitizeInput(formData.roomId),
      startDate: new Date(formData.startDate),
      endDate: new Date(formData.endDate),
      totalRate: parseFloat(formData.totalRate) || 0,
      paidAmount: parseFloat(formData.paidAmount) || 0,
      status: sanitizeInput(formData.status),
      notes: sanitizeInput(formData.notes) || ''
    };

    // Update row
    const updatedRow = [
      formData.bookingId, // Keep original ID
      sanitizedData.residentId,
      sanitizedData.roomId,
      sanitizedData.startDate,
      sanitizedData.endDate,
      sanitizedData.totalRate,
      sanitizedData.paidAmount,
      sanitizedData.status,
      sanitizedData.notes
    ];

    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);

    // Update dashboard
    updateDashboardData();

    return makeResponse('success', `Booking ${formData.bookingId} updated successfully!`);
  } catch (e) {
    logError('updateBooking', e);
    return makeResponse('error', e.message);
  }
}

function deleteBooking(bookingId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET_NAME);
    if (!sheet) throw new Error("Bookings sheet not found");

    const rowIndex = findRowIndexById(sheet, bookingId, BOOK_ID_COL);
    if (rowIndex === -1) throw new Error("Booking not found");

    // Delete the row
    sheet.deleteRow(rowIndex);

    // Update dashboard
    updateDashboardData();

    return makeResponse('success', `Booking ${bookingId} deleted successfully!`);
  } catch (e) {
    logError('deleteBooking', e);
    return makeResponse('error', e.message);
  }
}

/**
 * NEW: Updates an existing resident's details.
 */
function updateResident(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESIDENTS_SHEET_NAME);
    if (!sheet) throw new Error("Residents sheet not found");

    const rowIndex = findRowIndexById(sheet, formData.residentId, RES_ID_COL);
    if (rowIndex === -1) throw new Error("Resident not found");

    // Sanitize inputs
    const sanitizedData = {
      name: sanitizeInput(formData.name),
      contact: sanitizeInput(formData.contact),
      email: sanitizeInput(formData.email),
      status: sanitizeInput(formData.status),
      notes: sanitizeInput(formData.notes)
    };

    // Update row
    const updatedRow = [
      formData.residentId, // Keep original ID
      sanitizedData.name,
      sanitizedData.contact,
      sanitizedData.email,
      sanitizedData.status,
      new Date(), // Update timestamp
      sanitizedData.notes
    ];

    sheet.getRange(rowIndex, 1, 1, updatedRow.length).setValues([updatedRow]);

    // Update dashboard
    updateDashboardData();

    return makeResponse('success', `Resident ${sanitizedData.name} updated successfully!`);
  } catch (e) {
    logError('updateResident', e);
    return makeResponse('error', e.message);
  }
}

/**
 * Utility: Standardized response object
 */
function makeResponse(status, message, data = null) {
  return { status, message, data };
}

/**
 * Utility: Centralized error logging
 */
function logError(location, error) {
  console.error(`[${location}] ${error.message}`);
}



/**
 * Utility: Fallback for unexpected errors
 */
function globalErrorHandler(e) {
  logError('global', e);
  return makeResponse('error', 'An unexpected error occurred. Please try again.');
}

// --- DASHBOARD DATA & SETUP ---
function getDashboardData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
    const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
    const paymentsSheet = ss.getSheetByName(PAYMENTS_SHEET_NAME);
    const residentsSheet = ss.getSheetByName(RESIDENTS_SHEET_NAME); // Get residents sheet
    if (!roomsSheet || !bookingsSheet || !paymentsSheet || !residentsSheet) { return { error: "A required sheet is missing." }; }
    
    // Occupancy and Revenue calculation (same as before)
    const roomData = (roomsSheet.getLastRow() > 1) ? roomsSheet.getDataRange().getValues().slice(1) : [];
    let totalCapacity = 0, totalOccupied = 0;
    roomData.forEach(r => { totalCapacity += (parseInt(r[ROOM_CAPACITY_COL]) || 0); totalOccupied += (parseInt(r[ROOM_OCCUPIED_COL]) || 0); });
    const occupancy = totalCapacity > 0 ? (totalOccupied / totalCapacity) * 100 : 0;
    
    const paymentsData = (paymentsSheet.getLastRow() > 1) ? paymentsSheet.getDataRange().getValues().slice(1) : [];
    const now = new Date();
    let monthlyRevenue = 0;
    paymentsData.forEach(p => {
        const payDate = new Date(p[PAY_DATE_COL]);
        if (p[PAY_STATUS_COL] === 'Paid' && payDate.getMonth() === now.getMonth() && payDate.getFullYear() === now.getFullYear()) {
            monthlyRevenue += (parseFloat(p[PAY_AMOUNT_COL]) || 0);
        }
    });
    
    // Bookings data (same as before)
    const bookingsData = (bookingsSheet.getLastRow() > 1) ? bookingsSheet.getDataRange().getValues().slice(1) : [];
    const today = new Date(); today.setHours(0,0,0,0);
    const sevenDays = new Date(); sevenDays.setDate(today.getDate() + 7);
    const upcomingCheckins = bookingsData.filter(b => new Date(b[BOOK_START_DATE_COL]) >= today && new Date(b[BOOK_START_DATE_COL]) <= sevenDays && b[BOOK_STATUS_COL] === 'Confirmed');
    const upcomingCheckouts = bookingsData.filter(b => new Date(b[BOOK_END_DATE_COL]) >= today && new Date(b[BOOK_END_DATE_COL]) <= sevenDays && b[BOOK_STATUS_COL] === 'Checked-In');
    const recentBookings = bookingsData.slice(-5).reverse().map(b => ({ bookingId: b[BOOK_ID_COL], residentId: b[BOOK_RESIDENT_ID_COL], roomId: b[BOOK_ROOM_ID_COL], status: b[BOOK_STATUS_COL] }));

    // NEW: Get active residents
    const residentsData = (residentsSheet.getLastRow() > 1) ? residentsSheet.getDataRange().getValues().slice(1) : [];
    const activeResidents = residentsData
        .filter(r => r[RES_STATUS_COL] === 'Active')
        .map(r => ({ residentId: r[RES_ID_COL], name: r[RES_NAME_COL], contact: r[RES_CONTACT_COL] }));

    return {
        userEmail: Session.getActiveUser().getEmail(),
        totalOccupancy: occupancy.toFixed(1),
        totalRevenue: monthlyRevenue.toFixed(2),
        upcomingCheckins: upcomingCheckins.map(b => ({ bookingId: b[BOOK_ID_COL], residentId: b[BOOK_RESIDENT_ID_COL], roomId: b[BOOK_ROOM_ID_COL], date: new Date(b[BOOK_START_DATE_COL]).toLocaleDateString('en-GB') })),
        upcomingCheckouts: upcomingCheckouts.map(b => ({ bookingId: b[BOOK_ID_COL], residentId: b[BOOK_RESIDENT_ID_COL], roomId: b[BOOK_ROOM_ID_COL], date: new Date(b[BOOK_END_DATE_COL]).toLocaleDateString('en-GB') })),
        recentBookings: recentBookings,
        activeResidents: activeResidents // Add new data to the return object
    };
}

function updateDashboardData() { getDashboardData(); }
function sanitizeInput(input) { return (typeof input === 'string') ? input.replace(/</g, '&lt;').replace(/>/g, '&gt;') : input; }
function initialHostelSetup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create all required sheets if they don't exist
    const sheetNames = [
      DASHBOARD_SHEET_NAME,
      ROOM_TYPES_SHEET_NAME,
      ROOMS_SHEET_NAME,
      RESIDENTS_SHEET_NAME,
      BOOKINGS_SHEET_NAME,
      PAYMENTS_SHEET_NAME,
      README_SHEET_NAME
    ];
    
    sheetNames.forEach(name => {
      if (!ss.getSheetByName(name)) {
        ss.insertSheet(name);
      }
    });
    
    // Set up headers for each sheet
    setupSheetHeaders();
    
    SpreadsheetApp.getUi().alert('Setup complete! All sheets have been initialized.');
  } catch (e) {
    console.error('Setup error:', e);
    SpreadsheetApp.getUi().alert('Setup failed: ' + e.message);
  }
}

function setupSheetHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Room Types headers
  const roomTypesSheet = ss.getSheetByName(ROOM_TYPES_SHEET_NAME);
  if (roomTypesSheet && roomTypesSheet.getLastRow() === 0) {
    roomTypesSheet.getRange(1, 1, 1, 6).setValues([
      ['Type ID', 'Description', 'Capacity', 'Semester Rate 1', 'Semester Rate 2', 'Semester Rate 3']
    ]);
  }
  
  // Rooms headers
  const roomsSheet = ss.getSheetByName(ROOMS_SHEET_NAME);
  if (roomsSheet && roomsSheet.getLastRow() === 0) {
    roomsSheet.getRange(1, 1, 1, 5).setValues([
      ['Room ID', 'Type ID', 'Capacity', 'Occupied', 'Status']
    ]);
  }
  
  // Residents headers
  const residentsSheet = ss.getSheetByName(RESIDENTS_SHEET_NAME);
  if (residentsSheet && residentsSheet.getLastRow() === 0) {
    residentsSheet.getRange(1, 1, 1, 7).setValues([
      ['Resident ID', 'Name', 'Contact', 'Email', 'Status', 'Last Payment', 'Notes']
    ]);
  }
  
  // Bookings headers
  const bookingsSheet = ss.getSheetByName(BOOKINGS_SHEET_NAME);
  if (bookingsSheet && bookingsSheet.getLastRow() === 0) {
    bookingsSheet.getRange(1, 1, 1, 9).setValues([
      ['Booking ID', 'Resident ID', 'Room ID', 'Start Date', 'End Date', 'Total Rate', 'Paid Amount', 'Status', 'Notes']
    ]);
  }
  
  // Payments headers
  const paymentsSheet = ss.getSheetByName(PAYMENTS_SHEET_NAME);
  if (paymentsSheet && paymentsSheet.getLastRow() === 0) {
    paymentsSheet.getRange(1, 1, 1, 7).setValues([
      ['Payment ID', 'Booking ID', 'Resident ID', 'Amount', 'Date', 'Method', 'Status']
    ]);
  }
}

// Cache frequently accessed data
const CACHE_DURATION = 21600; // 6 hours
const cache = CacheService.getScriptCache();

class DataService {
  static getFromCache(key) {
    const cached = cache.get(key);
    return cached ? JSON.parse(cached) : null;
  }

  static setCache(key, data) {
    cache.put(key, JSON.stringify(data), CACHE_DURATION);
  }

  static async addResident(formData) {
    const lock = LockService.getScriptLock();
    try {
      // Wait up to 30 seconds for other processes to complete
      lock.waitLock(30000);
      
      const sheet = SpreadsheetApp.getActiveSpreadsheet()
        .getSheetByName(RESIDENTS_SHEET_NAME);
      
      if (!sheet) throw new Error("Residents sheet not found");

      // Validate data before processing
      const validatedData = this.validateResidentData(formData);
      
      // Generate new ID with locking to prevent duplicates
      const newId = this.getNextIdWithLock(sheet, "RES", RES_ID_COL);
      
      // Prepare row data
      const newRow = [
        newId,
        validatedData.name,
        validatedData.contact,
        validatedData.email,
        validatedData.status,
        new Date(),
        validatedData.notes
      ];

      // Batch write operations
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, newRow.length)
        .setValues([newRow]);

      // Clear cache to force refresh
      cache.remove('residents');
      
      // Update dashboard in background
      this.triggerDashboardUpdate();

      return { 
        status: 'success', 
        message: `Resident ${validatedData.name} added successfully!`,
        data: { residentId: newId }
      };

    } catch (error) {
      console.error('[addResident]', error);
      return { status: 'error', message: error.message };
    } finally {
      lock.releaseLock();
    }
  }

  static validateResidentData(data) {
    const required = ['name', 'contact', 'status'];
    for (const field of required) {
      if (!data[field] || !data[field].trim()) {
        throw new Error(`${field} is required`);
      }
    }

    return {
      name: sanitizeInput(data.name),
      contact: sanitizeInput(data.contact),
      email: sanitizeInput(data.email) || '',
      status: sanitizeInput(data.status),
      notes: sanitizeInput(data.notes) || ''
    };
  }

  static triggerDashboardUpdate() {
    // Simple dashboard update - triggers don't work in all contexts
    try {
      updateDashboardData();
    } catch (e) {
      console.error('Dashboard update failed:', e);
    }
  }

  static getNextIdWithLock(sheet, prefix, idColumn) {
    // Simplified ID generation for lock context
    return getNextId(sheet, prefix, idColumn);
  }
}
