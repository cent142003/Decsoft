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

function addBooking(formData) { /* Omitted for brevity */ }
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
function addPayment(formData) { /* Omitted for brevity */ }
function updateBooking(formData) { /* Omitted for brevity */ }
function deleteBooking(bookingId) { /* Omitted for brevity */ }

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
 * Example: Robust addResident function
 */
function addResident(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RESIDENTS_SHEET_NAME);
    if (!sheet) throw new Error("Residents sheet not found.");
    // Sanitize all input
    const name = sanitizeInput(formData.name);
    const contact = sanitizeInput(formData.contact);
    const email = sanitizeInput(formData.email);
    const status = sanitizeInput(formData.status);
    const notes = sanitizeInput(formData.notes);

    const newId = getNextId(sheet, "RES", RES_ID_COL);
    const newRow = [newId, name, contact, email, status, "", notes];
    sheet.appendRow(newRow);

    updateDashboardData();
    return makeResponse('success', `Resident ${name} added!`, { residentId: newId });
  } catch (e) {
    logError('addResident', e);
    return makeResponse('error', e.message);
  }
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
    const bookingsData = (bookingsData.getLastRow() > 1) ? bookingsData.getDataRange().getValues().slice(1) : [];
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
function initialHostelSetup() { /* Omitted for brevity */ }

const FORM_TYPE = document.getElementById('formType').value;

function setupForm() {
  const title = document.getElementById('formTitle');
  if (!FORM_TYPE) {
    title.textContent = 'Error';
    showError({ message: 'Form type not specified. Please reload.' });
    return;
  }
  // ...existing code...
}

class HostelApp {
  constructor() {
    // Cache DOM elements
    this.formContainer = document.getElementById('formContainer');
    this.context = document.getElementById('context')?.value || 'WebApp';
    this.formType = document.getElementById('formType')?.value;
    this.recordId = document.getElementById('recordId')?.value;
    
    this.isSubmitting = false;
    this.pendingRequests = new Set();
    
    // Initialize with error boundary
    try {
      this.initializeEventListeners();
      this.setupFormValidation();
    } catch (error) {
      console.error('Initialization error:', error);
      this.showError('Failed to initialize application');
    }
  }

  static start() {
    let attempts = 0;
    const maxAttempts = 10;
    
    function tryInitialize() {
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        const app = new HostelApp();
        app.showForm();
        return;
      }
      
      if (++attempts < maxAttempts) {
        setTimeout(tryInitialize, 100);
      } else {
        console.error('Failed to initialize Google Apps Script');
        document.getElementById('formTitle').textContent = 'Failed to load form';
      }
    }
    
    tryInitialize();
  }

  initializeEventListeners() {
    // Use event delegation for better performance
    this.formContainer.addEventListener('submit', this.handleFormSubmit.bind(this));
    this.formContainer.addEventListener('click', this.handleClick.bind(this));
    
    // Add input validation listeners
    this.formContainer.addEventListener('input', this.debounce(this.validateInput.bind(this), HostelApp.DEBOUNCE_DELAY));
  }

  setupFormValidation() {
    const forms = this.formContainer.querySelectorAll('form');
    forms.forEach(form => {
      // Add custom validation rules
      const inputs = form.querySelectorAll('input[required], select[required]');
      inputs.forEach(input => {
        input.addEventListener('invalid', (e) => {
          e.preventDefault();
          input.classList.add('invalid');
        });
      });
    });
  }

  async handleFormSubmit(event) {
    if (!event.target.matches('form')) return;
    event.preventDefault();
    
    if (this.isSubmitting) return;
    
    const form = event.target;
    const submitBtn = form.querySelector('.submit-btn');
    
    try {
      this.isSubmitting = true;
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<div class="spinner"></div>';

      const formData = this.getFormData(form);
      if (!this.validateForm(formData)) {
        throw new Error('Please fill in all required fields correctly');
      }

      const result = await this.submitFormWithRetry(formData);
      
      if (result.status === 'success') {
        this.handleSuccess(result);
      } else {
        throw new Error(result.message);
      }
    } catch (error) {
      this.handleError(error);
    } finally {
      this.isSubmitting = false;
      submitBtn.disabled = false;
      submitBtn.textContent = 'Submit';
    }
  }

  // Helper methods with performance optimizations
  getFormData(form) {
    const formData = {};
    const elements = form.querySelectorAll('[data-field]');
    elements.forEach(element => {
      const field = element.dataset.field;
      let value = element.value.trim();
      
      // Type coercion for numbers
      if (element.type === 'number') {
        value = parseFloat(value) || 0;
      }
      
      formData[field] = value;
    });
    return formData;
  }

  validateForm(formData) {
    // Add validation rules here
    return Object.entries(formData).every(([key, value]) => {
      if (this.requiredFields.includes(key)) {
        return value !== '' && value != null;
      }
      return true;
    });
  }

  async submitFormWithRetry(formData) {
    for (let i = 0; i < HostelApp.RETRY_ATTEMPTS; i++) {
      try {
        return await this.submitToServer(formData);
      } catch (error) {
        if (i === HostelApp.RETRY_ATTEMPTS - 1) throw error;
        await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
      }
    }
  }

  submitToServer(formData) {
    return new Promise((resolve, reject) => {
      const requestId = Date.now();
      this.pendingRequests.add(requestId);

      google.script.run
        .withSuccessHandler(result => {
          this.pendingRequests.delete(requestId);
          resolve(result);
        })
        .withFailureHandler(error => {
          this.pendingRequests.delete(requestId);
          reject(error);
        })
        [this.getServerFunction()](formData);
    });
  }

  // Utility methods
  debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  handleSuccess(result) {
    alert(result.message);
    if (window.opener?.loadDashboardData) {
      window.opener.loadDashboardData();
    }
    this.closeModal();
  }

  handleError(error) {
    console.error('Form submission error:', error);
    alert(`Error: ${error.message || 'Failed to submit form'}`);
  }
}

// Static properties for HostelApp class
HostelApp.DEBOUNCE_DELAY = 300; // ms
HostelApp.RETRY_ATTEMPTS = 3;

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
      if (!data[field]?.trim()) {
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
    // Schedule dashboard update to run in background
    ScriptApp.newTrigger('updateDashboardData')
      .timeBased()
      .after(1000)
      .create();
  }
}
