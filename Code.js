// ==================== GLOBAL CONFIGURATION ====================
const SPREADSHEET_ID = '1U1iGLK968X-QlMLI4rDXMWxtruK_HCgNa-VaV_bd1No';
const CACHE_PREFIX = 'SCHOOL_ERP_';
const SESSION_EXPIRY = 30;

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('School Management ERP')
    .setFaviconUrl('https://img.icons8.com/color/48/000000/school.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== CACHE ====================
function getC(key) { try { return CacheService.getScriptCache().get(CACHE_PREFIX + key); } catch(e) { return null; } }
function setC(key, val, sec) { try { CacheService.getScriptCache().put(CACHE_PREFIX + key, val, sec || 21600); } catch(e) {} }
function delC(key) { try { var keys = key ? [key] : ['SETTINGS','STUDENTS','TEACHERS','FEES','SALARY','EXPENSES']; var c = CacheService.getScriptCache(); keys.forEach(function(k) { c.remove(CACHE_PREFIX + k); }); } catch(e) {} }

// ==================== ERROR LOGGING ====================
function logErr(src, desc, err) {
  try {
    var lock = LockService.getScriptLock(); lock.waitLock(5000);
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('ErrorLogs');
    if (!sheet) { sheet = doc.insertSheet('ErrorLogs'); sheet.getRange(1,1,1,4).setValues([['Timestamp','Source','Description','Error']]); sheet.setFrozenRows(1); }
    sheet.appendRow([new Date(), src, desc, err ? err.toString() : '']);
    lock.releaseLock();
  } catch(e) { console.error(src, desc, err); }
}

// ==================== AUTH ====================
function adminLogin(pwd) {
  try {
    var s = getSettings();
    if (!s || !s.adminPassword) { createDefaultSettings(); return { success: false, message: 'System initialized. Default: admin123' }; }
    if (pwd === s.adminPassword) {
      var token = Utilities.getUuid();
      setC('SESSION_' + token, JSON.stringify({ role: 'admin', time: new Date().toISOString() }), 3600);
      return { success: true, token: token, settings: s };
    }
    return { success: false, message: 'Invalid password' };
  } catch(e) { return { success: false, message: 'Login error' }; }
}

function validateSession(token) {
  if (!token) return false;
  var d = getC('SESSION_' + token); if (!d) return false;
  try {
    var session = JSON.parse(d);
    var exp = new Date(session.time); exp.setMinutes(exp.getMinutes() + SESSION_EXPIRY);
    if (new Date() > exp) { delC('SESSION_' + token); return false; }
    session.time = new Date().toISOString();
    setC('SESSION_' + token, JSON.stringify(session), 3600);
    return true;
  } catch(e) { return false; }
}

function adminLogout(token) { if (token) delC('SESSION_' + token); return { success: true }; }

// ==================== SETTINGS ====================
function getSettings() {
  var cached = getC('SETTINGS'); if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Settings');
    if (!sheet) return createDefaultSettings();
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return createDefaultSettings();
    var settings = {};
    for (var i = 1; i < data.length; i++) { if (data[i][0]) settings[data[i][0]] = data[i][1] || ''; }
    setC('SETTINGS', JSON.stringify(settings));
    return settings;
  } catch(e) { return {}; }
}

function createDefaultSettings() {
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Settings');
    if (sheet) sheet.clear(); else sheet = doc.insertSheet('Settings');
    var defaults = [
      ['schoolName', 'The Excellence School'], ['adminPassword', 'admin123'],
      ['logoUrl', 'https://img.icons8.com/color/96/000000/school.png'],
      ['address', '123 Education Lane, Model Town, Lahore'], ['phone', '+92-300-1234567'],
      ['email', 'info@excellenceschool.edu.pk'], ['website', 'www.excellenceschool.edu.pk'],
      ['principalName', 'Dr. Muhammad Ahmed'], ['motto', 'Knowledge is Power'],
      ['primaryColor', '#1a237e'], ['secondaryColor', '#0d47a1'], ['accentColor', '#ff6f00'],
      ['theme', 'light'], ['language', 'en'], ['session', '2024-2025']
    ];
    sheet.getRange(1,1,1,2).setValues([['Key','Value']]);
    sheet.getRange(2,1,defaults.length,2).setValues(defaults);
    sheet.setFrozenRows(1); delC('SETTINGS');
    var obj = {}; defaults.forEach(function(r) { obj[r[0]] = r[1]; });
    setC('SETTINGS', JSON.stringify(obj)); return obj;
  } catch(e) { logErr('createDefaultSettings', 'Init error', e); return {}; }
}

function updateSettings(obj) {
  var lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Settings');
    if (!sheet) { sheet = doc.insertSheet('Settings'); sheet.getRange(1,1,1,2).setValues([['Key','Value']]); }
    var existing = {}; var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { if (data[i][0]) existing[data[i][0]] = data[i][1]; }
    var merged = {}; for (var k in existing) merged[k] = existing[k]; for (var k in obj) merged[k] = obj[k];
    var entries = []; for (var k in merged) entries.push([k, merged[k]]);
    sheet.clear(); sheet.getRange(1,1,1,2).setValues([['Key','Value']]);
    if (entries.length > 0) sheet.getRange(2,1,entries.length,2).setValues(entries);
    sheet.setFrozenRows(1); delC('SETTINGS'); return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==================== CLASSES & SECTIONS ====================
function getClasses() { return ['Play Group','Nursery','Prep','Class 1','Class 2','Class 3','Class 4','Class 5','Class 6','Class 7','Class 8','Class 9','Class 10','Class 11','Class 12']; }
function getSections() { return ['A','B','C','D']; }

// ==================== STUDENT CRUD ====================
function addStudent(data, imgBase64) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Students');
    if (!sheet) { sheet = doc.insertSheet('Students'); createStudentHeaders(sheet); }
    var sid = 'STU-' + Utilities.getUuid().substring(0,8).toUpperCase();
    var grNo = 'GR-' + new Date().getFullYear() + '-' + String(sheet.getLastRow()).padStart(4,'0');
    var img = imgBase64 ? saveImage(sid, imgBase64) : (data.imageUrl || '');
    sheet.appendRow([sid, grNo, data.fullName||'', data.fatherName||'', data.motherName||'', data.gender||'', data.dob||'', data.age||'', data.cnic||'', data.phone||'', data.whatsapp||'', data.email||'', data.address||'', data.city||'', data.country||'Pakistan', data.religion||'', data.bloodGroup||'', data.class||'', data.section||'', data.rollNo||'', data.house||'', new Date(), data.fee||0, data.discount||0, data.status||'Active', data.notes||'', img, data.fatherPhone||'', data.motherPhone||'', data.emergencyContact||'']);
    delC('STUDENTS'); return { success: true, studentId: sid, grNo: grNo };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateStudent(sid, updates) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Students');
    if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues(); var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === sid) {
        for (var k in updates) { var idx = headers.indexOf(k); if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]); }
        delC('STUDENTS'); return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function deleteStudent(sid) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Students'); if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length-1; i >= 1; i--) {
      if (data[i][0] === sid) { archiveRecord('ArchiveStudents', data[i]); sheet.deleteRow(i+1); delC('STUDENTS'); return { success: true }; }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getStudents() {
  var cached = getC('STUDENTS'); if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Students');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) { var obj = {}; headers.forEach(function(h, idx) { obj[h] = data[i][idx] || ''; }); result.push(obj); }
    setC('STUDENTS', JSON.stringify(result)); return result;
  } catch(e) { return []; }
}

function getStudentById(sid) { var students = getStudents(); return students.find(function(s) { return s.StudentID === sid; }) || null; }

// ==================== TEACHER CRUD ====================
function addTeacher(data, imgBase64) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Teachers');
    if (!sheet) { sheet = doc.insertSheet('Teachers'); createTeacherHeaders(sheet); }
    var tid = 'TCH-' + Utilities.getUuid().substring(0,8).toUpperCase();
    var empCode = 'EMP-' + String(sheet.getLastRow()).padStart(4,'0');
    var img = imgBase64 ? saveImage(tid, imgBase64) : (data.imageUrl || '');
    sheet.appendRow([tid, empCode, data.name||'', data.fatherName||'', data.gender||'', data.dob||'', data.cnic||'', data.phone||'', data.whatsapp||'', data.email||'', data.address||'', data.city||'', data.qualification||'', data.experience||'', data.subjects||'', data.classTeacher||'', new Date(), data.salaryType||'Monthly', data.perLecturePay||0, data.monthlyPackage||0, data.status||'Active', img, data.notes||'']);
    delC('TEACHERS'); return { success: true, teacherId: tid, empCode: empCode };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateTeacher(tid, updates) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Teachers'); if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues(); var headers = data[0];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === tid) { for (var k in updates) { var idx = headers.indexOf(k); if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]); } delC('TEACHERS'); return { success: true }; }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function deleteTeacher(tid) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Teachers'); if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length-1; i >= 1; i--) { if (data[i][0] === tid) { archiveRecord('ArchiveTeachers', data[i]); sheet.deleteRow(i+1); delC('TEACHERS'); return { success: true }; } }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getTeachers() {
  var cached = getC('TEACHERS'); if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Teachers'); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) { var obj = {}; headers.forEach(function(h, idx) { obj[h] = data[i][idx] || ''; }); result.push(obj); }
    setC('TEACHERS', JSON.stringify(result)); return result;
  } catch(e) { return []; }
}

// ==================== ARCHIVE ====================
function archiveRecord(sheetName, rowData) {
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName(sheetName);
    if (!sheet) { sheet = doc.insertSheet(sheetName); if (sheetName === 'ArchiveStudents') createStudentHeaders(sheet); else createTeacherHeaders(sheet); }
    sheet.appendRow(rowData);
  } catch(e) { logErr('archiveRecord', 'Archive error', e); }
}

function getArchivedStudents() { try { var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ArchiveStudents'); if (!sheet) return []; return sheet.getDataRange().getValues().slice(1); } catch(e) { return []; } }
function getArchivedTeachers() { try { var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ArchiveTeachers'); if (!sheet) return []; return sheet.getDataRange().getValues().slice(1); } catch(e) { return []; } }

// ==================== FEE MANAGEMENT ====================
function addFeePayment(studentId, amount, month, receiptNo) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('FeeHistory');
    if (!sheet) { sheet = doc.insertSheet('FeeHistory'); sheet.getRange(1,1,1,5).setValues([['StudentID','Date','Amount','Month','ReceiptNo']]); sheet.setFrozenRows(1); }
    sheet.appendRow([studentId, new Date(), amount, month, receiptNo || Utilities.getUuid().substring(0,8)]);
    delC('FEES'); return { success: true };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getAllFeeHistory() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('FeeHistory'); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; return data.slice(1).map(function(r) { var o = {}; headers.forEach(function(h,i) { o[h] = r[i] || ''; }); return o; });
  } catch(e) { return []; }
}

// ==================== SALARY MANAGEMENT ====================
function addSalaryPayment(teacherId, month, amount, lectures, deductions) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('SalaryHistory');
    if (!sheet) { sheet = doc.insertSheet('SalaryHistory'); sheet.getRange(1,1,1,7).setValues([['TeacherID','Month','Amount','Lectures','Deductions','NetPay','Date']]); sheet.setFrozenRows(1); }
    sheet.appendRow([teacherId, month, amount, lectures||0, deductions||0, amount-(deductions||0), new Date()]);
    delC('SALARY'); return { success: true };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getAllSalaryHistory() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SalaryHistory'); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; return data.slice(1).map(function(r) { var o = {}; headers.forEach(function(h,i) { o[h] = r[i] || ''; }); return o; });
  } catch(e) { return []; }
}

// ==================== EXPENSE MANAGEMENT ====================
function addExpense(data) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('Expenses');
    if (!sheet) { sheet = doc.insertSheet('Expenses'); sheet.getRange(1,1,1,8).setValues([['ExpenseID','Date','Category','Description','Amount','PaymentMode','BillNo','Notes']]); sheet.setFrozenRows(1); }
    var expId = 'EXP-' + Utilities.getUuid().substring(0,8).toUpperCase();
    sheet.appendRow([expId, data.date || new Date(), data.category || 'General', data.description || '', parseFloat(data.amount) || 0, data.paymentMode || 'Cash', data.billNo || '', data.notes || '']);
    delC('EXPENSES'); return { success: true, expenseId: expId };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function updateExpense(expId, data) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Expenses'); if (!sheet) return { success: false };
    var allData = sheet.getDataRange().getValues(); var headers = allData[0];
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][0] === expId) {
        var updates = { 'Date': data.date || allData[i][1], 'Category': data.category || allData[i][2], 'Description': data.description || allData[i][3], 'Amount': parseFloat(data.amount) || allData[i][4], 'PaymentMode': data.paymentMode || allData[i][5], 'BillNo': data.billNo || allData[i][6], 'Notes': data.notes || allData[i][7] };
        for (var k in updates) { var idx = headers.indexOf(k); if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]); }
        delC('EXPENSES'); return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function deleteExpense(expId) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Expenses'); if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length-1; i >= 1; i--) { if (data[i][0] === expId) { sheet.deleteRow(i+1); delC('EXPENSES'); return { success: true }; } }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getExpenses(month, year) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Expenses'); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) { var obj = {}; headers.forEach(function(h, idx) { obj[h] = data[i][idx] || ''; }); result.push(obj); }
    if (month || year) {
      result = result.filter(function(r) {
        var d = new Date(r.Date); var m = d.getMonth() + 1; var y = d.getFullYear();
        if (month && year) return m === parseInt(month) && y === parseInt(year);
        if (month) return m === parseInt(month); if (year) return y === parseInt(year); return true;
      });
    }
    return result;
  } catch(e) { return []; }
}

function getExpenseCategories() { return ['Salary','Rent','Utilities','Stationery','Books','Maintenance','Transport','Food/Canteen','Events','Marketing','Software','Internet','Phone Bill','Office Supplies','Cleaning','Security','General','Other']; }

// ==================== FINANCIAL REPORT ====================
function getFinancialReport(month, year) {
  try {
    var feeHistory = getAllFeeHistory(); var expenses = getExpenses(month, year);
    var filteredFees = feeHistory;
    if (month || year) {
      filteredFees = feeHistory.filter(function(f) {
        var d = new Date(f.Date); var m = d.getMonth() + 1; var y = d.getFullYear();
        if (month && year) return m === parseInt(month) && y === parseInt(year);
        if (month) return m === parseInt(month); if (year) return y === parseInt(year); return true;
      });
    }
    var totalIncome = filteredFees.reduce(function(s, f) { return s + parseFloat(f.Amount || 0); }, 0);
    var totalExpenses = expenses.reduce(function(s, e) { return s + parseFloat(e.Amount || 0); }, 0);
    var profit = totalIncome - totalExpenses;
    var expensesByCategory = {};
    expenses.forEach(function(e) { var cat = e.Category || 'General'; expensesByCategory[cat] = (expensesByCategory[cat] || 0) + parseFloat(e.Amount || 0); });
    var monthlyBreakdown = []; var currentYear = year || new Date().getFullYear();
    for (var m = 1; m <= 12; m++) {
      var monthFees = feeHistory.filter(function(f) { var d = new Date(f.Date); return d.getMonth()+1 === m && d.getFullYear() === currentYear; });
      var monthExpenses = expenses.filter(function(e) { var d = new Date(e.Date); return d.getMonth()+1 === m && d.getFullYear() === currentYear; });
      monthlyBreakdown.push({ month: m, monthName: new Date(currentYear, m-1).toLocaleString('default',{month:'long'}), income: monthFees.reduce(function(s,f){return s+parseFloat(f.Amount||0);},0), expenses: monthExpenses.reduce(function(s,e){return s+parseFloat(e.Amount||0);},0) });
    }
    return { success: true, totalIncome: totalIncome, totalExpenses: totalExpenses, profit: profit, expensesByCategory: expensesByCategory, monthlyBreakdown: monthlyBreakdown, expenseCount: expenses.length, feeCount: filteredFees.length, period: { month: month||'All', year: year||'All' } };
  } catch(e) { return { success: false, message: e.toString() }; }
}

// ==================== IMAGE ====================
function saveImage(id, base64Data) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', id + '.png');
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var parents = DriveApp.getFileById(SPREADSHEET_ID).getParents();
    if (!parents.hasNext()) return '';
    var parentFolder = parents.next(); var folder;
    var folders = parentFolder.getFoldersByName('School_Images');
    if (folders.hasNext()) folder = folders.next(); else folder = parentFolder.createFolder('School_Images');
    var file = folder.createFile(blob); file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) { return ''; }
}

// ==================== MAINTENANCE ====================
function resetProject() { try { delC(); var doc = SpreadsheetApp.openById(SPREADSHEET_ID); var sheet = doc.getSheetByName('Settings'); if (sheet) sheet.clear(); createDefaultSettings(); return { success: true, message: 'Reset. Password: admin123' }; } catch(e) { return { success: false }; } }
function clearSystemCache() { delC(); return { success: true, message: 'Cache cleared' }; }

// ==================== DUMMY DATA ====================
function generateDummyData() {
  try {
    addStudent({ fullName:'Ahmed Khan', fatherName:'Muhammad Khan', motherName:'Fatima Khan', gender:'Male', dob:'2015-03-15', age:'9', cnic:'35202-1234567-1', phone:'0300-1111111', whatsapp:'0300-1111111', email:'parent@test.com', address:'123 Main St', city:'Lahore', religion:'Islam', bloodGroup:'B+', class:'Class 4', section:'A', rollNo:'01', house:'Blue', fee:5000, discount:500, status:'Active', fatherPhone:'0300-9999999', motherPhone:'0300-8888888', emergencyContact:'0300-7777777' });
    addStudent({ fullName:'Fatima Ali', fatherName:'Ali Raza', motherName:'Sara Ali', gender:'Female', dob:'2012-07-22', age:'12', cnic:'35202-9876543-2', phone:'0300-2222222', whatsapp:'0300-2222222', email:'parent2@test.com', address:'456 Model Town', city:'Karachi', religion:'Islam', bloodGroup:'A+', class:'Class 7', section:'B', rollNo:'15', house:'Green', fee:6000, discount:0, status:'Active', fatherPhone:'0300-6666666', motherPhone:'0300-5555555', emergencyContact:'0300-4444444' });
    addStudent({ fullName:'Usman Raza', fatherName:'Raza Ahmed', motherName:'Noor Raza', gender:'Male', dob:'2018-11-10', age:'6', cnic:'35202-5555555-3', phone:'0300-3333333', whatsapp:'0300-3333333', email:'parent3@test.com', address:'789 College Rd', city:'Islamabad', religion:'Islam', bloodGroup:'O+', class:'Class 1', section:'C', rollNo:'08', house:'Red', fee:4000, discount:200, status:'Active', fatherPhone:'0300-3333333', motherPhone:'0300-2222222', emergencyContact:'0300-1111111' });
    addTeacher({ name:'Prof. Abdul Qadir', fatherName:'Late Ahmed', gender:'Male', dob:'1980-05-15', cnic:'35202-1111111-1', phone:'0300-6666666', whatsapp:'0300-6666666', email:'qadir@school.edu.pk', address:'789 College Rd', city:'Lahore', qualification:'M.Sc Mathematics', experience:'15 years', subjects:'Mathematics', classTeacher:'Class 10', salaryType:'Monthly', perLecturePay:0, monthlyPackage:75000, status:'Active' });
    addTeacher({ name:'Ms. Sara Ahmed', fatherName:'Ahmed Ali', gender:'Female', dob:'1985-08-20', cnic:'35202-2222222-2', phone:'0300-7777777', whatsapp:'0300-7777777', email:'sara@school.edu.pk', address:'123 Garden Town', city:'Lahore', qualification:'M.A English', experience:'10 years', subjects:'English', classTeacher:'Class 7', salaryType:'Monthly', perLecturePay:0, monthlyPackage:60000, status:'Active' });
    return { success: true, message: 'Dummy data generated!' };
  } catch(e) { return { success: false }; }
}

// ==================== EXAMINATION REGISTRATION ====================
function registerForExam(studentId, examType, examYear, subjects, feeStatus, specialPermission) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('ExamRegistrations');
    if (!sheet) { sheet = doc.insertSheet('ExamRegistrations'); sheet.getRange(1,1,1,12).setValues([['RegistrationID','StudentID','StudentName','Class','Section','ExamType','ExamYear','Subjects','FeeStatus','PermissionBy','RegistrationDate','Status']]); sheet.setFrozenRows(1); }
    var regId = 'EXR-' + Utilities.getUuid().substring(0,8).toUpperCase();
    var student = getStudentById(studentId); if (!student) return { success: false, message: 'Student not found' };
    sheet.appendRow([regId, studentId, student.FullName, student.Class, student.Section, examType, examYear, Array.isArray(subjects) ? subjects.join(', ') : subjects, feeStatus, specialPermission || '', new Date(), 'Registered']);
    return { success: true, registrationId: regId };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getExamRegistrations(examType, examYear, className) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ExamRegistrations'); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) { var obj = {}; headers.forEach(function(h, idx) { obj[h] = data[i][idx] || ''; }); result.push(obj); }
    if (examType) result = result.filter(function(r) { return r.ExamType === examType; });
    if (examYear) result = result.filter(function(r) { return r.ExamYear === examYear; });
    if (className) result = result.filter(function(r) { return r.Class === className; });
    return result;
  } catch(e) { return []; }
}

function deleteExamRegistration(regId) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ExamRegistrations'); if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length-1; i >= 1; i--) { if (data[i][0] === regId) { sheet.deleteRow(i+1); return { success: true }; } }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getExamTypes() { return ['Annual Examination','Supplementary Examination','Mid-Term Examination','Term Examination']; }
function getDefaultSubjects() { return ['English','Urdu','Mathematics','Science','Islamiyat','Social Studies','Computer Science','Physics','Chemistry','Biology','History','Geography','Civics','Economics','Art & Drawing','Physical Education']; }

// ==================== RESULT CARDS ====================
function saveResultCard(data) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('ResultCards');
    if (!sheet) { sheet = doc.insertSheet('ResultCards'); sheet.getRange(1,1,1,15).setValues([['ResultID','StudentID','StudentName','Class','Section','ExamType','ExamYear','Subjects','Marks','TotalMarks','Percentage','Grade','Remarks','Date','ImageURL']]); sheet.setFrozenRows(1); }
    var resId = 'RES-' + Utilities.getUuid().substring(0,8).toUpperCase();
    var student = getStudentById(data.studentId);
    sheet.appendRow([resId, data.studentId, student ? student.FullName : '', student ? student.Class : '', student ? student.Section : '', data.examType, data.examYear, JSON.stringify(data.subjects || []), JSON.stringify(data.marks || {}), data.totalMarks || 0, data.percentage || 0, data.grade || '', data.remarks || '', new Date(), data.imageUrl || '']);
    return { success: true, resultId: resId };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getResultCards(studentId, examType, examYear, className) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ResultCards'); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) { var obj = {}; headers.forEach(function(h, idx) { obj[h] = data[i][idx] || ''; }); result.push(obj); }
    if (studentId) result = result.filter(function(r) { return r.StudentID === studentId; });
    if (examType) result = result.filter(function(r) { return r.ExamType === examType; });
    if (examYear) result = result.filter(function(r) { return r.ExamYear === examYear; });
    if (className) result = result.filter(function(r) { return r.Class === className; });
    return result;
  } catch(e) { return []; }
}

function deleteResultCard(resId) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ResultCards'); if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length-1; i >= 1; i--) { if (data[i][0] === resId) { sheet.deleteRow(i+1); return { success: true }; } }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

// ==================== TEST RESULTS ====================
function saveTestResult(data) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName('TestResults');
    if (!sheet) { sheet = doc.insertSheet('TestResults'); sheet.getRange(1,1,1,12).setValues([['TestID','StudentID','StudentName','Class','Section','TestType','Subject','Marks','TotalMarks','Percentage','Date','Remarks']]); sheet.setFrozenRows(1); }
    var testId = 'TST-' + Utilities.getUuid().substring(0,8).toUpperCase();
    var student = getStudentById(data.studentId);
    sheet.appendRow([testId, data.studentId, student ? student.FullName : '', student ? student.Class : '', student ? student.Section : '', data.testType, data.subject, data.marks || 0, data.totalMarks || 0, data.percentage || 0, data.date || new Date(), data.remarks || '']);
    return { success: true, testId: testId };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getTestResults(studentId, testType, subject, className) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('TestResults'); if (!sheet) return [];
    var data = sheet.getDataRange().getValues(); if (data.length < 2) return [];
    var headers = data[0]; var result = [];
    for (var i = 1; i < data.length; i++) { var obj = {}; headers.forEach(function(h, idx) { obj[h] = data[i][idx] || ''; }); result.push(obj); }
    if (studentId) result = result.filter(function(r) { return r.StudentID === studentId; });
    if (testType) result = result.filter(function(r) { return r.TestType === testType; });
    if (subject) result = result.filter(function(r) { return r.Subject === subject; });
    if (className) result = result.filter(function(r) { return r.Class === className; });
    return result;
  } catch(e) { return []; }
}

function deleteTestResult(testId) {
  var lock = LockService.getScriptLock(); lock.waitLock(5000);
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('TestResults'); if (!sheet) return { success: false };
    var data = sheet.getDataRange().getValues();
    for (var i = data.length-1; i >= 1; i--) { if (data[i][0] === testId) { sheet.deleteRow(i+1); return { success: true }; } }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getTestTypes() { return ['Weekly Test','Monthly Test','Quarterly Test','Half-Yearly Test','Pre-Board Test']; }

// ==================== CSV EXPORT/IMPORT ====================
function exportAllDataToCSV() {
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ['Students','Teachers','FeeHistory','SalaryHistory','Expenses','ExamRegistrations','ResultCards','TestResults'];
    var csvContent = '';
    sheets.forEach(function(sheetName) {
      var sheet = doc.getSheetByName(sheetName);
      if (sheet) {
        csvContent += '=== ' + sheetName + ' ===\n';
        var data = sheet.getDataRange().getValues();
        data.forEach(function(row) { csvContent += row.map(function(cell) { var val = cell ? cell.toString() : ''; return '"' + val.replace(/"/g, '""') + '"'; }).join(',') + '\n'; });
        csvContent += '\n\n';
      }
    });
    return { success: true, csvData: csvContent, fileName: 'SchoolERP_Backup_' + new Date().toISOString().split('T')[0] + '.csv' };
  } catch(e) { return { success: false }; }
}

function importCSVData(csvContent, sheetName) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var rows = csvContent.split('\n').filter(function(r) { return r.trim(); });
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName(sheetName); if (!sheet) { sheet = doc.insertSheet(sheetName); }
    var data = rows.map(function(r) { var matches = r.match(/"([^"]*)"|([^,]+)/g) || []; return matches.map(function(m) { return m.replace(/^"|"$/g, '').trim(); }); });
    if (data.length > 0) { sheet.clear(); sheet.getRange(1, 1, data.length, data[0].length).setValues(data); sheet.setFrozenRows(1); }
    return { success: true, message: 'Imported ' + data.length + ' rows to ' + sheetName };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function getCSVSheets() { try { var doc = SpreadsheetApp.openById(SPREADSHEET_ID); return doc.getSheets().map(function(s) { return s.getName(); }); } catch(e) { return []; } }

function deleteDataByYear(sheetName, year) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName(sheetName); if (!sheet) return { success: false, message: 'Sheet not found' };
    var data = sheet.getDataRange().getValues(); var headers = data[0];
    var dateCols = ['JoiningDate','Date','AdmissionDate','RegistrationDate']; var dateColIndex = -1;
    dateCols.forEach(function(dc) { var idx = headers.indexOf(dc); if (idx !== -1) dateColIndex = idx; });
    if (dateColIndex === -1) return { success: false, message: 'No date column found' };
    for (var i = data.length - 1; i >= 1; i--) { var d = new Date(data[i][dateColIndex]); if (d.getFullYear() === parseInt(year)) { sheet.deleteRow(i + 1); } }
    return { success: true, message: 'Deleted records from year ' + year };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==================== SHEET HEADERS ====================
function createStudentHeaders(sheet) { sheet.getRange(1,1,1,30).setValues([['StudentID','GRNo','FullName','FatherName','MotherName','Gender','DOB','Age','CNIC','Phone','WhatsApp','Email','Address','City','Country','Religion','BloodGroup','Class','Section','RollNo','House','AdmissionDate','Fee','Discount','Status','Notes','ImageURL','FatherPhone','MotherPhone','EmergencyContact']]); sheet.setFrozenRows(1); }
function createTeacherHeaders(sheet) { sheet.getRange(1,1,1,22).setValues([['TeacherID','EmpCode','Name','FatherName','Gender','DOB','CNIC','Phone','WhatsApp','Email','Address','City','Qualification','Experience','Subjects','ClassTeacher','JoiningDate','SalaryType','PerLecturePay','MonthlyPackage','Status','ImageURL','Notes']]); sheet.setFrozenRows(1); }

function onOpen() { try { createDefaultSettings(); } catch(e) {} }
