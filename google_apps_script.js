// ================================================================
// ग्राम गाडी ट्रॅकर — Google Apps Script v2
// ================================================================
// Setup:
//   1. Google Sheets उघडा → Extensions → Apps Script
//   2. हे सर्व code paste करा (Replace करा)
//   3. Deploy → New deployment → Web app
//      Execute as: Me | Who has access: Anyone
//   4. URL copy करा → App मध्ये Developer Settings मध्ये paste करा
// ================================================================

// Sheet names — बदलू नका
const S_TRIPS    = 'Trips';
const S_SETTINGS = 'Settings';
const S_USERS    = 'Users';
const S_STOPS    = 'Stops';
const S_PAYMENTS = 'Payments';
const S_SUMMARY  = 'Monthly Summary';

// ---------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result = {};
    switch (data.action) {
      case 'addTrip':    result = addTrip(data.trip);         break;
      case 'syncBulk':   result = syncBulk(data.trips);       break;
      case 'saveSettings': result = saveSettings(data.settings); break;
      case 'getSettings':  result = getSettings();             break;
      case 'getUsers':     result = getUsers();                break;
      case 'getStops':     result = getStops();                break;
      case 'addPayment':   result = addPayment(data.payment);  break;
      default: result = { error: 'Unknown action: ' + data.action };
    }
    return jsonResponse({ success: true, ...result });
  } catch(err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function doGet(e) {
  // Test endpoint — browser मध्ये URL टाकून test करता येईल
  if (e.parameter.test) {
    return jsonResponse({ status: 'ग्राम गाडी ट्रॅकर Script चालू आहे ✅', time: new Date().toISOString(), sheets: getSheetNames() });
  }
  return jsonResponse({ status: 'OK', time: new Date().toISOString() });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

// ---------------------------------------------------------------
// TRIPS SHEET
// ---------------------------------------------------------------
function ensureTripsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_TRIPS);
  if (!sh) {
    sh = ss.insertSheet(S_TRIPS);
    const headers = ['ID','दिनांक','वेळ','चालक ID','चालकाचे नाव','थांब्याचे नाव','रक्कम (₹)','Latitude','Longitude','महिना','वर्ष','Sync Time'];
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setBackground('#1a6b3a').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1,80); sh.setColumnWidth(2,90); sh.setColumnWidth(5,140); sh.setColumnWidth(6,160);
  }
  return sh;
}

function addTrip(trip) {
  const sh = ensureTripsSheet();
  // Duplicate check
  const data = sh.getDataRange().getValues();
  if (data.some((r,i) => i>0 && String(r[0]) === String(trip.id))) {
    return { skipped: true, reason: 'duplicate' };
  }
  sh.appendRow([trip.id, trip.date, trip.time, trip.driverId, trip.driverName, trip.stopName, trip.amount, trip.lat||0, trip.lng||0, trip.month, trip.year, new Date().toISOString()]);
  // Alternate row color
  const lr = sh.getLastRow();
  if (lr % 2 === 0) sh.getRange(lr, 1, 1, 12).setBackground('#f0f8f3');
  updateMonthlySummary();
  return { added: true, row: lr };
}

function syncBulk(trips) {
  const sh = ensureTripsSheet();
  const data = sh.getDataRange().getValues();
  const existing = new Set(data.slice(1).map(r => String(r[0])));
  let added = 0, skipped = 0;
  trips.forEach(trip => {
    if (existing.has(String(trip.id))) { skipped++; return; }
    sh.appendRow([trip.id, trip.date, trip.time, trip.driverId, trip.driverName, trip.stopName, trip.amount, trip.lat||0, trip.lng||0, trip.month, trip.year, new Date().toISOString()]);
    added++;
  });
  if (added > 0) updateMonthlySummary();
  return { added, skipped };
}

// ---------------------------------------------------------------
// SETTINGS SHEET (Central config — सर्व devices साठी)
// ---------------------------------------------------------------
function ensureSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(S_SETTINGS);
    sh.appendRow(['Key','Value','Description']);
    sh.getRange(1,1,1,3).setBackground('#1565c0').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1,160); sh.setColumnWidth(2,220); sh.setColumnWidth(3,260);
    // Default values
    const defaults = [
      ['appName','ग्राम गाडी ट्रॅकर','App चे नाव'],
      ['tagline','ग्रामपंचायत वाहन व्यवस्थापन','App Tagline'],
      ['logo','🚌','Logo Emoji'],
      ['gramName','ग्रामपंचायत','ग्रामपंचायतीचे नाव'],
      ['color','#1a6b3a','Theme Color'],
      ['defaultRate','150','Default प्रति फेरी दर (₹)'],
      ['fenceMeters','10','Geo-fence अंतर (मीटर)'],
      ['drvPhone','','चालकाचा WhatsApp नंबर'],
      ['srpPhone','','सरपंचाचा WhatsApp नंबर'],
    ];
    defaults.forEach(r => sh.appendRow(r));
  }
  return sh;
}

function saveSettings(settings) {
  const sh = ensureSettingsSheet();
  const data = sh.getDataRange().getValues();
  const keyRow = {};
  data.forEach((r, i) => { if (i > 0) keyRow[r[0]] = i+1; });
  Object.entries(settings).forEach(([k, v]) => {
    if (keyRow[k]) sh.getRange(keyRow[k], 2).setValue(v);
    else sh.appendRow([k, v, '']);
  });
  return { saved: true };
}

function getSettings() {
  ensureSettingsSheet();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_SETTINGS);
  const data = sh.getDataRange().getValues();
  const settings = {};
  data.forEach((r,i) => { if(i>0 && r[0]) settings[r[0]] = r[1]; });
  return { settings };
}

// ---------------------------------------------------------------
// USERS SHEET
// ---------------------------------------------------------------
function ensureUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_USERS);
  if (!sh) {
    sh = ss.insertSheet(S_USERS);
    sh.appendRow(['Login ID','Password','Role','नाव','Phone']);
    sh.getRange(1,1,1,5).setBackground('#6a1b9a').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    // Sample users
    sh.appendRow(['driver1','pass123','user','रामराव शिंदे','9876543210']);
    sh.appendRow(['admin1','admin123','admin','ग्रामसेवक जाधव','9876543220']);
  }
  return sh;
}

function getUsers() {
  ensureUsersSheet();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_USERS);
  const data = sh.getDataRange().getValues();
  const users = data.slice(1).filter(r=>r[0]).map(r=>({ id:r[0], pass:r[1], role:r[2], name:r[3], phone:r[4]||'' }));
  return { users };
}

// ---------------------------------------------------------------
// STOPS SHEET
// ---------------------------------------------------------------
function ensureStopsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_STOPS);
  if (!sh) {
    sh = ss.insertSheet(S_STOPS);
    sh.appendRow(['ID','थांब्याचे नाव','Latitude','Longitude','प्रति फेरी दर (₹)']);
    sh.getRange(1,1,1,5).setBackground('#e67e00').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
    // Sample stops
    sh.appendRow([1,'ग्रामपंचायत कार्यालय',17.6805,74.0183,150]);
    sh.appendRow([2,'मुख्य बाजारपेठ',17.6820,74.0200,120]);
    sh.appendRow([3,'प्राथमिक शाळा',17.6790,74.0165,100]);
    sh.appendRow([4,'आरोग्य केंद्र',17.6835,74.0210,130]);
    sh.appendRow([5,'रेल्वे स्थानक',17.6770,74.0150,200]);
  }
  return sh;
}

function getStops() {
  ensureStopsSheet();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_STOPS);
  const data = sh.getDataRange().getValues();
  const stops = data.slice(1).filter(r=>r[0]).map(r=>({ id:r[0], name:r[1], lat:r[2], lng:r[3], rate:r[4] }));
  return { stops };
}

// ---------------------------------------------------------------
// PAYMENTS SHEET
// ---------------------------------------------------------------
function addPayment(payment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(S_PAYMENTS);
  if (!sh) {
    sh = ss.insertSheet(S_PAYMENTS);
    sh.appendRow(['महिना','वर्ष','चालक ID','चालकाचे नाव','एकूण फेऱ्या','एकूण रक्कम (₹)','बिल तारीख']);
    sh.getRange(1,1,1,7).setBackground('#c0392b').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow([payment.month,payment.year,payment.driverId,payment.driverName,payment.totalTrips,payment.totalAmount,new Date().toLocaleString('mr-IN')]);
  return { saved: true };
}

// ---------------------------------------------------------------
// MONTHLY SUMMARY (auto-generated)
// ---------------------------------------------------------------
function updateMonthlySummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trSh = ss.getSheetByName(S_TRIPS);
  if (!trSh) return;
  let sumSh = ss.getSheetByName(S_SUMMARY);
  if (!sumSh) {
    sumSh = ss.insertSheet(S_SUMMARY);
    sumSh.setFrozenRows(1);
  }
  const data = trSh.getDataRange().getValues().slice(1).filter(r=>r[0]);
  const grouped = {};
  data.forEach(r => {
    const key = `${r[10]}_${r[9]}_${r[4]}`;
    if (!grouped[key]) grouped[key] = { year:r[10], month:r[9], name:r[4], count:0, total:0 };
    grouped[key].count++; grouped[key].total += Number(r[6]);
  });
  sumSh.clearContents();
  sumSh.appendRow(['वर्ष','महिना','चालकाचे नाव','एकूण फेऱ्या','एकूण रक्कम (₹)']);
  sumSh.getRange(1,1,1,5).setBackground('#1a6b3a').setFontColor('white').setFontWeight('bold');
  Object.values(grouped).forEach(g => sumSh.appendRow([g.year,g.month,g.name,g.count,g.total]));
}

// ---------------------------------------------------------------
// SETUP — पहिल्यांदा run करा सर्व sheets तयार होण्यासाठी
// ---------------------------------------------------------------
function setupAllSheets() {
  ensureTripsSheet();
  ensureSettingsSheet();
  ensureUsersSheet();
  ensureStopsSheet();
  Logger.log('✅ सर्व sheets तयार झाल्या!');
  Logger.log('Sheets: ' + getSheetNames().join(', '));
}
