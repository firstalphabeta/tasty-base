// Bakery & Catering app backend. Attach this script to the bakery spreadsheet.
const SHEETS = {
  Settings: ['id', 'key', 'value'],
  Categories: ['id', 'name'],
  Items: ['id', 'name', 'category', 'price', 'cost', 'active', 'image'],
  Sales: ['id', 'date', 'itemId', 'name', 'category', 'qty', 'price', 'cost', 'customer', 'discount', 'receipt'],
  Expenses: ['id', 'date', 'category', 'description', 'amount'],
  Orders: ['id', 'created', 'customer', 'phone', 'details', 'pickup', 'total', 'deposit', 'status', 'cost', 'revenueDate']
};
const TEXT_COLS = { Sales: [2], Expenses: [2], Orders: [2, 4, 6, 11] };

function doGet() {
  return ContentService.createTextOutput('Bakery API is running');
}

function doPost(e) {
  let out;
  try {
    const request = JSON.parse(e.postData.contents || '{}');
    const actions = { getAll, save, saveMany, remove, getPinStatus };
    if (request.fn === 'setPin') {
      const args = Array.isArray(request.args) ? request.args : [];
      const currentPin = args[0] ?? request.pin ?? '';
      const nextPin = args[1] ?? '';
      out = { ok: true, data: setPin(String(currentPin), String(nextPin)) };
      return ContentService.createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (request.fn === 'getPinStatus') {
      out = { ok: true, data: getPinStatus() };
      return ContentService.createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (!Object.prototype.hasOwnProperty.call(actions, request.fn)) throw new Error('Unknown action');
    requirePin_(request.pin || '');
    out = { ok: true, data: actions[request.fn](...(request.args || [])) };
  } catch (error) {
    out = { ok: false, error: error.message || String(error) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  if (!Object.prototype.hasOwnProperty.call(SHEETS, name)) throw new Error('Unknown sheet');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    const headers = SHEETS[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    (TEXT_COLS[name] || []).forEach(column => sheet.getRange(1, column, sheet.getMaxRows(), 1).setNumberFormat('@'));
  } else {
    const headers = SHEETS[name];
    if (sheet.getLastColumn() < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

function seed_() {
  put_('Settings', { key: 'businessName', value: 'My Bakery' });
  put_('Settings', { key: 'ownerName', value: 'Boss' });
  put_('Settings', { key: 'currency', value: '$' });
  ['Cakes', 'Pastries', 'Drinks', 'Catering'].forEach(name => put_('Categories', { name }));
  put_('Items', { name: 'Vanilla Cupcake', category: 'Pastries', price: 5, cost: 1.5, active: true });
  put_('Items', { name: 'Red Velvet Cake', category: 'Cakes', price: 85, cost: 32, active: true });
  put_('Items', { name: 'Zobo / Fruit Drink', category: 'Drinks', price: 3, cost: 0.8, active: true });
}

function getAll() {
  const timezone = Session.getScriptTimeZone();
  const output = { _meta: { orderFinance: true } };
  Object.keys(SHEETS).forEach(name => {
    const sheet = sheet_(name), headers = SHEETS[name];
    if (name === 'Settings' && sheet.getLastRow() < 2) seed_();
    const lastRow = sheet.getLastRow();
    output[name] = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
      .filter(row => row[0] !== '')
      .map(row => Object.fromEntries(headers.map((header, index) => [
        header,
        row[index] instanceof Date ? Utilities.formatDate(row[index], timezone, 'yyyy-MM-dd') : row[index]
      ])));
  });
  return output;
}

function put_(name, record) {
  const sheet = sheet_(name), headers = SHEETS[name];
  if (record.id) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const index = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().indexOf(record.id);
      if (index > -1) {
        sheet.getRange(index + 2, 1, 1, headers.length).setValues([headers.map(header => record[header] ?? '')]);
        return record;
      }
    }
  }
  record.id = record.id || Utilities.getUuid().slice(0, 8);
  sheet.appendRow(headers.map(header => record[header] ?? ''));
  return record;
}

function locked_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function save(name, record) {
  return locked_(() => put_(name, record));
}

function saveMany(name, records) {
  return locked_(() => records.map(record => put_(name, record)));
}

function remove(name, id) {
  return locked_(() => {
    const sheet = sheet_(name), lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const index = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().indexOf(id);
    if (index > -1) sheet.deleteRow(index + 2);
    return index > -1;
  });
}

function pinHash_(pin, salt) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    salt + pin,
    Utilities.Charset.UTF_8
  );
  return digest.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
}

function requirePin_(pin) {
  const properties = PropertiesService.getScriptProperties();
  const savedHash = properties.getProperty('BAKERY_PIN_HASH');
  if (!savedHash) return;
  const salt = properties.getProperty('BAKERY_PIN_SALT') || '';
  if (pinHash_(String(pin), salt) !== savedHash) throw new Error('PIN_REQUIRED');
}

function getPinStatus() {
  const properties = PropertiesService.getScriptProperties();
  return {
    enabled: Boolean(properties.getProperty('BAKERY_PIN_HASH')),
    hint: properties.getProperty('BAKERY_PIN_HASH') ? 'PIN enabled' : 'PIN disabled'
  };
}

function setPin(currentPin, nextPin) {
  const properties = PropertiesService.getScriptProperties();
  const savedHash = properties.getProperty('BAKERY_PIN_HASH');
  if (savedHash) requirePin_(String(currentPin || ''));
  nextPin = String(nextPin || '');
  if (nextPin && !/^\d{4,12}$/.test(nextPin)) throw new Error('PIN must be 4 to 12 digits');
  if (!nextPin) {
    properties.deleteProperty('BAKERY_PIN_HASH');
    properties.deleteProperty('BAKERY_PIN_SALT');
    return { enabled: false };
  }
  const salt = Utilities.getUuid();
  properties.setProperties({ BAKERY_PIN_SALT: salt, BAKERY_PIN_HASH: pinHash_(nextPin, salt) });
  return { enabled: true };
}