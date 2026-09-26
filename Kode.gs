/**
 * Rhazes Nota - Google Apps Script backend
 * Database: Google Sheets | File output: Google Drive
 */

const APP = Object.freeze({
  NAME: 'Rhazes Nota',
  STORE_NAME: 'Rhazes Digital Store',
  DB_PROPERTY: 'RHAZES_DB_ID',
  FOLDER_PROPERTY: 'RHAZES_PDF_FOLDER_ID',
  TIMEZONE: 'Asia/Jakarta',
  SHEETS: {
    SETTINGS: 'Pengaturan',
    CUSTOMERS: 'Pelanggan',
    INVOICES: 'Invoice',
    ITEMS: 'Item Invoice'
  }
});

const HEADERS = Object.freeze({
  Pengaturan: ['Kunci', 'Nilai'],
  Pelanggan: ['ID', 'Nama', 'WhatsApp', 'Alamat', 'Terakhir Transaksi'],
  Invoice: [
    'ID', 'No Nota', 'Tanggal', 'Jatuh Tempo', 'Nama Pelanggan', 'WhatsApp',
    'Alamat', 'Subtotal', 'Diskon', 'Pajak', 'Ongkir', 'Total', 'Status',
    'Pembayaran', 'Tipe Nota', 'Catatan', 'Dibuat', 'PDF ID', 'PDF URL'
  ],
  'Item Invoice': ['Invoice ID', 'Nama Item', 'Qty', 'Harga', 'Jumlah']
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Rhazes Nota — Rhazes Digital Store')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
}

/** Jalankan sekali dari editor Apps Script sebelum deployment. */
function setupApplication() {
  const db = getDatabase_();
  const folder = getPdfFolder_();
  initializeDatabase_(db);
  return {
    success: true,
    databaseName: db.getName(),
    databaseUrl: db.getUrl(),
    pdfFolderName: folder.getName(),
    pdfFolderUrl: folder.getUrl()
  };
}

function getBootstrap() {
  try {
    const db = getDatabase_();
    initializeDatabase_(db);
    return {
      success: true,
      settings: readSettings_(db),
      invoices: listInvoices_(db, 100)
    };
  } catch (error) {
    return failure_(error);
  }
}

function saveSettings(input) {
  try {
    input = input || {};
    const db = getDatabase_();
    initializeDatabase_(db);
    const current = readSettings_(db);
    const settings = {
      storeName: clean_(input.storeName, 120) || APP.STORE_NAME,
      tagline: clean_(input.tagline, 160),
      address: clean_(input.address, 500),
      phone: normalizePhone_(input.phone),
      email: clean_(input.email, 160),
      logoUrl: cleanUrl_(input.logoUrl),
      bankName: clean_(input.bankName, 80),
      accountNumber: clean_(input.accountNumber, 80),
      accountName: clean_(input.accountName, 120),
      invoicePrefix: (clean_(input.invoicePrefix, 12) || current.invoicePrefix || 'RDS').toUpperCase(),
      defaultTax: number_(input.defaultTax, 0, 100),
      defaultNotes: clean_(input.defaultNotes, 600),
      themeColor: '#7c3aed'
    };

    const sheet = db.getSheetByName(APP.SHEETS.SETTINGS);
    const rows = Object.keys(settings).map(key => [key, String(settings[key] ?? '')]);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 2).setValues([HEADERS.Pengaturan]);
    if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
    styleHeader_(sheet, 2);
    return { success: true, settings: settings };
  } catch (error) {
    return failure_(error);
  }
}

function saveInvoice(input) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const db = getDatabase_();
    initializeDatabase_(db);
    const settings = readSettings_(db);
    const invoice = validateInvoice_(input);
    invoice.id = Utilities.getUuid();
    invoice.number = nextInvoiceNumber_(db, settings.invoicePrefix, invoice.date);
    invoice.createdAt = new Date();

    const subtotal = invoice.items.reduce((sum, item) => sum + item.total, 0);
    const discount = Math.min(invoice.discount, subtotal);
    const taxable = Math.max(0, subtotal - discount);
    const taxAmount = Math.round(taxable * invoice.tax / 100);
    const grandTotal = Math.max(0, taxable + taxAmount + invoice.shipping);

    invoice.subtotal = subtotal;
    invoice.discount = discount;
    invoice.taxAmount = taxAmount;
    invoice.total = grandTotal;

    const invoiceSheet = db.getSheetByName(APP.SHEETS.INVOICES);
    invoiceSheet.appendRow([
      invoice.id, invoice.number, invoice.date, invoice.dueDate || '', invoice.customerName,
      invoice.customerPhone, invoice.customerAddress, subtotal, discount, invoice.tax,
      invoice.shipping, grandTotal, invoice.status, invoice.paymentMethod, invoice.template,
      invoice.notes, invoice.createdAt, '', ''
    ]);

    const itemRows = invoice.items.map(item => [
      invoice.id, item.name, item.qty, item.price, item.total
    ]);
    const itemSheet = db.getSheetByName(APP.SHEETS.ITEMS);
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, 5).setValues(itemRows);
    upsertCustomer_(db, invoice);

    return {
      success: true,
      invoice: invoiceForClient_(invoice),
      invoices: listInvoices_(db, 100)
    };
  } catch (error) {
    return failure_(error);
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function getInvoice(invoiceId) {
  try {
    const db = getDatabase_();
    const invoice = findInvoice_(db, clean_(invoiceId, 100));
    if (!invoice) throw new Error('Nota tidak ditemukan.');
    return { success: true, invoice: invoiceForClient_(invoice) };
  } catch (error) {
    return failure_(error);
  }
}

function getInvoices() {
  try {
    const db = getDatabase_();
    return { success: true, invoices: listInvoices_(db, 100) };
  } catch (error) {
    return failure_(error);
  }
}

function updateInvoiceStatus(invoiceId, status) {
  try {
    const allowed = ['Belum Lunas', 'Lunas', 'Dibatalkan'];
    if (!allowed.includes(status)) throw new Error('Status tidak valid.');
    const db = getDatabase_();
    const sheet = db.getSheetByName(APP.SHEETS.INVOICES);
    const values = sheet.getDataRange().getValues();
    const rowIndex = values.findIndex((row, index) => index > 0 && String(row[0]) === String(invoiceId));
    if (rowIndex < 0) throw new Error('Nota tidak ditemukan.');

    const oldPdfId = String(values[rowIndex][17] || '');
    sheet.getRange(rowIndex + 1, 13).setValue(status);
    sheet.getRange(rowIndex + 1, 18, 1, 2).clearContent();
    if (oldPdfId) {
      try { DriveApp.getFileById(oldPdfId).setTrashed(true); } catch (ignored) {}
    }
    return { success: true, invoices: listInvoices_(db, 100) };
  } catch (error) {
    return failure_(error);
  }
}

function generateInvoicePdf(invoiceId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const db = getDatabase_();
    const invoice = findInvoice_(db, clean_(invoiceId, 100));
    if (!invoice) throw new Error('Nota tidak ditemukan.');

    if (invoice.pdfId) {
      try {
        const existing = DriveApp.getFileById(invoice.pdfId);
        if (!existing.isTrashed()) {
          return { success: true, url: invoice.pdfUrl || existing.getUrl(), fileName: existing.getName() };
        }
      } catch (ignored) {}
    }

    const settings = readSettings_(db);
    const folder = getPdfFolder_();
    const tempDoc = DocumentApp.create('TEMP-' + invoice.number);
    const docFile = DriveApp.getFileById(tempDoc.getId());
    const body = tempDoc.getBody();
    body.setMarginTop(32).setMarginBottom(32).setMarginLeft(36).setMarginRight(36);
    buildPdfDocument_(body, invoice, settings);
    tempDoc.saveAndClose();

    const pdfName = invoice.number + ' - ' + safeFileName_(invoice.customerName) + '.pdf';
    const pdfFile = folder.createFile(docFile.getAs(MimeType.PDF).setName(pdfName));
    docFile.setTrashed(true);

    let url = pdfFile.getUrl();
    try {
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      url = 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view';
    } catch (ignored) {
      // Beberapa akun Workspace melarang berbagi publik; URL tetap dapat dibuka pemilik.
    }

    const sheet = db.getSheetByName(APP.SHEETS.INVOICES);
    const values = sheet.getDataRange().getValues();
    const rowIndex = values.findIndex((row, index) => index > 0 && String(row[0]) === invoice.id);
    if (rowIndex >= 0) sheet.getRange(rowIndex + 1, 18, 1, 2).setValues([[pdfFile.getId(), url]]);

    return { success: true, url: url, fileName: pdfName };
  } catch (error) {
    return failure_(error);
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function getDatabase_() {
  const properties = PropertiesService.getScriptProperties();
  const storedId = properties.getProperty(APP.DB_PROPERTY);
  if (storedId) {
    try { return SpreadsheetApp.openById(storedId); } catch (ignored) {}
  }
  const db = SpreadsheetApp.create('Database - Rhazes Nota');
  properties.setProperty(APP.DB_PROPERTY, db.getId());
  return db;
}

function getPdfFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const storedId = properties.getProperty(APP.FOLDER_PROPERTY);
  if (storedId) {
    try { return DriveApp.getFolderById(storedId); } catch (ignored) {}
  }
  const folder = DriveApp.createFolder('PDF Nota - Rhazes Digital Store');
  properties.setProperty(APP.FOLDER_PROPERTY, folder.getId());
  return folder;
}

function initializeDatabase_(db) {
  const first = db.getSheets()[0];
  if (first && first.getName() === 'Sheet1' && db.getSheets().length === 1) {
    first.setName(APP.SHEETS.SETTINGS);
  }

  Object.keys(HEADERS).forEach(name => {
    let sheet = db.getSheetByName(name);
    if (!sheet) sheet = db.insertSheet(name);
    const headers = HEADERS[name];
    if (sheet.getLastRow() === 0 || String(sheet.getRange(1, 1).getValue()) !== headers[0]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sheet.setFrozenRows(1);
    styleHeader_(sheet, headers.length);
  });

  const settingsSheet = db.getSheetByName(APP.SHEETS.SETTINGS);
  if (settingsSheet.getLastRow() < 2) {
    const defaults = defaultSettings_();
    const rows = Object.keys(defaults).map(key => [key, String(defaults[key])]);
    settingsSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  const invoiceSheet = db.getSheetByName(APP.SHEETS.INVOICES);
  invoiceSheet.getRange('C:D').setNumberFormat('dd/MM/yyyy');
  invoiceSheet.getRange('H:L').setNumberFormat('Rp #,##0');
  invoiceSheet.autoResizeColumns(1, HEADERS.Invoice.length);

  const itemSheet = db.getSheetByName(APP.SHEETS.ITEMS);
  itemSheet.getRange('C:C').setNumberFormat('0.##');
  itemSheet.getRange('D:E').setNumberFormat('Rp #,##0');
}

function defaultSettings_() {
  return {
    storeName: APP.STORE_NAME,
    tagline: 'Solusi Produk Digital Terpercaya',
    address: '',
    phone: '',
    email: '',
    logoUrl: '',
    bankName: '',
    accountNumber: '',
    accountName: '',
    invoicePrefix: 'RDS',
    defaultTax: 0,
    defaultNotes: 'Terima kasih telah berbelanja di Rhazes Digital Store.',
    themeColor: '#7c3aed'
  };
}

function readSettings_(db) {
  const settings = defaultSettings_();
  const sheet = db.getSheetByName(APP.SHEETS.SETTINGS);
  if (sheet.getLastRow() < 2) return settings;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(row => {
    const key = String(row[0] || '');
    if (key) settings[key] = row[1];
  });
  settings.defaultTax = number_(settings.defaultTax, 0, 100);
  return settings;
}

function validateInvoice_(input) {
  input = input || {};
  const rawItems = Array.isArray(input.items) ? input.items : [];
  const items = rawItems.map(item => {
    const name = clean_(item.name, 180);
    const qty = number_(item.qty, 0.01, 999999);
    const price = Math.round(number_(item.price, 0, 999999999999));
    return { name: name, qty: qty, price: price, total: Math.round(qty * price) };
  }).filter(item => item.name && item.qty > 0);

  if (!clean_(input.customerName, 160)) throw new Error('Nama pelanggan wajib diisi.');
  if (!items.length) throw new Error('Tambahkan minimal satu item.');
  if (items.length > 100) throw new Error('Maksimal 100 item per nota.');

  const template = ['simple', 'professional'].includes(input.template) ? input.template : 'professional';
  const status = ['Belum Lunas', 'Lunas'].includes(input.status) ? input.status : 'Belum Lunas';
  return {
    date: parseDate_(input.date) || new Date(),
    dueDate: parseDate_(input.dueDate),
    customerName: clean_(input.customerName, 160),
    customerPhone: normalizePhone_(input.customerPhone),
    customerAddress: clean_(input.customerAddress, 500),
    items: items,
    discount: Math.round(number_(input.discount, 0, 999999999999)),
    tax: number_(input.tax, 0, 100),
    shipping: Math.round(number_(input.shipping, 0, 999999999999)),
    status: status,
    paymentMethod: clean_(input.paymentMethod, 80),
    template: template,
    notes: clean_(input.notes, 1000)
  };
}

function nextInvoiceNumber_(db, prefix, date) {
  const datePart = Utilities.formatDate(date, APP.TIMEZONE, 'yyyyMMdd');
  const base = (prefix || 'RDS') + '-' + datePart + '-';
  const sheet = db.getSheetByName(APP.SHEETS.INVOICES);
  let max = 0;
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => {
      const value = String(row[0] || '');
      if (value.indexOf(base) === 0) max = Math.max(max, Number(value.slice(base.length)) || 0);
    });
  }
  return base + String(max + 1).padStart(3, '0');
}

function listInvoices_(db, limit) {
  const sheet = db.getSheetByName(APP.SHEETS.INVOICES);
  if (sheet.getLastRow() < 2) return [];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.Invoice.length).getValues();
  return rows.slice(-limit).reverse().map(row => ({
    id: String(row[0]),
    number: String(row[1]),
    date: dateString_(row[2]),
    customerName: String(row[4]),
    total: number_(row[11], 0, Number.MAX_SAFE_INTEGER),
    status: String(row[12]),
    template: String(row[14]),
    pdfUrl: String(row[18] || '')
  }));
}

function findInvoice_(db, invoiceId) {
  const sheet = db.getSheetByName(APP.SHEETS.INVOICES);
  if (sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.Invoice.length).getValues();
  const row = rows.find(data => String(data[0]) === String(invoiceId));
  if (!row) return null;

  const itemSheet = db.getSheetByName(APP.SHEETS.ITEMS);
  const itemRows = itemSheet.getLastRow() > 1
    ? itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, 5).getValues()
    : [];
  const items = itemRows.filter(item => String(item[0]) === String(invoiceId)).map(item => ({
    name: String(item[1]), qty: Number(item[2]), price: Number(item[3]), total: Number(item[4])
  }));

  const subtotal = Number(row[7]) || 0;
  const discount = Number(row[8]) || 0;
  const taxPercent = Number(row[9]) || 0;
  return {
    id: String(row[0]), number: String(row[1]), date: row[2], dueDate: row[3],
    customerName: String(row[4]), customerPhone: String(row[5]), customerAddress: String(row[6]),
    subtotal: subtotal, discount: discount, tax: taxPercent,
    taxAmount: Math.round(Math.max(0, subtotal - discount) * taxPercent / 100),
    shipping: Number(row[10]) || 0, total: Number(row[11]) || 0,
    status: String(row[12]), paymentMethod: String(row[13]), template: String(row[14]),
    notes: String(row[15]), createdAt: row[16], pdfId: String(row[17] || ''),
    pdfUrl: String(row[18] || ''), items: items
  };
}

function invoiceForClient_(invoice) {
  return Object.assign({}, invoice, {
    date: dateString_(invoice.date),
    dueDate: dateString_(invoice.dueDate),
    createdAt: dateTimeString_(invoice.createdAt)
  });
}

function upsertCustomer_(db, invoice) {
  const sheet = db.getSheetByName(APP.SHEETS.CUSTOMERS);
  const rows = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
    : [];
  const phone = invoice.customerPhone;
  const index = rows.findIndex(row =>
    (phone && String(row[2]) === phone) || (!phone && String(row[1]).toLowerCase() === invoice.customerName.toLowerCase())
  );
  const values = [Utilities.getUuid(), invoice.customerName, phone, invoice.customerAddress, invoice.date];
  if (index >= 0) {
    values[0] = rows[index][0];
    sheet.getRange(index + 2, 1, 1, 5).setValues([values]);
  } else {
    sheet.appendRow(values);
  }
}

function buildPdfDocument_(body, invoice, settings) {
  const purple = '#6D28D9';
  const dark = '#17121F';
  const gray = '#6B7280';
  const professional = invoice.template === 'professional';

  const header = body.appendTable([[settings.storeName || APP.STORE_NAME, professional ? 'INVOICE' : 'NOTA']]);
  header.setBorderWidth(0);
  header.getCell(0, 0).getChild(0).asParagraph().setHeading(DocumentApp.ParagraphHeading.HEADING1);
  header.getCell(0, 0).editAsText().setForegroundColor(professional ? purple : dark).setBold(true);
  header.getCell(0, 1).editAsText().setForegroundColor(professional ? purple : dark).setBold(true).setFontSize(18);
  header.getCell(0, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);

  if (settings.tagline) {
    body.appendParagraph(String(settings.tagline)).editAsText().setForegroundColor(gray).setFontSize(9);
  }
  const contact = [settings.address, settings.phone, settings.email].filter(Boolean).join(' • ');
  if (contact) body.appendParagraph(contact).editAsText().setForegroundColor(gray).setFontSize(8);
  body.appendHorizontalRule();

  const info = body.appendTable([
    ['DITAGIHKAN KEPADA', 'DETAIL NOTA'],
    [invoice.customerName, 'No: ' + invoice.number],
    [invoice.customerPhone || '-', 'Tanggal: ' + dateDisplay_(invoice.date)],
    [invoice.customerAddress || '-', 'Status: ' + invoice.status]
  ]);
  info.setBorderWidth(0);
  for (let c = 0; c < 2; c++) {
    info.getCell(0, c).editAsText().setBold(true).setForegroundColor(professional ? purple : dark).setFontSize(9);
  }
  for (let r = 1; r < info.getNumRows(); r++) {
    for (let c = 0; c < 2; c++) info.getCell(r, c).editAsText().setFontSize(9);
  }
  body.appendParagraph('');

  const itemRows = [['ITEM', 'QTY', 'HARGA', 'JUMLAH']];
  invoice.items.forEach(item => itemRows.push([
    item.name, formatNumber_(item.qty), rupiah_(item.price), rupiah_(item.total)
  ]));
  const table = body.appendTable(itemRows);
  for (let c = 0; c < 4; c++) {
    table.getCell(0, c).setBackgroundColor(professional ? purple : dark);
    table.getCell(0, c).editAsText().setBold(true).setForegroundColor('#FFFFFF').setFontSize(9);
  }
  for (let r = 1; r < table.getNumRows(); r++) {
    for (let c = 0; c < 4; c++) table.getCell(r, c).editAsText().setFontSize(9);
  }

  body.appendParagraph('');
  const totals = [
    ['Subtotal', rupiah_(invoice.subtotal)],
    ['Diskon', '- ' + rupiah_(invoice.discount)],
    ['Pajak (' + formatNumber_(invoice.tax) + '%)', rupiah_(invoice.taxAmount)],
    ['Ongkir/Biaya lain', rupiah_(invoice.shipping)],
    ['TOTAL', rupiah_(invoice.total)]
  ];
  const totalTable = body.appendTable(totals);
  totalTable.setBorderWidth(0);
  for (let r = 0; r < totalTable.getNumRows(); r++) {
    totalTable.getCell(r, 1).getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    for (let c = 0; c < 2; c++) {
      totalTable.getCell(r, c).editAsText().setFontSize(r === totals.length - 1 ? 12 : 9);
    }
  }
  for (let c = 0; c < 2; c++) {
    totalTable.getCell(totals.length - 1, c).editAsText().setBold(true).setForegroundColor(professional ? purple : dark);
  }

  if (invoice.paymentMethod) {
    body.appendParagraph('Metode pembayaran: ' + invoice.paymentMethod).editAsText().setFontSize(9).setBold(true);
  }
  if (settings.bankName || settings.accountNumber) {
    body.appendParagraph(['Pembayaran', settings.bankName, settings.accountNumber, settings.accountName].filter(Boolean).join(' • '))
      .editAsText().setFontSize(9).setForegroundColor(gray);
  }
  if (invoice.notes) {
    body.appendParagraph('Catatan: ' + invoice.notes).editAsText().setFontSize(9).setForegroundColor(gray);
  }
  const footer = body.appendParagraph('Terima kasih atas kepercayaan Anda.');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.editAsText().setForegroundColor(professional ? purple : dark).setBold(true).setFontSize(9);
}

function styleHeader_(sheet, columns) {
  sheet.getRange(1, 1, 1, columns)
    .setBackground('#2E1065').setFontColor('#FFFFFF').setFontWeight('bold');
}

function parseDate_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  const match = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0);
}

function dateString_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date)) return '';
  return Utilities.formatDate(date, APP.TIMEZONE, 'yyyy-MM-dd');
}

function dateTimeString_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date)) return '';
  return Utilities.formatDate(date, APP.TIMEZONE, 'dd/MM/yyyy HH:mm');
}

function dateDisplay_(value) {
  if (!value) return '-';
  const date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, APP.TIMEZONE, 'dd MMMM yyyy');
}

function clean_(value, maxLength) {
  return String(value == null ? '' : value).replace(/[\u0000-\u001F\u007F]/g, ' ').trim().slice(0, maxLength || 500);
}

function cleanUrl_(value) {
  const url = clean_(value, 500);
  if (!url) return '';
  if (!/^https:\/\//i.test(url)) throw new Error('URL logo harus diawali https://');
  return url;
}

function normalizePhone_(value) {
  let phone = clean_(value, 30).replace(/[^\d+]/g, '');
  if (phone.indexOf('+62') === 0) phone = '62' + phone.slice(3);
  if (phone.indexOf('0') === 0) phone = '62' + phone.slice(1);
  return phone;
}

function number_(value, min, max) {
  const number = Number(value);
  if (!isFinite(number)) return min || 0;
  return Math.min(max, Math.max(min, number));
}

function rupiah_(value) {
  return 'Rp ' + Math.round(Number(value) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function formatNumber_(value) {
  return Number(value || 0).toLocaleString('id-ID');
}

function safeFileName_(value) {
  return clean_(value, 80).replace(/[\\/:*?\"<>|]/g, '-');
}

function failure_(error) {
  console.error(error && error.stack ? error.stack : error);
  return { success: false, message: error && error.message ? error.message : 'Terjadi kesalahan pada server.' };
}
