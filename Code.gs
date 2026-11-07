const SHEMAS = {
  Users: ['id', 'created_at', 'name', 'currency', 'timezone', 'fiscal_year_start', 'updated_at'],
  Transactions: ['id', 'created_at', 'updated_at', 'type', 'amount', 'category_id', 'transaction_date', 'note', 'source', 'created_by'],
  Categories: ['id', 'created_at', 'updated_at', 'name', 'type', 'icon', 'color', 'is_default', 'is_active', 'sort_order'],
  Accounts: ['id', 'created_at', 'updated_at', 'name', 'is_active', 'account_number'],
  AppSettings: ['id', 'created_at', 'key', 'value', 'updated_at'],
  AuditLogs: ['id', 'created_at', 'action', 'entity_type', 'entity_id', 'details']
};

const DEFAULT_CATEGORIES = [
  {name:'Makanan',type:'expense',icon:'restaurant',color:'#238467'},
  {name:'Transportasi',type:'expense',icon:'directions_car',color:'#5689a2'},
  {name:'Tagihan',type:'expense',icon:'receipt_long',color:'#8d82bf'},
  {name:'Belanja',type:'expense',icon:'shopping_bag',color:'#e0a36e'},
  {name:'Hiburan',type:'expense',icon:'movie',color:'#ce7e72'},
  {name:'Kesehatan',type:'expense',icon:'medical_services',color:'#6eaa83'},
  {name:'Lainnya',type:'expense',icon:'category',color:'#889790'},
  {name:'Gaji',type:'income',icon:'payments',color:'#238467'},
  {name:'Bonus',type:'income',icon:'redeem',color:'#e0a36e'},
  {name:'Freelance',type:'income',icon:'work',color:'#5689a2'},
  {name:'Hasil usaha',type:'income',icon:'storefront',color:'#8d82bf'},
  {name:'Investasi',type:'income',icon:'trending_up',color:'#6eaa83'},
  {name:'Lainnya',type:'income',icon:'category',color:'#889790'}
];

function doGet(e) {
  if (e && e.parameter && e.parameter.api === '1') {
    return _jsonpApi(e);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('RuangUang — Keuangan Pribadi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _jsonpApi(e) {
  try {
    const callback = String(e.parameter.callback || '');
    if (!/^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback.')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    const routes = {
      loginUser: loginUser,
      registerUser: registerUser,
      getInitialData: getInitialData,
      getDashboard: getDashboard,
      getTransactions: getTransactions,
      getTransactionById: getTransactionById,
      createTransaction: createTransaction,
      updateTransaction: updateTransaction,
      deleteTransaction: deleteTransaction,
      getCategories: getCategories,
      getAccounts: getAccounts,
      createAccount: createAccount,
      updateAccount: updateAccount,
      createCategory: createCategory,
      updateCategory: updateCategory,
      deleteCategory: deleteCategory,
      getReport: getReport,
      getUserSettings: getUserSettings,
      updateUserSettings: updateUserSettings,
      exportTransactions: exportTransactions,
      getAuditLogs: getAuditLogs
    };
    const method = String(e.parameter.method || '');
    const args = JSON.parse(String(e.parameter.args || '[]'));
    if (!Array.isArray(args) || !routes[method]) {
      throw new Error('Permintaan tidak valid.');
    }
    const result = routes[method].apply(null, args);
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (err) {
    const callback = String(e.parameter.callback || '');
    if (!/^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback.')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    const result = _fail(err);
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function setupSpreadsheet() {
  try {
    const ss = _databaseSpreadsheet();
    Object.keys(SHEMAS).forEach(function(name) {
      let sheet = ss.getSheetByName(name) || ss.insertSheet(name);
      const headers = SHEMAS[name];
      const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      const hasHeaders = firstRow.some(function(v) { return String(v).trim() !== ''; });
      if (!hasHeaders) {
        sheet.appendRow(headers);
      } else if (headers.some(function(h, i) { return String(firstRow[i] || '').trim() !== h; })) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#0d6e60')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.setRowHeight(1, 32);
    });
    const users = ss.getSheetByName('Users');
    if (users.getLastRow() < 2) {
      const now = new Date();
      users.appendRow([_uuid(), now, 'Saya', 'IDR', 'Asia/Jakarta', 1, now]);
    }
    const categories = ss.getSheetByName('Categories');
    if (categories.getLastRow() < 2) {
      const now = new Date();
      DEFAULT_CATEGORIES.forEach(function(c, i) {
        categories.appendRow([_uuid(), now, now, c.name, c.type, c.icon, c.color, true, true, i + 1]);
      });
    }
    const accounts = ss.getSheetByName('Accounts');
    if (accounts.getLastRow() < 2) {
      const now = new Date();
      accounts.appendRow([_uuid(), now, now, 'Tunai', true]);
    }
    const settings = ss.getSheetByName('AppSettings');
    if (settings.getLastRow() < 2) {
      const now = new Date();
      settings.appendRow([_uuid(), now, 'initialized', 'true', now]);
    }
    return {success:true, message:'Database siap.', data:{}};
  } catch (err) {
    return _fail(err);
  }
}

function getInitialData(token) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const settings = _settings();
    const categories = _categories(null, true);
    const now = new Date();
    const month = now.getMonth() + 1;
    const year = now.getFullYear();
    const dashboard = _dashboard(month, year);
    const tx = _filteredTransactions({month:month, year:year, limit:8, sortBy:'transactionDate', sortDirection:'desc'});
    return _ok('Data awal berhasil dimuat.', {
      settings:settings,
      categories:categories,
      transactions:tx.items,
      dashboard:dashboard
    });
  } catch (err) {
    return _fail(err);
  }
}

function getDashboard(token, month, year) {
  try {
    _useSession(token);
    setupSpreadsheet();
    return _ok('Ringkasan berhasil dimuat.', _dashboard(_int(month, new Date().getMonth()+1), _int(year, new Date().getFullYear())));
  } catch (err) {
    return _fail(err);
  }
}

function getTransactions(token, filters) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const result = _filteredTransactions(filters || {});
    return _ok('Transaksi berhasil dimuat.', result);
  } catch (err) {
    return _fail(err);
  }
}

function getTransactionById(token, id) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const record = _findById('Transactions', id);
    if (!record) throw new Error('Transaksi tidak ditemukan.');
    return _ok('Detail transaksi berhasil dimuat.', _transactionObject(record));
  } catch (err) {
    return _fail(err);
  }
}

function createTransaction(token, payload) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const clean = _validateTransaction(payload);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const now = new Date();
      const id = _uuid();
      const row = [id, now, now, clean.type, clean.amount, clean.categoryId, clean.transactionDate, clean.note, clean.source, 'single-user'];
      _sheet('Transactions').appendRow(row);
      _audit('create', 'transaction', id, {type:clean.type, amount:clean.amount, categoryId:clean.categoryId});
      return _ok('Transaksi berhasil ditambahkan.', {
        id:id,
        transaction:_transactionObject(_rowObject(SHEMAS.Transactions, row)),
        dashboard:_dashboard(_monthFromDate(clean.transactionDate), _yearFromDate(clean.transactionDate))
      });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return _fail(err);
  }
}

function updateTransaction(token, id, payload) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const clean = _validateTransaction(payload);
    const sheet = _sheet('Transactions');
    const rowNumber = _rowNumberById(sheet, id);
    if (!rowNumber) throw new Error('Transaksi tidak ditemukan.');
    const previous = _findById('Transactions', id);
    const now = new Date();
    const row = [id, previous.created_at, now, clean.type, clean.amount, clean.categoryId, clean.transactionDate, clean.note, clean.source, previous.created_by || 'single-user'];
    sheet.getRange(rowNumber, 1, 1, SHEMAS.Transactions.length).setValues([row]);
    _audit('update', 'transaction', id, {before:{type:previous.type,amount:previous.amount},after:{type:clean.type,amount:clean.amount}});
    return _ok('Transaksi berhasil diperbarui.', {
      id:id,
      transaction:_transactionObject(_rowObject(SHEMAS.Transactions, row)),
      dashboard:_dashboard(_monthFromDate(clean.transactionDate), _yearFromDate(clean.transactionDate))
    });
  } catch (err) {
    return _fail(err);
  }
}

function deleteTransaction(token, id) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const sheet = _sheet('Transactions');
    const rowNumber = _rowNumberById(sheet, id);
    if (!rowNumber) throw new Error('Transaksi tidak ditemukan.');
    const record = _findById('Transactions', id);
    sheet.deleteRow(rowNumber);
    _audit('delete', 'transaction', id, {type:record.type, amount:record.amount});
    return _ok('Transaksi berhasil dihapus.', {
      deletedId:id,
      dashboard:_dashboard(_monthFromDate(record.transaction_date), _yearFromDate(record.transaction_date))
    });
  } catch (err) {
    return _fail(err);
  }
}

function getCategories(token, type) {
  try {
    _useSession(token);
    setupSpreadsheet();
    return _ok('Kategori berhasil dimuat.', {categories:_categories(type || null, false)});
  } catch (err) {
    return _fail(err);
  }
}

function getAccounts(token) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const accounts = _accounts(false);
    const transactions = _read('Transactions').map(_transactionObject);
    accounts.forEach(function(account) {
      account.balance = transactions.reduce(function(total, transaction) {
        if (transaction.source !== account.id) return total;
        return total + (transaction.type === 'income' ? transaction.amount : -transaction.amount);
      }, 0);
    });
    return _ok('Rekening berhasil dimuat.', {accounts:accounts});
  } catch (err) {
    return _fail(err);
  }
}

function createAccount(token, payload) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const data = payload || {};
    const name = String(data.name || '').trim();
    const accountNumber = String(data.accountNumber || '').trim();
    if (!name || name.length > 60) throw new Error('Nama rekening wajib diisi (maksimal 60 karakter).');
    if (accountNumber.length > 40) throw new Error('Nomor rekening maksimal 40 karakter.');
    if (_accounts(false).some(function(a) { return a.name.toLowerCase() === name.toLowerCase(); })) {
      throw new Error('Nama rekening tersebut sudah digunakan.');
    }
    const now = new Date();
    const id = _uuid();
    const row = [id, now, now, name, true, accountNumber];
    _sheet('Accounts').appendRow(row);
    _audit('create', 'account', id, {name:name});
    return _ok('Rekening berhasil ditambahkan.', {account:_accountObject(_rowObject(SHEMAS.Accounts, row))});
  } catch (err) {
    return _fail(err);
  }
}

function updateAccount(token, id, payload) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const sheet = _sheet('Accounts');
    const rowNumber = _rowNumberById(sheet, id);
    if (!rowNumber) throw new Error('Rekening tidak ditemukan.');
    const old = _findById('Accounts', id);
    const data = payload || {};
    const name = data.name === undefined ? String(old.name) : String(data.name).trim();
    const accountNumber = data.accountNumber === undefined ? String(old.account_number || '') : String(data.accountNumber).trim();
    if (!name || name.length > 60) throw new Error('Nama rekening wajib diisi (maksimal 60 karakter).');
    if (accountNumber.length > 40) throw new Error('Nomor rekening maksimal 40 karakter.');
    if (_accounts(false).some(function(a) { return a.id !== String(id) && a.name.toLowerCase() === name.toLowerCase(); })) {
      throw new Error('Nama rekening tersebut sudah digunakan.');
    }
    const now = new Date();
    const row = [String(id), old.created_at, now, name,
      data.isActive === undefined ? _bool(old.is_active) : Boolean(data.isActive), accountNumber];
    sheet.getRange(rowNumber, 1, 1, SHEMAS.Accounts.length).setValues([row]);
    _audit('update', 'account', id, {name:name, isActive:row[4]});
    return _ok('Rekening berhasil diperbarui.', {account:_accountObject(_rowObject(SHEMAS.Accounts, row))});
  } catch (err) {
    return _fail(err);
  }
}

function createCategory(token, payload) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const data = payload || {};
    const name = String(data.name || '').trim();
    const type = String(data.type || '');
    if (!name || name.length > 60) throw new Error('Nama kategori wajib diisi (maksimal 60 karakter).');
    if (type !== 'income' && type !== 'expense') throw new Error('Jenis kategori tidak valid.');
    const duplicate = _categories(null, false).some(function(c) {
      return c.type === type && c.name.toLowerCase() === name.toLowerCase();
    });
    if (duplicate) throw new Error('Nama kategori tersebut sudah digunakan.');
    const now = new Date();
    const id = _uuid();
    const row = [id, now, now, name, type, _safeIcon(data.icon), _safeColor(data.color), false, true, _sheet('Categories').getLastRow()];
    _sheet('Categories').appendRow(row);
    _audit('create', 'category', id, {name:name, type:type});
    return _ok('Kategori berhasil dibuat.', {id:id, category:_categoryObject(_rowObject(SHEMAS.Categories, row))});
  } catch (err) {
    return _fail(err);
  }
}

function updateCategory(token, id, payload) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const sheet = _sheet('Categories');
    const rowNumber = _rowNumberById(sheet, id);
    if (!rowNumber) throw new Error('Kategori tidak ditemukan.');
    const old = _findById('Categories', id);
    const data = payload || {};
    const name = data.name === undefined ? old.name : String(data.name).trim();
    if (!name || name.length > 60) throw new Error('Nama kategori wajib diisi (maksimal 60 karakter).');
    const duplicate = _categories(null, false).some(function(c) {
      return c.id !== id && c.type === old.type && c.name.toLowerCase() === name.toLowerCase();
    });
    if (duplicate) throw new Error('Nama kategori tersebut sudah digunakan.');
    const now = new Date();
    const row = [id, old.created_at, now, name, old.type, data.icon === undefined ? old.icon : _safeIcon(data.icon),
      data.color === undefined ? old.color : _safeColor(data.color), _bool(old.is_default),
      data.isActive === undefined ? _bool(old.is_active) : Boolean(data.isActive), old.sort_order];
    sheet.getRange(rowNumber, 1, 1, SHEMAS.Categories.length).setValues([row]);
    _audit('update', 'category', id, {name:name, isActive:row[8]});
    return _ok('Kategori berhasil diperbarui.', {id:id, category:_categoryObject(_rowObject(SHEMAS.Categories, row))});
  } catch (err) {
    return _fail(err);
  }
}

function deleteCategory(token, id) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const sheet = _sheet('Categories');
    const rowNumber = _rowNumberById(sheet, id);
    if (!rowNumber) throw new Error('Kategori tidak ditemukan.');
    const category = _findById('Categories', id);
    if (_bool(category.is_default)) throw new Error('Kategori bawaan tidak dapat dihapus. Nonaktifkan kategori jika tidak ingin menggunakannya.');
    const used = _read('Transactions').some(function(t) { return String(t.category_id) === String(id); });
    if (used) throw new Error('Kategori sudah digunakan pada transaksi. Nonaktifkan agar riwayat tetap terjaga.');
    sheet.deleteRow(rowNumber);
    _audit('delete', 'category', id, {name:category.name});
    return _ok('Kategori berhasil dihapus.', {deletedId:id});
  } catch (err) {
    return _fail(err);
  }
}

function getReport(token, month, year, date) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const m = _int(month, new Date().getMonth()+1);
    const y = _int(year, new Date().getFullYear());
    const dailyDate = String(date || '').trim();
    const filters = dailyDate ? {date:dailyDate, limit:5000, sortBy:'transactionDate', sortDirection:'desc'} :
      {month:m, year:y, limit:5000, sortBy:'transactionDate', sortDirection:'desc'};
    const transactions = _filteredTransactions(filters).items;
    const data = _dashboard(m, y);
    const income = _sum(transactions, 'income');
    const expense = _sum(transactions, 'expense');
    return _ok('Laporan berhasil dimuat.', {
      summary:{
        balance:data.balance,
        income:income,
        expense:expense,
        savingsRate:income ? Math.round(((income-expense)/income)*100) : 0,
        transactionCount:transactions.length
      },
      categoryBreakdown:data.categoryBreakdown,
      transactions:transactions
    });
  } catch (err) {
    return _fail(err);
  }
}

function getUserSettings(token) {
  try {
    _useSession(token);
    setupSpreadsheet();
    return _ok('Pengaturan berhasil dimuat.', {settings:_settings()});
  } catch (err) {
    return _fail(err);
  }
}

function updateUserSettings(token, payload) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const current = _settings();
    const data = payload || {};
    const name = data.name === undefined ? current.name : String(data.name).trim();
    const currency = data.currency === undefined ? current.currency : String(data.currency).toUpperCase();
    const timezone = data.timezone === undefined ? current.timezone : String(data.timezone);
    const fiscal = data.fiscalYearStart === undefined ? current.fiscalYearStart : _int(data.fiscalYearStart, 1);
    if (!name || name.length > 80) throw new Error('Nama perlu diisi (maksimal 80 karakter).');
    if (!/^[A-Z]{3}$/.test(currency)) throw new Error('Kode mata uang tidak valid.');
    if (fiscal < 1 || fiscal > 12) throw new Error('Awal tahun anggaran harus antara 1 sampai 12.');
    const sheet = _sheet('Users');
    const rowNumber = _rowNumberById(sheet, current.id);
    const now = new Date();
    const row = [current.id, current.createdAt || now, name, currency, timezone, fiscal, now];
    if (rowNumber) sheet.getRange(rowNumber, 1, 1, SHEMAS.Users.length).setValues([row]);
    else sheet.appendRow(row);
    _audit('update', 'user_settings', current.id, {currency:currency, fiscalYearStart:fiscal});
    return _ok('Pengaturan berhasil diperbarui.', {settings:_settings()});
  } catch (err) {
    return _fail(err);
  }
}

function exportTransactions(token, month, year, format, date) {
  try {
    _useSession(token);
    setupSpreadsheet();
    if (format && String(format).toLowerCase() !== 'csv') throw new Error('Format ekspor yang didukung adalah CSV.');
    const m = _int(month, new Date().getMonth()+1);
    const y = _int(year, new Date().getFullYear());
    const dailyDate = String(date || '').trim();
    const filters = dailyDate ? {date:dailyDate, limit:50000, sortBy:'transactionDate', sortDirection:'asc'} :
      {month:m, year:y, limit:50000, sortBy:'transactionDate', sortDirection:'asc'};
    const items = _filteredTransactions(filters).items;
    const rows = [['Tanggal','Jenis','Kategori','Nominal','Catatan']];
    items.forEach(function(t) {
      rows.push([t.transactionDate,t.type === 'income' ? 'Pemasukan' : 'Pengeluaran',t.categoryName,t.amount,t.note]);
    });
    const content = rows.map(function(row) { return row.map(_csvCell).join(','); }).join('\r\n');
    return _ok('Data ekspor berhasil disiapkan.', {
      format:'csv',
      filename:dailyDate ? 'laporan-keuangan-' + dailyDate + '.csv' : 'laporan-keuangan-' + y + '-' + ('0'+m).slice(-2) + '.csv',
      content:content
    });
  } catch (err) {
    return _fail(err);
  }
}

function getAuditLogs(token, limit) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const max = Math.min(500, Math.max(1, _int(limit, 100)));
    const logs = _read('AuditLogs').sort(function(a,b) { return new Date(b.created_at)-new Date(a.created_at); }).slice(0,max).map(function(x) {
      return {id:String(x.id),createdAt:_dateTime(x.created_at),action:String(x.action),entityType:String(x.entity_type),entityId:String(x.entity_id),details:_parseJson(x.details)};
    });
    return _ok('Log aktivitas berhasil dimuat.', {logs:logs});
  } catch (err) {
    return _fail(err);
  }
}

function resetApplicationData(token) {
  try {
    _useSession(token);
    setupSpreadsheet();
    const sheet = _sheet('Transactions');
    const count = Math.max(0, sheet.getLastRow()-1);
    if (count) sheet.deleteRows(2, count);
    const resetAt = new Date();
    _audit('reset', 'application', '', {deletedTransactions:count});
    return _ok('Data transaksi berhasil dihapus.', {resetAt:_dateTime(resetAt),deletedTransactions:count});
  } catch (err) {
    return _fail(err);
  }
}

function _dashboard(month, year) {
  const m = _int(month, new Date().getMonth()+1);
  const y = _int(year, new Date().getFullYear());
  const all = _read('Transactions').map(_transactionObject);
  const categories = _categories(null, false);
  const categoryMap = {};
  categories.forEach(function(c) { categoryMap[c.id] = c; });
  const period = all.filter(function(t) {
    const d = _dateParts(t.transactionDate);
    return d && d.month === m && d.year === y;
  });
  const income = _sum(period, 'income');
  const expense = _sum(period, 'expense');
  const balance = all.reduce(function(total,t) { return total + (t.type === 'income' ? t.amount : -t.amount); }, 0);
  const previousDate = new Date(y, m-2, 1);
  const pm = previousDate.getMonth()+1, py = previousDate.getFullYear();
  const previous = all.filter(function(t) {
    const d = _dateParts(t.transactionDate);
    return d && d.month === pm && d.year === py;
  });
  const previousBalance = all.filter(function(t) {
    const d = _dateParts(t.transactionDate);
    return d && (d.year < py || (d.year === py && d.month <= pm));
  }).reduce(function(total,t) { return total + (t.type === 'income' ? t.amount : -t.amount); }, 0);
  const categoryTotals = {};
  period.filter(function(t) { return t.type === 'expense'; }).forEach(function(t) {
    const id = t.categoryId || '';
    if (!categoryTotals[id]) categoryTotals[id] = 0;
    categoryTotals[id] += t.amount;
  });
  const categoryBreakdown = Object.keys(categoryTotals).map(function(id) {
    const category = categoryMap[id];
    return {categoryId:id,name:category ? category.name : 'Tanpa kategori',amount:categoryTotals[id]};
  }).sort(function(a,b) { return b.amount-a.amount; });
  const trend = [];
  for (let i=5; i>=0; i--) {
    const d = new Date(y,m-1-i,1);
    const tm=d.getMonth()+1, ty=d.getFullYear();
    const records=all.filter(function(t) {
      const parts=_dateParts(t.transactionDate);
      return parts && parts.month===tm && parts.year===ty;
    });
    trend.push({
      label:['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'][tm-1],
      month:tm,year:ty,income:_sum(records,'income'),expense:_sum(records,'expense')
    });
  }
  const days = new Date(y,m,0).getDate();
  return {
    balance:balance,
    income:income,
    expense:expense,
    savingsRate:income ? Math.round(((income-expense)/income)*100) : 0,
    transactionCount:period.length,
    averageDailyExpense:Math.round(expense/days),
    previousPeriod:{balance:previousBalance,income:_sum(previous,'income'),expense:_sum(previous,'expense')},
    categoryBreakdown:categoryBreakdown,
    cashFlowTrend:trend
  };
}

function _filteredTransactions(filters) {
  const f=filters||{};
  let items=_read('Transactions').map(_transactionObject);
  if (f.date) items=items.filter(function(t) { return t.transactionDate===String(f.date); });
  else if (f.month && f.year) items=items.filter(function(t) { const d=_dateParts(t.transactionDate); return d && d.month===_int(f.month,0) && d.year===_int(f.year,0); });
  if (f.type) items=items.filter(function(t) { return t.type===String(f.type); });
  if (f.categoryId) items=items.filter(function(t) { return t.categoryId===String(f.categoryId); });
  const query=String(f.search||f.query||'').trim().toLowerCase();
  const categoryMap={};
  _categories(null,false).forEach(function(c){categoryMap[c.id]=c.name});
  const accountMap={};
  _accounts(false).forEach(function(a){accountMap[a.id]=a.name});
  items=items.map(function(t){
    t.categoryName=categoryMap[t.categoryId]||'Tanpa kategori';
    t.accountName=accountMap[t.source]||'';
    return t;
  });
  if(query) items=items.filter(function(t){return [t.note,t.categoryName,t.type,t.amount].join(' ').toLowerCase().indexOf(query)!==-1;});
  const sort=String(f.sortBy||'transactionDate');
  const dir=String(f.sortDirection||'desc').toLowerCase()==='asc'?1:-1;
  items.sort(function(a,b){
    let av=a[sort],bv=b[sort];
    if(sort==='amount'){av=Number(av);bv=Number(bv);return (av-bv)*dir}
    if(sort==='createdAt'){av=new Date(av).getTime();bv=new Date(bv).getTime();return (av-bv)*dir}
    return String(av||'').localeCompare(String(bv||''))*dir;
  });
  const filteredTotal=items.length;
  const offset=Math.max(0,_int(f.offset,0));
  const limit=Math.min(50000,Math.max(1,_int(f.limit,500)));
  return {items:items.slice(offset,offset+limit),total:_read('Transactions').length,filteredTotal:filteredTotal};
}

function _validateTransaction(payload) {
  const p=payload||{};
  const type=String(p.type||'');
  if(type!=='income'&&type!=='expense')throw new Error('Jenis transaksi tidak valid.');
  const amount=Number(p.amount);
  if(!isFinite(amount)||amount<=0)throw new Error('Nominal transaksi harus lebih dari nol.');
  const categoryId=String(p.categoryId||p.category_id||'');
  const category=_findById('Categories',categoryId);
  if(!category)throw new Error('Kategori tidak ditemukan.');
  if(String(category.type)!==type)throw new Error('Jenis kategori tidak sesuai dengan transaksi.');
  if(!_bool(category.is_active))throw new Error('Kategori ini sedang tidak aktif.');
  const transactionDate=String(p.transactionDate||p.transaction_date||'');
  if(!_dateParts(transactionDate))throw new Error('Tanggal transaksi tidak valid.');
  const note=String(p.note||'').trim();
  if(note.length>240)throw new Error('Catatan maksimal 240 karakter.');
  let source=String(p.source||'manual');
  const accountId=String(p.accountId||(type==='income'?p.source:'')||'');
  if(type==='income'&&!accountId)throw new Error('Pilih rekening sumber pemasukan yang aktif.');
  if(accountId){
    const account=_findById('Accounts',accountId);
    if(!account||!_bool(account.is_active))throw new Error('Pilih rekening sumber transaksi yang aktif.');
    source=accountId;
  }
  return {type:type,amount:amount,categoryId:categoryId,transactionDate:transactionDate,note:note,source:source};
}

function _settings() {
  const rows=_read('Users');
  const r=rows[0]||{};
  return {
    id:String(r.id||''),
    name:String(r.name||'Saya'),
    currency:String(r.currency||'IDR'),
    timezone:String(r.timezone||'Asia/Jakarta'),
    fiscalYearStart:_int(r.fiscal_year_start,1),
    createdAt:_dateTime(r.created_at)
  };
}

function _categories(type, activeOnly) {
  return _read('Categories').filter(function(r) {
    return (!type||String(r.type)===String(type))&&(!activeOnly||_bool(r.is_active));
  }).map(_categoryObject).sort(function(a,b){return Number(a.sortOrder)-Number(b.sortOrder)||a.name.localeCompare(b.name)});
}

function _accounts(activeOnly) {
  return _read('Accounts').filter(function(r) {
    return !activeOnly || _bool(r.is_active);
  }).map(_accountObject).sort(function(a,b) {
    return a.name.localeCompare(b.name);
  });
}

function _accountObject(r) {
  return {
    id:String(r.id||''),
    name:String(r.name||''),
    accountNumber:String(r.account_number||''),
    isActive:_bool(r.is_active),
    createdAt:_dateTime(r.created_at),
    updatedAt:_dateTime(r.updated_at)
  };
}

function _categoryObject(r) {
  return {
    id:String(r.id||''),
    name:String(r.name||''),
    type:String(r.type||'expense'),
    icon:String(r.icon||'sell'),
    color:String(r.color||'#238467'),
    isDefault:_bool(r.is_default),
    isActive:_bool(r.is_active),
    sortOrder:_int(r.sort_order,0),
    createdAt:_dateTime(r.created_at),
    updatedAt:_dateTime(r.updated_at)
  };
}

function _transactionObject(r) {
  return {
    id:String(r.id||''),
    createdAt:_dateTime(r.created_at),
    updatedAt:_dateTime(r.updated_at),
    type:String(r.type||'expense'),
    amount:Number(r.amount)||0,
    categoryId:String(r.category_id||''),
    transactionDate:_dateOnly(r.transaction_date),
    note:String(r.note||''),
    source:String(r.source||'manual')
  };
}

function _sheet(name) {
  const ss=_databaseSpreadsheet();
  return ss.getSheetByName(name)||ss.insertSheet(name);
}

function _read(name) {
  const sheet=_sheet(name);
  const values=sheet.getDataRange().getValues();
  if(values.length<2)return [];
  const headers=values[0].map(String);
  return values.slice(1).filter(function(row){return row.some(function(v){return v!==''&&v!==null})}).map(function(row){return _rowObject(headers,row)});
}

function _rowObject(headers,row) {
  const result={};
  headers.forEach(function(h,i){result[h]=row[i]});
  return result;
}

function _findById(name,id) {
  return _read(name).filter(function(r){return String(r.id)===String(id)})[0]||null;
}

function _rowNumberById(sheet,id) {
  const last=sheet.getLastRow();
  if(last<2)return 0;
  const ids=sheet.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<ids.length;i++)if(String(ids[i][0])===String(id))return i+2;
  return 0;
}

function _sum(items,type) {
  return items.filter(function(t){return t.type===type}).reduce(function(total,t){return total+(Number(t.amount)||0)},0);
}

function _audit(action,entityType,entityId,details) {
  _sheet('AuditLogs').appendRow([_uuid(),new Date(),action,entityType,String(entityId||''),JSON.stringify(details||{})]);
}

function _dateParts(value) {
  if(value instanceof Date&&!isNaN(value.getTime()))return {year:value.getFullYear(),month:value.getMonth()+1,day:value.getDate()};
  const s=String(value||'').trim();
  const m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(m)return {year:Number(m[1]),month:Number(m[2]),day:Number(m[3])};
  const d=new Date(s);
  if(isNaN(d.getTime()))return null;
  return {year:d.getFullYear(),month:d.getMonth()+1,day:d.getDate()};
}

function _monthFromDate(value){const p=_dateParts(value);return p?p.month:new Date().getMonth()+1}
function _yearFromDate(value){const p=_dateParts(value);return p?p.year:new Date().getFullYear()}

function _dateOnly(value) {
  if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const p=_dateParts(value);
  return p?String(p.year)+'-'+('0'+p.month).slice(-2)+'-'+('0'+p.day).slice(-2):String(value||'');
}

function _dateTime(value) {
  if(value instanceof Date&&!isNaN(value.getTime()))return Utilities.formatDate(value,Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm:ss");
  return value?String(value):'';
}

function _uuid() {
  return Utilities.getUuid();
}

function _int(value,fallback) {
  const n=parseInt(value,10);
  return isNaN(n)?fallback:n;
}

function _bool(value) {
  return value===true||String(value).toLowerCase()==='true'||String(value)==='1';
}

function _safeIcon(value) {
  const s=String(value||'sell');
  return /^[a-z0-9_]{2,40}$/.test(s)?s:'sell';
}

function _safeColor(value) {
  const s=String(value||'#238467');
  return /^#[0-9a-fA-F]{6}$/.test(s)?s:'#238467';
}

function _csvCell(value) {
  const s=String(value===null||value===undefined?'':value);
  return '"'+s.replace(/"/g,'""')+'"';
}

function _parseJson(value) {
  try{return JSON.parse(String(value||'{}'))}catch(err){return {}}
}

function _ok(message,data) {
  return {success:true,message:message,data:data};
}

function _fail(err) {
  return {success:false,message:err&&err.message?err.message:String(err),data:null};
}

var ACTIVE_USER_SPREADSHEET_ID = '';

function _databaseSpreadsheet() {
  return ACTIVE_USER_SPREADSHEET_ID
    ? SpreadsheetApp.openById(ACTIVE_USER_SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function _useSession(token) {
  const key = 'RUANGUANG_SESSION_' + String(token || '');
  const stored = PropertiesService.getScriptProperties().getProperty(key);
  if (!stored) throw new Error('Sesi tidak valid. Silakan masuk kembali.');
  const parts = stored.split('|');
  if (parts.length !== 2 || Date.now() - Number(parts[1]) > 30 * 24 * 60 * 60 * 1000) {
    PropertiesService.getScriptProperties().deleteProperty(key);
    throw new Error('Sesi berakhir. Silakan masuk kembali.');
  }
  ACTIVE_USER_SPREADSHEET_ID = parts[0];
}

function _authUserKey(username) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(username).toLowerCase(), Utilities.Charset.UTF_8);
  return 'RUANGUANG_USER_' + Utilities.base64Encode(bytes).replace(/[^a-zA-Z0-9]/g, '');
}

function _passwordHash(salt, password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + '|' + String(password), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(bytes);
}

function _createSession(spreadsheetId) {
  const token = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty('RUANGUANG_SESSION_' + token, String(spreadsheetId) + '|' + Date.now());
  return token;
}

function registerUser(username, password) {
  try {
    const name = String(username || '').trim();
    const secret = String(password || '');
    if (!/^[a-zA-Z0-9._-]{3,30}$/.test(name)) throw new Error('Nama pengguna harus berisi 3–30 karakter: huruf, angka, titik, garis bawah, atau tanda hubung.');
    if (secret.length < 8 || secret.length > 128) throw new Error('Kata sandi harus terdiri dari 8–128 karakter.');
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const properties = PropertiesService.getScriptProperties();
      const userKey = _authUserKey(name);
      if (properties.getProperty(userKey)) throw new Error('Nama pengguna tersebut sudah terdaftar.');
      const spreadsheet = SpreadsheetApp.create('RuangUang - ' + name);
      ACTIVE_USER_SPREADSHEET_ID = spreadsheet.getId();
      const setup = setupSpreadsheet();
      if (!setup.success) throw new Error(setup.message || 'Akun belum dapat disiapkan.');
      const users = _sheet('Users');
      if (users.getLastRow() >= 2) users.getRange(2, 3).setValue(name);
      const salt = Utilities.getUuid();
      properties.setProperty(userKey, JSON.stringify({
        username:name,
        salt:salt,
        passwordHash:_passwordHash(salt, secret),
        spreadsheetId:spreadsheet.getId()
      }));
      return _ok('Akun berhasil dibuat.', {token:_createSession(spreadsheet.getId()), username:name});
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return _fail(err);
  }
}

function loginUser(username, password) {
  try {
    const name = String(username || '').trim();
    const record = PropertiesService.getScriptProperties().getProperty(_authUserKey(name));
    if (!record) throw new Error('Nama pengguna atau kata sandi tidak sesuai.');
    const user = JSON.parse(record);
    if (_passwordHash(user.salt, password) !== user.passwordHash) throw new Error('Nama pengguna atau kata sandi tidak sesuai.');
    return _ok('Berhasil masuk.', {token:_createSession(user.spreadsheetId), username:user.username});
  } catch (err) {
    return _fail(err);
  }
}
