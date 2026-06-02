// ═══════════════════════════════════════════════════════════════
// PharmaTrack Pro — Code.gs (Backend)
// Built by SME Systems Studio · Patience Oduori
// oduoripatience@gmail.com
// ═══════════════════════════════════════════════════════════════
//
// HOW TO INSTALL:
// 1. Open your PharmaTrack Pro DB Google Sheet
// 2. Extensions → Apps Script
// 3. Delete everything → paste this file → save as Code.gs
// 4. Create a new file → name it index.html → paste the HTML file
// 5. Deploy → New Deployment → Web App
// 6. Execute as: Me | Who has access: Anyone
// 7. Copy the web app URL → open in browser
// ═══════════════════════════════════════════════════════════════

// ── SHEET NAMES ─────────────────────────────────────────────────
var SS = {
  USERS:      'Users',
  INVENTORY:  'Inventory',
  SALES:      'Sales',
  SALE_ITEMS: 'SaleItems',
  CUSTOMERS:  'Customers',
  EXPENSES:   'Expenses',
  SETTINGS:   'Settings'
};

// ── SERVE THE WEB APP ─────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('PharmaTrack Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── HELPER: GET SHEET ─────────────────────────────────────────────
function sheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// ── HELPER: SHEET DATA AS OBJECTS ────────────────────────────────
function sheetToObjects(name) {
  var s = sheet(name);
  if (!s) return [];
  var rows = s.getDataRange().getValues();
  if (rows.length < 2) return [];
  var headers = rows[0];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = rows[i][j];
    }
    result.push(obj);
  }
  return result;
}

// ── HELPER: UNIQUE ID ─────────────────────────────────────────────
function uid(prefix) {
  return prefix + '-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
}

// ── HELPER: FORMAT DATE ───────────────────────────────────────────
function fmtDate(d) {
  if (!d) return '';
  try { return Utilities.formatDate(new Date(d), 'UTC', 'yyyy-MM-dd'); }
  catch(e) { return String(d).substring(0, 10); }
}

// ════════════════════════════════════════════════════════════════
// AUTHENTICATION
// ════════════════════════════════════════════════════════════════
function loginUser(username, password) {
  try {
    var s = sheet(SS.USERS);
    var data = s.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === username &&
          data[i][2] === password &&
          data[i][6] === true) {
        return {
          success: true,
          user: {
            id:       data[i][0],
            username: data[i][1],
            role:     data[i][3],
            name:     data[i][4],
            email:    data[i][5]
          }
        };
      }
    }
    return { success: false, error: 'Invalid username or password' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getUsers() {
  try {
    var users = sheetToObjects(SS.USERS).map(function(u) {
      return { id: u['ID'], username: u['Username'], role: u['Role'],
               name: u['Name'], email: u['Email'], active: u['Active'] };
    });
    return { success: true, data: users };
  } catch(e) { return { success: false, error: e.message }; }
}

function addUser(data) {
  try {
    var s = sheet(SS.USERS);
    s.appendRow([uid('USR'), data.username, data.password,
                 data.role, data.name, data.email || '', true]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// INVENTORY
// ════════════════════════════════════════════════════════════════
function getInventory() {
  try {
    var s = sheet(SS.INVENTORY);
    var data = s.getDataRange().getValues();
    var today = new Date();
    var medicines = [];

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][10] === false) continue;

      var stock    = Number(data[i][5]) || 0;
      var minStock = Number(data[i][6]) || 0;
      var expiry   = data[i][4] ? new Date(data[i][4]) : null;
      var days     = expiry ? Math.floor((expiry - today) / 86400000) : null;

      var status = 'In Stock';
      if (stock === 0)              status = 'Out of Stock';
      else if (stock <= minStock)   status = 'Low Stock';
      else if (days !== null && days <= 60) status = 'Expiring Soon';

      medicines.push({
        id:           data[i][0],
        name:         data[i][1],
        category:     data[i][2],
        batch:        data[i][3],
        expiry:       fmtDate(data[i][4]),
        stock:        stock,
        minStock:     minStock,
        costPrice:    Number(data[i][7]) || 0,
        sellingPrice: Number(data[i][8]) || 0,
        description:  data[i][9] || '',
        status:       status,
        daysToExpiry: days
      });
    }
    return { success: true, data: medicines };
  } catch(e) { return { success: false, error: e.message }; }
}

function addMedicine(data) {
  try {
    var s = sheet(SS.INVENTORY);
    var id = uid('MED');
    s.appendRow([id, data.name, data.category, data.batch || '',
                 data.expiry || '', Number(data.stock) || 0,
                 Number(data.minStock) || 10,
                 Number(data.costPrice) || 0,
                 Number(data.sellingPrice) || 0,
                 data.description || '', true]);
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

function updateMedicine(data) {
  try {
    var s    = sheet(SS.INVENTORY);
    var rows = s.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        s.getRange(i + 1, 2, 1, 9).setValues([[
          data.name, data.category, data.batch || '',
          data.expiry || '', Number(data.stock) || 0,
          Number(data.minStock) || 10,
          Number(data.costPrice) || 0,
          Number(data.sellingPrice) || 0,
          data.description || ''
        ]]);
        return { success: true };
      }
    }
    return { success: false, error: 'Medicine not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

function deleteMedicine(id) {
  try {
    var s    = sheet(SS.INVENTORY);
    var rows = s.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        s.getRange(i + 1, 11).setValue(false);
        return { success: true };
      }
    }
    return { success: false, error: 'Not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// SALES
// ════════════════════════════════════════════════════════════════
function createSale(saleData) {
  try {
    var salesSheet = sheet(SS.SALES);
    var itemsSheet = sheet(SS.SALE_ITEMS);
    var invSheet   = sheet(SS.INVENTORY);

    var saleId    = uid('SALE');
    var receiptNo = 'RCP-' + String(Date.now()).slice(-6);
    var now       = new Date();
    var dateStr   = fmtDate(now);
    var timeStr   = Utilities.formatDate(now, 'UTC', 'HH:mm:ss');

    // Save sale header row
    salesSheet.appendRow([
      saleId, dateStr, timeStr,
      saleData.customerName  || 'Walk-in',
      saleData.customerPhone || '',
      Number(saleData.total),
      saleData.paymentMethod,
      Number(saleData.amountPaid),
      Number(saleData.change)  || 0,
      saleData.cashier         || 'Cashier',
      receiptNo,
      saleData.notes           || ''
    ]);

    // Save line items + deduct inventory
    var invData = invSheet.getDataRange().getValues();
    for (var i = 0; i < saleData.items.length; i++) {
      var item = saleData.items[i];
      itemsSheet.appendRow([
        saleId, item.name,
        Number(item.qty), Number(item.price),
        Number(item.qty) * Number(item.price),
        item.category || ''
      ]);

      // Deduct stock
      for (var j = 1; j < invData.length; j++) {
        if (invData[j][1] === item.name) {
          var newStock = Math.max(0, Number(invData[j][5]) - Number(item.qty));
          invSheet.getRange(j + 1, 6).setValue(newStock);
          invData[j][5] = newStock;
          break;
        }
      }
    }

    // Update customer record
    if (saleData.customerName && saleData.customerName !== 'Walk-in') {
      _updateCustomerSpend(saleData.customerName,
                           saleData.customerPhone,
                           Number(saleData.total));
    }

    return { success: true, saleId: saleId, receiptNo: receiptNo };
  } catch(e) { return { success: false, error: e.message }; }
}

function getSales(limit) {
  try {
    var s    = sheet(SS.SALES);
    var data = s.getDataRange().getValues();
    var max  = limit || 100;
    var sales = [];

    for (var i = data.length - 1; i >= 1 && sales.length < max; i--) {
      if (!data[i][0]) continue;
      sales.push({
        id: data[i][0], date: fmtDate(data[i][1]), time: data[i][2],
        customer: data[i][3], phone: data[i][4],
        total: Number(data[i][5]), payment: data[i][6],
        amountPaid: Number(data[i][7]), change: Number(data[i][8]),
        cashier: data[i][9], receiptNo: data[i][10], notes: data[i][11]
      });
    }
    return { success: true, data: sales };
  } catch(e) { return { success: false, error: e.message }; }
}

function getSaleItems(saleId) {
  try {
    var s    = sheet(SS.SALE_ITEMS);
    var data = s.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === saleId) {
        items.push({
          saleId: data[i][0], name: data[i][1],
          qty: Number(data[i][2]), price: Number(data[i][3]),
          total: Number(data[i][4]), category: data[i][5]
        });
      }
    }
    return { success: true, data: items };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// CUSTOMERS
// ════════════════════════════════════════════════════════════════
function getCustomers() {
  try {
    var s    = sheet(SS.CUSTOMERS);
    var data = s.getDataRange().getValues();
    var customers = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      customers.push({
        id: data[i][0], name: data[i][1], phone: data[i][2],
        email: data[i][3], address: data[i][4],
        totalSpent: Number(data[i][5]) || 0,
        visits: Number(data[i][6]) || 0,
        dateAdded: fmtDate(data[i][7]),
        notes: data[i][8] || ''
      });
    }
    return { success: true, data: customers };
  } catch(e) { return { success: false, error: e.message }; }
}

function addCustomer(data) {
  try {
    var s  = sheet(SS.CUSTOMERS);
    var id = uid('CUST');
    s.appendRow([id, data.name, data.phone, data.email || '',
                 data.address || '', 0, 0, fmtDate(new Date()), data.notes || '']);
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

function updateCustomer(data) {
  try {
    var s    = sheet(SS.CUSTOMERS);
    var rows = s.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.id) {
        s.getRange(i + 1, 2, 1, 4).setValues([[
          data.name, data.phone, data.email || '', data.address || ''
        ]]);
        if (data.notes !== undefined) s.getRange(i + 1, 9).setValue(data.notes);
        return { success: true };
      }
    }
    return { success: false, error: 'Customer not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

function _updateCustomerSpend(name, phone, amount) {
  var s    = sheet(SS.CUSTOMERS);
  var data = s.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === name || (phone && data[i][2] === phone)) {
      s.getRange(i + 1, 6).setValue(Number(data[i][5]) + amount);
      s.getRange(i + 1, 7).setValue(Number(data[i][6]) + 1);
      return;
    }
  }
  // Auto-register new customer
  if (name && name !== 'Walk-in') {
    sheet(SS.CUSTOMERS).appendRow([
      uid('CUST'), name, phone || '', '', '', amount, 1, fmtDate(new Date()), ''
    ]);
  }
}

// ════════════════════════════════════════════════════════════════
// EXPENSES
// ════════════════════════════════════════════════════════════════
function getExpenses() {
  try {
    var s    = sheet(SS.EXPENSES);
    var data = s.getDataRange().getValues();
    var expenses = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      expenses.push({
        id: data[i][0], date: fmtDate(data[i][1]),
        category: data[i][2], description: data[i][3],
        amount: Number(data[i][4]), recordedBy: data[i][5], notes: data[i][6]
      });
    }
    return { success: true, data: expenses };
  } catch(e) { return { success: false, error: e.message }; }
}

function addExpense(data) {
  try {
    var s  = sheet(SS.EXPENSES);
    var id = uid('EXP');
    s.appendRow([id, fmtDate(new Date()), data.category,
                 data.description, Number(data.amount),
                 data.recordedBy || '', data.notes || '']);
    return { success: true, id: id };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// DASHBOARD
// ════════════════════════════════════════════════════════════════
function getDashboardData() {
  try {
    var salesSheet = sheet(SS.SALES);
    var salesData  = salesSheet.getDataRange().getValues();

    var today     = fmtDate(new Date());
    var thisMonth = today.substring(0, 7);

    var todayRev = 0, todayTxns = 0;
    var monthRev = 0, monthTxns = 0;
    var totalRev = 0, totalTxns = 0;

    // Revenue per category
    var itemsSheet = sheet(SS.SALE_ITEMS);
    var itemsData  = itemsSheet.getDataRange().getValues();
    var catRevenue = {};

    for (var i = 1; i < itemsData.length; i++) {
      var cat = itemsData[i][5] || 'Other';
      catRevenue[cat] = (catRevenue[cat] || 0) + Number(itemsData[i][4]);
    }

    // Sales totals
    var trend = {};
    for (var k = 6; k >= 0; k--) {
      var d = new Date();
      d.setDate(d.getDate() - k);
      trend[fmtDate(d)] = 0;
    }

    var recentSales = [];
    for (var j = salesData.length - 1; j >= 1; j--) {
      if (!salesData[j][0]) continue;
      var sDate  = fmtDate(salesData[j][1]);
      var amount = Number(salesData[j][5]) || 0;

      totalRev += amount; totalTxns++;
      if (sDate === today)              { todayRev += amount; todayTxns++; }
      if (sDate.startsWith(thisMonth))  { monthRev += amount; monthTxns++; }
      if (trend.hasOwnProperty(sDate))  { trend[sDate] += amount; }
      if (recentSales.length < 8) {
        recentSales.push({
          date: sDate, customer: salesData[j][3],
          total: amount, payment: salesData[j][6],
          cashier: salesData[j][9], receiptNo: salesData[j][10]
        });
      }
    }

    // Inventory stats
    var inv        = getInventory().data || [];
    var lowStock   = inv.filter(function(m){ return m.status==='Low Stock'; }).length;
    var expiring   = inv.filter(function(m){ return m.status==='Expiring Soon'; }).length;
    var outOfStock = inv.filter(function(m){ return m.status==='Out of Stock'; }).length;

    return {
      success: true,
      data: {
        todayRevenue:  todayRev,  todayTxns:  todayTxns,
        monthRevenue:  monthRev,  monthTxns:  monthTxns,
        totalRevenue:  totalRev,  totalTxns:  totalTxns,
        lowStock:      lowStock,  expiringSoon: expiring,
        outOfStock:    outOfStock, totalMedicines: inv.length,
        recentSales:   recentSales,
        trend:         trend,
        catRevenue:    catRevenue
      }
    };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// REPORTS
// ════════════════════════════════════════════════════════════════
function getSalesReport(startDate, endDate) {
  try {
    var s    = sheet(SS.SALES);
    var data = s.getDataRange().getValues();
    var sales = [];
    var total = 0;

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var d = fmtDate(data[i][1]);
      if (d >= startDate && d <= endDate) {
        var amt = Number(data[i][5]) || 0;
        total += amt;
        sales.push({
          date: d, customer: data[i][3], receiptNo: data[i][10],
          total: amt, payment: data[i][6], cashier: data[i][9]
        });
      }
    }
    return { success: true, data: sales, total: total };
  } catch(e) { return { success: false, error: e.message }; }
}

function getProfitLoss(startDate, endDate) {
  try {
    // Revenue
    var salesReport = getSalesReport(startDate, endDate);
    var revenue = salesReport.total || 0;

    // Cost of sales from sale items x cost price
    var itemsSheet = sheet(SS.SALE_ITEMS);
    var invSheet   = sheet(SS.INVENTORY);
    var itemsData  = itemsSheet.getDataRange().getValues();
    var invData    = invSheet.getDataRange().getValues();

    // Build cost lookup
    var costMap = {};
    for (var j = 1; j < invData.length; j++) {
      costMap[invData[j][1]] = Number(invData[j][7]) || 0;
    }

    // Get sale IDs within date range
    var salesIds = {};
    var salesSheet = sheet(SS.SALES);
    var salesData  = salesSheet.getDataRange().getValues();
    for (var k = 1; k < salesData.length; k++) {
      var d = fmtDate(salesData[k][1]);
      if (d >= startDate && d <= endDate && salesData[k][0]) {
        salesIds[salesData[k][0]] = true;
      }
    }

    var cogs = 0;
    for (var i = 1; i < itemsData.length; i++) {
      if (salesIds[itemsData[i][0]]) {
        cogs += (costMap[itemsData[i][1]] || 0) * (Number(itemsData[i][2]) || 0);
      }
    }

    // Expenses in range
    var expSheet = sheet(SS.EXPENSES);
    var expData  = expSheet.getDataRange().getValues();
    var expenses = 0;
    for (var m = 1; m < expData.length; m++) {
      var ed = fmtDate(expData[m][1]);
      if (ed >= startDate && ed <= endDate) {
        expenses += Number(expData[m][4]) || 0;
      }
    }

    var grossProfit = revenue - cogs;
    var netProfit   = grossProfit - expenses;

    return {
      success: true,
      data: {
        revenue: revenue, cogs: cogs,
        grossProfit: grossProfit, expenses: expenses,
        netProfit: netProfit,
        grossMargin: revenue > 0 ? ((grossProfit / revenue) * 100).toFixed(1) : 0,
        netMargin:   revenue > 0 ? ((netProfit / revenue)   * 100).toFixed(1) : 0
      }
    };
  } catch(e) { return { success: false, error: e.message }; }
}

function getExpiryReport() {
  try {
    var inv = getInventory().data || [];
    var today = new Date();
    var expiring = inv.filter(function(m) {
      return m.daysToExpiry !== null && m.daysToExpiry <= 90;
    }).sort(function(a, b) {
      return (a.daysToExpiry || 999) - (b.daysToExpiry || 999);
    });
    return { success: true, data: expiring };
  } catch(e) { return { success: false, error: e.message }; }
}

function getLowStockReport() {
  try {
    var inv = getInventory().data || [];
    var low = inv.filter(function(m) {
      return m.status === 'Low Stock' || m.status === 'Out of Stock';
    }).sort(function(a, b) { return a.stock - b.stock; });
    return { success: true, data: low };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// SETTINGS
// ════════════════════════════════════════════════════════════════
function getSettings() {
  try {
    var s    = sheet(SS.SETTINGS);
    var data = s.getDataRange().getValues();
    var cfg  = {
      pharmacyName:      'PharmaTrack Pro',
      currency:          'KES',
      address:           'Nairobi, Kenya',
      phone:             '',
      email:             '',
      expiryWarningDays: 60,
      lowStockAlert:     10,
      taxRate:           0,
      receiptFooter:     'Thank you for your business!'
    };
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) cfg[data[i][0]] = data[i][1];
    }
    return { success: true, data: cfg };
  } catch(e) { return { success: false, error: e.message }; }
}

function saveSettings(data) {
  try {
    var s = sheet(SS.SETTINGS);
    s.clearContents();
    Object.keys(data).forEach(function(k) {
      s.appendRow([k, data[k]]);
    });
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ════════════════════════════════════════════════════════════════
// INITIALIZE — Run once to seed data
// ════════════════════════════════════════════════════════════════
function initializeSystem() {
  try {
    // Add sample medicines to inventory if empty
    var invSheet = sheet(SS.INVENTORY);
    if (invSheet.getLastRow() <= 1) {
      var medicines = [
        ['MED-001','Amoxicillin 500mg','Antibiotics','BCH-001','2025-12-31',150,20,45,120,'',true],
        ['MED-002','Paracetamol 500mg','Analgesics','BCH-002','2026-06-30',400,50,8,25,'',true],
        ['MED-003','Metformin 500mg','Diabetes','BCH-003','2025-09-15',80,15,32,85,'',true],
        ['MED-004','Atorvastatin 10mg','Cardiac','BCH-004','2026-03-20',60,10,85,220,'',true],
        ['MED-005','Omeprazole 20mg','Gastrointestinal','BCH-005','2025-08-10',45,10,38,95,'',true],
        ['MED-006','Ciprofloxacin 250mg','Antibiotics','BCH-006','2026-01-15',120,20,55,140,'',true],
        ['MED-007','Salbutamol Inhaler','Respiratory','BCH-007','2025-07-20',30,8,380,950,'',true],
        ['MED-008','Ibuprofen 400mg','Analgesics','BCH-008','2026-04-30',200,30,12,35,'',true],
        ['MED-009','Insulin Glargine','Diabetes','BCH-009','2025-06-30',25,5,1200,2800,'',true],
        ['MED-010','Vitamin C 500mg','Supplements','BCH-010','2026-12-31',300,50,5,18,'',true],
      ];
      medicines.forEach(function(m) { invSheet.appendRow(m); });
    }

    // Default settings
    var settingsSheet = sheet(SS.SETTINGS);
    if (settingsSheet.getLastRow() <= 1) {
      [['pharmacyName','PharmaTrack Demo Pharmacy'],
       ['currency','KES'],['address','Nairobi, Kenya'],
       ['phone','+254 700 000 000'],['email','owner@pharmatrack.com'],
       ['expiryWarningDays','60'],['lowStockAlert','10'],
       ['taxRate','0'],['receiptFooter','Thank you for your business!']
      ].forEach(function(r) { settingsSheet.appendRow(r); });
    }

    return { success: true, message: 'System initialized with sample data!' };
  } catch(e) { return { success: false, error: e.message }; }
}

function updateExpense(data) {
  try {
    var s    = sheet(SS.EXPENSES);
    var rows = s.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        s.getRange(i+1, 3, 1, 4).setValues([[
          data.category, data.description,
          Number(data.amount), data.notes || ''
        ]]);
        return { success: true };
      }
    }
    return { success: false, error: 'Expense not found' };
  } catch(e) { return { success: false, error: e.message }; }
}