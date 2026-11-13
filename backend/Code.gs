
/***************************************************************
 * IMARA MILK DELIVERY HUB
 * Google Apps Script + Google Sheets backend
 *
 * Spreadsheet: Imara Milk Delivery Hub
 * Frontend: GitHub Pages
 *
 * Main features:
 * - Customer registration and profile retrieval
 * - Primary customer location registration
 * - Customer locations for Google Maps
 * - Delivery order creation and tracking
 * - Dashboard statistics
 * - Order status history
 * - Customer messaging and conversations
 *
 * IMPORTANT:
 * 1. Keep your existing spreadsheet ID.
 * 2. Run setupImaraDatabase() once.
 * 3. Deploy as a Web App after saving.
 * 4. GitHub Pages uses the JSONP doGet endpoint; no CORS POST is required.
 ***************************************************************/

const IMARA_SPREADSHEET_ID =
  '1apujiw_j742V6hrmzvG3GRd67P2yM2GMjNbFVW4jwpI';

const IMARA_SHEETS = {
  Customers: [
    'CustomerID',
    'FullName',
    'Phone',
    'Email',
    'CreatedAt',
    'UpdatedAt',
    'AccountStatus'
  ],

  UserLocations: [
    'LocationID',
    'CustomerID',
    'Label',
    'Address',
    'Latitude',
    'Longitude',
    'IsPrimary',
    'CreatedAt',
    'UpdatedAt'
  ],

  Orders: [
    'OrderID',
    'CustomerName',
    'Phone',
    'Address',
    'Latitude',
    'Longitude',
    'Quantity',
    'Status',
    'CreatedAt',
    'CustomerID',
    'LocationID',
    'Price',
    'Notes',
    'PaymentStatus',
    'PaymentReference',
    'AssignedStaff',
    'UpdatedAt'
  ],

  OrderStatusHistory: [
    'HistoryID',
    'OrderID',
    'PreviousStatus',
    'NewStatus',
    'ChangedBy',
    'Notes',
    'CreatedAt'
  ],

  Messages: [
    'MessageID',
    'CustomerID',
    'Sender',
    'SenderName',
    'Message',
    'CreatedAt',
    'ReadStatus'
  ],

  Staff: [
    'StaffID',
    'FullName',
    'Phone',
    'Email',
    'Role',
    'AccountStatus',
    'CreatedAt'
  ]
};


/* =============================================================
   DATABASE SETUP
============================================================= */

function getDatabase_() {
  return SpreadsheetApp.openById(IMARA_SPREADSHEET_ID);
}

function getSheet_(name) {
  const ss = getDatabase_();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function setupImaraDatabase() {
  return withScriptLock_(function () {
  Object.keys(IMARA_SHEETS).forEach(name => {
    const sheet = getSheet_(name);
    const headers = IMARA_SHEETS[name];

    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      return;
    }

    const existing = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(String);

    headers.forEach(header => {
      if (existing.indexOf(header) === -1) {
        const nextColumn = sheet.getLastColumn() + 1;
        sheet.getRange(1, nextColumn).setValue(header);
        existing.push(header);
      }
    });

    sheet.setFrozenRows(1);
  });

  return {
    success: true,
    message: 'Imara database setup completed.'
  };

  });
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];

  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(String);
}

function getRecords_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const headers = getHeaders_(sheet);
  const values = sheet
    .getRange(2, 1, lastRow - 1, headers.length)
    .getValues();

  return values.map(row => {
    const record = {};

    headers.forEach((header, index) => {
      let value = row[index];

      if (value instanceof Date) {
        value = value.toISOString();
      }

      record[header] = value;
    });

    return record;
  }).filter(record =>
    Object.values(record).some(value =>
      value !== '' && value !== null && value !== undefined
    )
  );
}

function appendRecord_(sheetName, record) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);

  const row = headers.map(header =>
    record[header] !== undefined ? record[header] : ''
  );

  sheet.appendRow(row);
  return record;
}

function updateRecord_(sheetName, key, keyValue, updates) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const keyColumn = headers.indexOf(key);

  if (keyColumn === -1) {
    throw new Error('Column not found: ' + key);
  }

  if (sheet.getLastRow() < 2) return false;

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, headers.length)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][keyColumn]) === String(keyValue)) {
      Object.keys(updates).forEach(field => {
        const column = headers.indexOf(field);

        if (column !== -1) {
          sheet.getRange(i + 2, column + 1).setValue(updates[field]);
        }
      });

      return true;
    }
  }

  return false;
}

function findRecord_(sheetName, key, value) {
  const records = getRecords_(sheetName);

  return records.find(record =>
    String(record[key] || '').trim() === String(value || '').trim()
  ) || null;
}

function makeId_(prefix) {
  return prefix + '-' +
    Utilities.getUuid().replace(/-/g, '').substring(0, 10).toUpperCase();
}

function cleanText_(value, maxLength) {
  return String(value === undefined || value === null ? '' : value)
    .trim()
    .substring(0, maxLength || 500);
}

function normalizePhone_(phone) {
  return String(phone || '').replace(/[\s()-]/g, '');
}

function validateCoordinates_(latitude, longitude) {
  const lat = Number(latitude);
  const lng = Number(longitude);

  return Number.isFinite(lat) &&
    Number.isFinite(lng) &&
    lat >= -90 && lat <= 90 &&
    lng >= -180 && lng <= 180;
}

function parseBoolean_(value) {
  return value === true ||
    String(value).toLowerCase() === 'true' ||
    String(value) === '1';
}

function now_() {
  return new Date().toISOString();
}

function jsonResponse_(data) {
  return data;
}


/* =============================================================
   CUSTOMER MANAGEMENT
============================================================= */

function createCustomer(payload) {
  return withScriptLock_(function () {
  payload = payload || {};

  const name = cleanText_(
    payload.FullName || payload.CustomerName || payload.customerName,
    120
  );

  const phone = cleanText_(
    payload.Phone || payload.phone,
    40
  );

  const email = cleanText_(
    payload.Email || payload.email,
    150
  );

  if (!name) {
    return { success: false, message: 'Customer name is required.' };
  }

  if (!phone) {
    return { success: false, message: 'Phone number is required.' };
  }

  const normalizedPhone = normalizePhone_(phone);

  const existing = getRecords_('Customers').find(customer =>
    normalizePhone_(customer.Phone) === normalizedPhone
  );

  if (existing) {
    return {
      success: true,
      existing: true,
      customer: existing,
      message: 'Existing customer profile retrieved.'
    };
  }

  const timestamp = now_();

  const customer = {
    CustomerID: makeId_('CUS'),
    FullName: name,
    Phone: phone,
    Email: email,
    CreatedAt: timestamp,
    UpdatedAt: timestamp,
    AccountStatus: 'Active'
  };

  appendRecord_('Customers', customer);

  return {
    success: true,
    existing: false,
    customer: customer,
    message: 'Customer registered successfully.'
  };

  });
}

function getCustomerProfile(payload) {
  payload = payload || {};

  const id = payload.CustomerID || payload.customerId;
  const phone = payload.Phone || payload.phone;

  let customer = id
    ? findRecord_('Customers', 'CustomerID', id)
    : null;

  if (!customer && phone) {
    const normalized = normalizePhone_(phone);

    customer = getRecords_('Customers').find(item =>
      normalizePhone_(item.Phone) === normalized
    ) || null;
  }

  if (!customer) {
    return { success: false, message: 'Customer not found.' };
  }

  return {
    success: true,
    customer: customer,
    locations: getCustomerLocations({
      CustomerID: customer.CustomerID
    })
  };
}

function updateCustomerProfile(payload) {
  return withScriptLock_(function () {
  payload = payload || {};

  const id = payload.CustomerID;

  if (!id) {
    return { success: false, message: 'CustomerID is required.' };
  }

  const customer = findRecord_('Customers', 'CustomerID', id);

  if (!customer) {
    return { success: false, message: 'Customer not found.' };
  }

  const updates = {
    UpdatedAt: now_()
  };

  if (payload.FullName !== undefined) {
    updates.FullName = cleanText_(payload.FullName, 120);
  }

  if (payload.Phone !== undefined) {
    updates.Phone = cleanText_(payload.Phone, 40);
  }

  if (payload.Email !== undefined) {
    updates.Email = cleanText_(payload.Email, 150);
  }

  updateRecord_('Customers', 'CustomerID', id, updates);

  return {
    success: true,
    message: 'Customer profile updated.'
  };

  });
}

function savePrimaryLocation(payload) {
  payload = payload || {};

  const customerId = cleanText_(
    payload.CustomerID || payload.customerId,
    100
  );

  const address = cleanText_(
    payload.Address || payload.address,
    500
  );

  const latitude = payload.Latitude !== undefined
    ? payload.Latitude : payload.latitude;

  const longitude = payload.Longitude !== undefined
    ? payload.Longitude : payload.longitude;

  const label = cleanText_(
    payload.Label || payload.label || 'Primary Location',
    80
  );

  if (!customerId) {
    return { success: false, message: 'CustomerID is required.' };
  }

  if (!findRecord_('Customers', 'CustomerID', customerId)) {
    return { success: false, message: 'Customer does not exist.' };
  }

  if (!address) {
    return { success: false, message: 'Delivery address is required.' };
  }

  if (!validateCoordinates_(latitude, longitude)) {
    return {
      success: false,
      message: 'Valid map coordinates are required. Select a map pin or use GPS.'
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const existingLocations = getRecords_('UserLocations')
      .filter(location =>
        String(location.CustomerID) === customerId
      );

    existingLocations.forEach(location => {
      if (parseBoolean_(location.IsPrimary)) {
        updateRecord_(
          'UserLocations',
          'LocationID',
          location.LocationID,
          {
            IsPrimary: false,
            UpdatedAt: now_()
          }
        );
      }
    });

    const timestamp = now_();

    const location = {
      LocationID: makeId_('LOC'),
      CustomerID: customerId,
      Label: label,
      Address: address,
      Latitude: Number(latitude),
      Longitude: Number(longitude),
      IsPrimary: true,
      CreatedAt: timestamp,
      UpdatedAt: timestamp
    };

    appendRecord_('UserLocations', location);

    return {
      success: true,
      location: location,
      message: 'Primary delivery location saved.'
    };

  } finally {
    lock.releaseLock();
  }
}

function getCustomerLocations(payload) {
  payload = payload || {};

  const customerId = payload.CustomerID || payload.customerId;

  let locations = getRecords_('UserLocations');

  if (customerId) {
    locations = locations.filter(location =>
      String(location.CustomerID) === String(customerId)
    );
  }

  return locations;
}


/*
 * Returns primary locations joined to their customer profiles.
 * Used by bootstrap() and the Google Maps frontend.
 */

function getCustomerMapLocations_() {
  const customers = getRecords_('Customers');
  const locations = getRecords_('UserLocations');

  const customerIndex = {};

  customers.forEach(customer => {
    const id = String(customer.CustomerID || '');
    if (id) customerIndex[id] = customer;
  });

  return locations
    .filter(location => {
      const customer = customerIndex[
        String(location.CustomerID || '')
      ];

      return parseBoolean_(location.IsPrimary) &&
        customer &&
        String(customer.AccountStatus || 'Active').toLowerCase() === 'active';
    })
    .map(location => {
      const customer = customerIndex[String(location.CustomerID)];

      return {
        LocationID: String(location.LocationID || ''),
        CustomerID: String(location.CustomerID || ''),
        CustomerName: String(customer.FullName || ''),
        Phone: String(customer.Phone || ''),
        Email: String(customer.Email || ''),
        Label: String(location.Label || ''),
        Address: String(location.Address || ''),
        Latitude: Number(location.Latitude),
        Longitude: Number(location.Longitude),
        IsPrimary: true,
        UpdatedAt: location.UpdatedAt || location.CreatedAt || ''
      };
    })
    .filter(location =>
      validateCoordinates_(location.Latitude, location.Longitude)
    );
}


/* =============================================================
   ORDER MANAGEMENT
============================================================= */

function getOrders(payload) {
  payload = payload || {};

  let orders = getRecords_('Orders');

  if (payload.CustomerID) {
    orders = orders.filter(order =>
      String(order.CustomerID) === String(payload.CustomerID)
    );
  }

  if (payload.Status) {
    orders = orders.filter(order =>
      String(order.Status).toLowerCase() ===
      String(payload.Status).toLowerCase()
    );
  }

  orders.sort((a, b) =>
    new Date(b.CreatedAt || 0) - new Date(a.CreatedAt || 0)
  );

  return orders;
}

function createOrder(payload) {
  return withScriptLock_(function () {
  payload = payload || {};

  const customerName = cleanText_(
    payload.CustomerName || payload.FullName || payload.customerName,
    120
  );

  const phone = cleanText_(payload.Phone || payload.phone, 40);

  const address = cleanText_(
    payload.Address || payload.address,
    500
  );

  const quantity = Number(
    payload.Quantity !== undefined ? payload.Quantity : payload.quantity
  );

  const price = Number(
    payload.Price !== undefined ? payload.Price : payload.price
  );

  const latitude = payload.Latitude !== undefined
    ? payload.Latitude : payload.latitude;

  const longitude = payload.Longitude !== undefined
    ? payload.Longitude : payload.longitude;

  const notes = cleanText_(payload.Notes || payload.notes, 1000);

  if (!customerName) {
    return { success: false, message: 'Customer name is required.' };
  }

  if (!phone) {
    return { success: false, message: 'Phone number is required.' };
  }

  if (!address) {
    return { success: false, message: 'Delivery address is required.' };
  }

  if (!Number.isFinite(quantity) || quantity <= 0) {
    return { success: false, message: 'Enter a valid milk quantity.' };
  }

  if (!Number.isFinite(price) || price < 0) {
    return { success: false, message: 'Enter a valid order price.' };
  }

  const customerId = cleanText_(
    payload.CustomerID || payload.customerId,
    100
  );

  const locationId = cleanText_(
    payload.LocationID || payload.locationId,
    100
  );

  if (customerId &&
      !findRecord_('Customers', 'CustomerID', customerId)) {
    return { success: false, message: 'Customer account was not found.' };
  }

  if (locationId) {
    const location = findRecord_(
      'UserLocations',
      'LocationID',
      locationId
    );

    if (!location ||
        String(location.CustomerID) !== String(customerId)) {
      return {
        success: false,
        message: 'The selected location does not belong to this customer.'
      };
    }
  }

  const hasCoordinates =
    latitude !== '' && latitude !== null && latitude !== undefined &&
    longitude !== '' && longitude !== null && longitude !== undefined;

  if (hasCoordinates && !validateCoordinates_(latitude, longitude)) {
    return { success: false, message: 'Invalid delivery coordinates.' };
  }

  const timestamp = now_();

  const order = {
    OrderID: makeId_('IM'),
    CustomerName: customerName,
    Phone: phone,
    Address: address,
    Latitude: hasCoordinates ? Number(latitude) : '',
    Longitude: hasCoordinates ? Number(longitude) : '',
    Quantity: quantity,
    Status: 'Pending',
    CreatedAt: timestamp,
    CustomerID: customerId,
    LocationID: locationId,
    Price: price,
    Notes: notes,
    PaymentStatus: 'Unpaid',
    PaymentReference: '',
    AssignedStaff: '',
    UpdatedAt: timestamp
  };

  appendRecord_('Orders', order);

  appendRecord_('OrderStatusHistory', {
    HistoryID: makeId_('HIS'),
    OrderID: order.OrderID,
    PreviousStatus: '',
    NewStatus: 'Pending',
    ChangedBy: 'System',
    Notes: 'Order created.',
    CreatedAt: timestamp
  });

  return {
    success: true,
    order: order,
    message: 'Delivery order created successfully.'
  };

  });
}

function updateOrderStatus(payload) {
  return withScriptLock_(function () {
  payload = payload || {};

  const orderId = payload.OrderID || payload.orderId;
  const newStatus = cleanText_(
    payload.Status || payload.status,
    40
  );

  const changedBy = cleanText_(
    payload.ChangedBy || payload.changedBy || 'Staff',
    120
  );

  const notes = cleanText_(payload.Notes || payload.notes, 500);

  const allowedStatuses = [
    'Pending',
    'Confirmed',
    'Preparing',
    'Out for Delivery',
    'Delivered',
    'Cancelled'
  ];

  if (!orderId || !newStatus) {
    return {
      success: false,
      message: 'OrderID and new status are required.'
    };
  }

  if (allowedStatuses.indexOf(newStatus) === -1) {
    return { success: false, message: 'Invalid order status.' };
  }

  const order = findRecord_('Orders', 'OrderID', orderId);

  if (!order) {
    return { success: false, message: 'Order not found.' };
  }

  const timestamp = now_();

  updateRecord_('Orders', 'OrderID', orderId, {
    Status: newStatus,
    UpdatedAt: timestamp
  });

  appendRecord_('OrderStatusHistory', {
    HistoryID: makeId_('HIS'),
    OrderID: orderId,
    PreviousStatus: order.Status,
    NewStatus: newStatus,
    ChangedBy: changedBy,
    Notes: notes,
    CreatedAt: timestamp
  });

  return {
    success: true,
    message: 'Order status updated.',
    OrderID: orderId,
    Status: newStatus
  };

  });
}

function getOrderHistory(payload) {
  payload = payload || {};

  const orderId = payload.OrderID || payload.orderId;

  if (!orderId) return [];

  return getRecords_('OrderStatusHistory')
    .filter(item => String(item.OrderID) === String(orderId))
    .sort((a, b) =>
      new Date(a.CreatedAt) - new Date(b.CreatedAt)
    );
}


/* =============================================================
   DASHBOARD STATISTICS
============================================================= */

function getDashboardStats_() {
  const orders = getRecords_('Orders');
  const customers = getRecords_('Customers');

  const pendingOrders = orders.filter(order =>
    String(order.Status).toLowerCase() === 'pending'
  ).length;

  const deliveredOrders = orders.filter(order =>
    String(order.Status).toLowerCase() === 'delivered'
  ).length;

  const activeOrders = orders.filter(order => {
    const status = String(order.Status).toLowerCase();

    return [
      'confirmed',
      'preparing',
      'out for delivery'
    ].indexOf(status) !== -1;
  }).length;

  const totalQuantity = orders.reduce((sum, order) =>
    sum + (Number(order.Quantity) || 0), 0
  );

  const totalRevenue = orders.reduce((sum, order) => {
    const price = Number(order.Price) || 0;

    const isCancelled =
      String(order.Status).toLowerCase() === 'cancelled';

    return sum + (isCancelled ? 0 : price);
  }, 0);

  return {
    totalOrders: orders.length,
    pendingOrders: pendingOrders,
    deliveredOrders: deliveredOrders,
    activeOrders: activeOrders,
    totalCustomers: customers.length,
    totalQuantity: totalQuantity,
    totalRevenue: totalRevenue
  };
}


/* =============================================================
   MESSAGING
============================================================= */

function messageForClient_(record) {
  return {
    MessageID: record.MessageID || '',
    CustomerID: record.CustomerID || '',
    sender: record.Sender || '',
    senderName: record.SenderName || '',
    message: record.Message || '',
    time: record.CreatedAt || '',
    readStatus: record.ReadStatus || 'Unread'
  };
}

function getMessages(payload) {
  payload = payload || {};

  const customerId = payload.CustomerID || payload.customerId;

  if (!customerId) {
    return {
      success: false,
      message: 'CustomerID is required.'
    };
  }

  return getRecords_('Messages')
    .filter(message =>
      String(message.CustomerID) === String(customerId)
    )
    .sort((a, b) =>
      new Date(a.CreatedAt) - new Date(b.CreatedAt)
    )
    .map(messageForClient_);
}

function sendMessage(payload) {
  return withScriptLock_(function () {
  payload = payload || {};

  const customerId = cleanText_(
    payload.CustomerID || payload.customerId,
    100
  );

  const message = cleanText_(
    payload.Message || payload.message,
    3000
  );

  const sender = cleanText_(
    payload.Sender || payload.sender || 'Customer',
    40
  );

  const senderName = cleanText_(
    payload.SenderName || payload.senderName || sender,
    120
  );

  if (!customerId || !message) {
    return {
      success: false,
      message: 'CustomerID and message are required.'
    };
  }

  if (!findRecord_('Customers', 'CustomerID', customerId)) {
    return { success: false, message: 'Customer not found.' };
  }

  const record = {
    MessageID: makeId_('MSG'),
    CustomerID: customerId,
    Sender: sender,
    SenderName: senderName,
    Message: message,
    CreatedAt: now_(),
    ReadStatus: 'Unread'
  };

  appendRecord_('Messages', record);

  return {
    success: true,
    message: messageForClient_(record)
  };

  });
}

function getConversationDirectory_() {
  const customers = getRecords_('Customers');
  const messages = getRecords_('Messages');

  return customers.map(customer => {
    const customerMessages = messages.filter(message =>
      String(message.CustomerID) === String(customer.CustomerID)
    );

    const latest = customerMessages.length
      ? customerMessages[customerMessages.length - 1]
      : null;

    return {
      CustomerID: customer.CustomerID,
      CustomerName: customer.FullName,
      Phone: customer.Phone,
      Email: customer.Email,
      MessageCount: customerMessages.length,
      LastMessage: latest ? latest.Message : '',
      LastMessageAt: latest ? latest.CreatedAt : ''
    };
  }).sort((a, b) =>
    new Date(b.LastMessageAt || 0) -
    new Date(a.LastMessageAt || 0)
  );
}

function getConversations() {
  return getConversationDirectory_();
}

function summarizeConversation(payload) {
  payload = payload || {};

  const customerId = payload.CustomerID || payload.customerId;

  if (!customerId) {
    return { success: false, message: 'CustomerID is required.' };
  }

  const messages = getMessages({ CustomerID: customerId });

  if (!Array.isArray(messages)) return messages;

  return {
    success: true,
    CustomerID: customerId,
    messageCount: messages.length,
    messages: messages
  };
}


/* =============================================================
   BOOTSTRAP
   The frontend should call this when loading the dashboard.
============================================================= */

function bootstrap(payload) {
  const conversations = getConversationDirectory_();

  const allMessages = getRecords_('Messages');
  const messagesByCustomer = {};

  allMessages.forEach(record => {
    const customerId = String(record.CustomerID || '');

    if (!customerId) return;

    if (!messagesByCustomer[customerId]) {
      messagesByCustomer[customerId] = [];
    }

    messagesByCustomer[customerId].push(
      messageForClient_(record)
    );
  });

  Object.keys(messagesByCustomer).forEach(customerId => {
    messagesByCustomer[customerId].sort((a, b) =>
      new Date(a.time) - new Date(b.time)
    );
  });

  return {
    success: true,
    stats: getDashboardStats_(),
    orders: getOrders(),
    customers: getRecords_('Customers'),
    locations: getCustomerMapLocations_(),
    conversations: conversations,
    messages: messagesByCustomer
  };
}


/* =============================================================
   API ROUTER
============================================================= */

function api(action, payload) {
  payload = payload || {};

  switch (String(action || '').toLowerCase()) {

    case 'setup':
    case 'setupim aradatabase':
    case 'setupimaradatabase':
      return setupImaraDatabase();

    case 'bootstrap':
      return bootstrap(payload);

    case 'createcustomer':
      return createCustomer(payload);

    case 'getcustomerprofile':
      return getCustomerProfile(payload);

    case 'updatecustomerprofile':
      return updateCustomerProfile(payload);

    case 'saveprimarylocation':
      return savePrimaryLocation(payload);

    case 'getcustomerlocations':
      return {
        success: true,
        locations: getCustomerLocations(payload)
      };

    case 'getorders':
      return {
        success: true,
        orders: getOrders(payload)
      };

    case 'createorder':
      return createOrder(payload);

    case 'updateorderstatus':
      return updateOrderStatus(payload);

    case 'getorderhistory':
      return {
        success: true,
        history: getOrderHistory(payload)
      };

    case 'getmessages':
      return {
        success: true,
        messages: getMessages(payload)
      };

    case 'sendmessage':
      return sendMessage(payload);

    case 'getconversations':
      return {
        success: true,
        conversations: getConversations()
      };

    case 'summarizeconversation':
      return summarizeConversation(payload);

    case 'getdashboardstats':
      return {
        success: true,
        stats: getDashboardStats_()
      };

    default:
      return {
        success: false,
        message: 'Unknown API action: ' + action
      };
  }
}


/* =============================================================
   HTTP / JSONP HANDLERS
============================================================= */

function parsePayload_(value) {
  if (!value) return {};

  if (typeof value === 'object') return value;

  try {
    return JSON.parse(value);
  } catch (error) {
    return {};
  }
}

function jsonpResponse_(data, callback) {
  const json = JSON.stringify({
    ok: true,
    data: data
  });

  if (callback) {
    const safeCallback = String(callback);

    if (!/^[a-zA-Z_$][0-9a-zA-Z_$.]*$/.test(safeCallback)) {
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: false,
          error: 'Invalid JSONP callback.'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(safeCallback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};

    const action = params.action || params.method || 'bootstrap';

    const payload = parsePayload_(
      params.payload || params.data
    );

    const result = api(action, payload);

    return jsonpResponse_(result, params.callback);

  } catch (error) {
    return jsonpResponse_({
      success: false,
      message: error.message || String(error)
    }, e && e.parameter ? e.parameter.callback : '');
  }
}

function doPost(e) {
  try {
    let body = {};

    if (e && e.postData && e.postData.contents) {
      body = parsePayload_(e.postData.contents);
    }

    const action = body.action || body.method || 'bootstrap';
    const payload = body.payload || body.data || {};

    const result = api(action, payload);

    return jsonResponse_(result);

  } catch (error) {
    return jsonResponse_({
      success: false,
      message: error.message || String(error)
    });
  }
}


/* =============================================================
   BACKEND TEST
============================================================= */

function testImaraBackend() {
  const result = bootstrap({});

  Logger.log(JSON.stringify(result, null, 2));

  return {
    success: true,
    customers: result.customers.length,
    locations: result.locations.length,
    orders: result.orders.length,
    stats: result.stats
  };
}
