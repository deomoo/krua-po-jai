// ════════════════════════════════════════════════════════════════
//  ครัวพอใจตามสั่ง — Google Apps Script Backend
//  วางทั้งหมดนี้ใน Google Apps Script แล้ว Deploy as Web App
// ════════════════════════════════════════════════════════════════

// ── 🔧 ตั้งค่าตรงนี้ก่อนใช้งาน ──────────────────────────────────
const CONFIG = {
  SHEET_ID:                '1uuHY6s4hhGxr_mJsop3CcegwBPQ0GLH-T_-7BjE-0oE',
  // LINE Messaging API (ใช้แทน LINE Notify ที่ปิดตัวแล้ว มี.ค. 2568)
  LINE_CHANNEL_ACCESS_TOKEN: 'YzlvR2hhPmXiT9QaWV9jx8xs+2j5VuMsuOq98sO1cDhIL/BtPlf6DsCw79nMh3+ePVeR0GU90+7PB/2XR5xe0xFWpXkeJmdtmJgNf/DuQSj5ham5xqA1kdgEwcfwOl1vldwWdnTKiwqZWhDO6MTE3AdB04t89/1O/w1cDnyilFU=',
  LINE_OWNER_USER_ID:        'Uccc1b5db6798aa769a8c4158a410b1f8',
  LINE_ADMIN_IDS:            ['Uccc1b5db6798aa769a8c4158a410b1f8', 'U90ab2da186f1ce82e4688bd6ebef8845'],  // เพิ่ม User ID แอดมินได้ที่นี่
  LINE_GROUP_ID:             'C224b763190c944f8b89f55ddc49bf566',
  GEMINI_API_KEY:            'AIzaSyAKStG9zR_z-g106Di53aVhf1Ra5TPjsgc',
  TIMEZONE:                'Asia/Bangkok',
  SHOP_NAME:               'ครัวพอใจตามสั่ง',
  DELIVERY_RADIUS_KM:      3,
};
// ─────────────────────────────────────────────────────────────────


// ════════════════════════════════════════════════════════════════
//  HTTP ENTRY POINTS
// ════════════════════════════════════════════════════════════════

/** รับ POST request จาก LIFF */
function doPost(e) {
  try {
    // ป้องกัน LINE verification request ที่ไม่มี postData
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput('OK');
    }
    const data = JSON.parse(e.postData.contents);

    // ── LINE Webhook Event (จาก LINE Platform) ──
    if (data.events && Array.isArray(data.events)) {
      data.events.forEach(ev => {
        // บันทึก Group ID เฉพาะครั้งแรกที่ยังไม่ได้ตั้งค่า
        if (ev.source && ev.source.type === 'group' && !CONFIG.LINE_GROUP_ID) {
          logError('GROUP_ID_FOUND', '✅ Group ID คือ: ' + ev.source.groupId);
        }
        // AI Order Bot — รับข้อความจากลูกค้าใน LINE OA (direct chat เท่านั้น)
        if (ev.type === 'message' && ev.message && ev.message.type === 'text'
            && ev.source && ev.source.type === 'user') {
          handleLineMessage(ev);
        }
      });
      return ContentService.createTextOutput('OK');
    }

    const action = data.action;
    let result;

    if      (action === 'submitOrder')       result = submitOrder(data);
    else if (action === 'getMenu')           result = getMenuItems();
    else if (action === 'checkStock')        result = checkStock(data.menuId);
    else if (action === 'updateOrderStatus') result = updateOrderStatus(data);
    else if (action === 'updateOrderItems')  result = updateOrderItems(data);
    else if (action === 'toggleMenu')        result = toggleMenu(data);
    else if (action === 'addMenu')           result = addMenu(data);
    else if (action === 'updateMenu')        result = updateMenu(data);
    else if (action === 'deleteMenu')        result = deleteMenu(data);
    else if (action === 'toggleIngredient')  result = toggleIngredient(data);
    else if (action === 'addExpense')        result = addExpenseFromAdmin(data);
    else if (action === 'deleteExpense')     result = deleteExpense(data);
    else if (action === 'addManualOrder')    result = addManualOrder(data);
    else if (action === 'addIncome')         result = addIncome(data);
    else result = { status: 'error', message: 'Unknown action: ' + action };

    return jsonResponse(result);
  } catch (err) {
    logError('doPost', err);
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/** รับ GET request (โหลดเมนู, ตรวจสต็อก) */
function doGet(e) {
  try {
    const action = e.parameter.action || 'menu';
    if (action === 'menu')        return jsonResponse(getMenuItems());
    if (action === 'stock')       return jsonResponse(getAllStock());
    if (action === 'orders')      return jsonResponse(getOrders());
    if (action === 'adminMenu')   return jsonResponse(getAdminMenu());
    if (action === 'ingredients') return jsonResponse(getIngredients());
    if (action === 'expenses')    return jsonResponse(getExpenses(e.parameter));
    if (action === 'incomes')     return jsonResponse(getIncomes(e.parameter));
    return jsonResponse({ status: 'ok', message: 'ครัวพอใจ API ready' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════════
//  1. SUBMIT ORDER
// ════════════════════════════════════════════════════════════════

function submitOrder(data) {
  const ss        = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const orderSheet = getOrCreateSheet(ss, 'ออเดอร์',
    ['เลขออเดอร์','วันที่','เวลา','ช่องทาง','ประเภท','ยอดรวม',
     'รายการเมนู','Add-ons','โน้ต','ที่อยู่','ชื่อลูกค้า','สถานะ','ประเภทบิล','LINE User ID']);

  // สร้างเลขออเดอร์ (reset ทุกวัน)
  const now0     = new Date();
  const todayStr0 = Utilities.formatDate(now0, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const allRows0  = orderSheet.getLastRow() > 1
    ? orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, 2).getValues()
    : [];
  const todayCount = allRows0.filter(r => {
    try { return Utilities.formatDate(new Date(r[1]), CONFIG.TIMEZONE, 'dd/MM/yyyy') === todayStr0; }
    catch(e) { return String(r[1]) === todayStr0; }
  }).length;
  const orderNum = '#' + String(todayCount + 1).padStart(3, '0');

  const now      = new Date();
  const dateStr  = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const timeStr  = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');

  // สรุปรายการจากเมนูปกติ
  const regularItems = Array.isArray(data.items) ? data.items : [];
  const itemsSummary = regularItems
    .map(i => `${i.name} x${i.qty}${i.addon ? ' (+'+i.addon+')' : ''}`)
    .join(', ');

  // รายการพิเศษ/นอกเมนู (ราคายังไม่ระบุ)
  const customList    = (data.customItems || []).filter(c => c.name && c.name.trim());
  const customSummary = customList.map(c => `${c.name.trim()} x1 [?]`).join(', ');
  const fullSummary   = [itemsSummary, customSummary].filter(Boolean).join(', ');
  const hasCustom     = true; // แอดมินยืนยันราคาทุกออเดอร์

  const total = regularItems.reduce((s, i) => s + i.price * i.qty, 0);

  // บันทึกลง Sheet
  orderSheet.appendRow([
    orderNum,
    dateStr,
    timeStr,
    'LINE OA',
    (data.deliveryType === 'deliver' || data.deliveryType === 'delivery') ? 'ส่ง'
      : data.deliveryType === 'dine-in' ? 'ทานที่ร้าน' : 'รับหน้าร้าน',
    total,            // ยอดรวมเฉพาะรายการที่มีราคา (รายการ [?] คิดหลังแอดมินระบุราคา)
    fullSummary,      // รายการทั้งหมด รวม [?]
    regularItems.filter(i => i.addon).map(i => i.addon).join(', '),
    data.note || '',
    data.address || '',
    data.customerName || 'ลูกค้า LINE',
    'รอทำ',
    data.paymentMethod === 'transfer' ? '💳 เงินโอน' : '💵 เงินสด',
    data.userId || '',
  ]);

  // หักสต็อกวัตถุดิบ (เฉพาะรายการปกติ)
  deductStock(ss, regularItems);

  // แจ้งเตือนเจ้าของร้าน
  sendOwnerNotify(orderNum, data, fullSummary, total, timeStr, hasCustom);

  // ไม่ส่งข้อความยืนยันให้ลูกค้าตอนรับออเดอร์ — ส่งเฉพาะตอนเสร็จ/ส่งเท่านั้น

  return { status: 'success', orderNum: orderNum, total: total, hasCustom: hasCustom };
}


// ════════════════════════════════════════════════════════════════
//  1b. ADD MANUAL ORDER (สร้างออเดอร์จากแอดมิน — โทรศัพท์/หน้าร้าน)
// ════════════════════════════════════════════════════════════════

function addManualOrder(data) {
  const ss         = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const orderSheet = getOrCreateSheet(ss, 'ออเดอร์',
    ['เลขออเดอร์','วันที่','เวลา','ช่องทาง','ประเภท','ยอดรวม',
     'รายการเมนู','Add-ons','โน้ต','ที่อยู่','ชื่อลูกค้า','สถานะ','ประเภทบิล','LINE User ID']);

  // สร้างเลขออเดอร์ (reset ทุกวัน)
  const now       = new Date();
  const todayStr  = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const allRows   = orderSheet.getLastRow() > 1
    ? orderSheet.getRange(2, 1, orderSheet.getLastRow() - 1, 2).getValues()
    : [];
  const todayCount = allRows.filter(r => {
    try { return Utilities.formatDate(new Date(r[1]), CONFIG.TIMEZONE, 'dd/MM/yyyy') === todayStr; }
    catch(e) { return String(r[1]) === todayStr; }
  }).length;
  const orderNum = '#' + String(todayCount + 1).padStart(3, '0');

  const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const timeStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');

  // แปลง type
  const typeLabel = (data.type === 'delivery') ? 'ส่ง'
    : (data.type === 'dine-in') ? 'ทานที่ร้าน' : 'รับหน้าร้าน';

  // แปลง payment
  const payLabel = (data.payment === 'transfer') ? '💳 เงินโอน'
    : (data.payment === 'installment') ? '🔄 คนละครึ่ง'
    : (data.payment === 'other') ? '🧾 อื่นๆ' : '💵 เงินสด';

  orderSheet.appendRow([
    orderNum,
    dateStr,
    timeStr,
    data.channel || 'โทรศัพท์',
    typeLabel,
    data.total || 0,
    data.items || '',
    '',                            // Add-ons (รวมอยู่ใน items string แล้ว)
    data.note || '',
    data.address || '',
    data.customerName || 'ลูกค้า',
    'รอทำ',
    payLabel,
    '',                            // LINE User ID (ไม่มีสำหรับออเดอร์โทรศัพท์)
  ]);

  // แจ้งเตือนเจ้าของร้าน
  try {
    const fakeData = {
      customerName: data.customerName || 'ลูกค้า',
      deliveryType: data.type,
      note: data.note || '',
      address: data.address || '',
      paymentMethod: data.payment,
    };
    sendOwnerNotify(orderNum, fakeData, data.items || '', data.total || 0, timeStr, false);
  } catch(e) { logError('addManualOrder.notify', e); }

  return { status: 'success', orderNum: orderNum, total: data.total || 0 };
}


// ════════════════════════════════════════════════════════════════
//  2. GET MENU ITEMS (อ่านจาก Sheet "เมนู")
// ════════════════════════════════════════════════════════════════

function getMenuItems() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet เมนู' };

  const rows  = sheet.getDataRange().getValues();
  const items = [];

  for (let i = 2; i < rows.length; i++) {   // เริ่มจากแถว 3 (index 2) ข้ามหัวตาราง
    const row = rows[i];
    if (!row[0]) continue;                   // ข้ามแถวว่าง
    items.push({
      id:       row[0],                      // A: รหัสเมนู
      name:     row[1],                      // B: ชื่อเมนู
      category: row[2],                      // C: หมวดหมู่
      price:    Number(row[3]) || 0,         // D: ราคา
      cost:     Number(row[4]) || 0,         // E: ต้นทุน
      status:   row[6] || 'มี',              // G: สถานะ (มี/หมด)
      addons:   row[7] ? String(row[7]).split(',').map(a => a.trim()) : [],
      image:    row[10] || '',               // K: URL รูปภาพ (คอลัมน์ K)
    });
  }

  return { status: 'success', items: items };
}


// ════════════════════════════════════════════════════════════════
//  3. CHECK / DEDUCT STOCK
// ════════════════════════════════════════════════════════════════

function checkStock(menuId) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (rows[i][0] === menuId) {
      return { status: 'success', inStock: rows[i][6] === 'มี', name: rows[i][1] };
    }
  }
  return { status: 'not_found' };
}

function getAllStock() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error' };
  const rows = sheet.getDataRange().getValues();
  const stock = {};
  for (let i = 2; i < rows.length; i++) {
    if (rows[i][0]) stock[rows[i][0]] = rows[i][6] === 'มี';
  }
  return { status: 'success', stock: stock };
}

/** หักสต็อกวัตถุดิบเมื่อมีออเดอร์ */
function deductStock(ss, items) {
  const ingSheet = ss.getSheetByName('วัตถุดิบ');
  if (!ingSheet) return;

  // Map: ชื่อเมนู → วัตถุดิบที่ใช้ (กำหนดเองได้)
  const RECIPE = {
    'MN-01': [{ ing: 'ING-01', qty: 1 }],   // ผัดกระเพราหมูกรอบ → หมูกรอบ 1 จาน
    'MN-02': [{ ing: 'ING-03', qty: 1 }],   // กุ้งสดผัดผักรวม → กุ้งสด 1 จาน
    'MN-03': [{ ing: 'ING-14', qty: 1 }],   // ทะเลรวม
    'MN-06': [{ ing: 'ING-03', qty: 1 }],   // ผัดไทยกุ้งสด
    'MN-10': [{ ing: 'ING-10', qty: 2 }],   // ไข่เจียว → ไข่ไก่ 2 ฟอง
  };

  const rows = ingSheet.getDataRange().getValues();

  items.forEach(item => {
    const recipe = RECIPE[item.id];
    if (!recipe) return;
    recipe.forEach(r => {
      for (let i = 2; i < rows.length; i++) {
        if (rows[i][0] === r.ing) {
          const curr = Number(rows[i][4]) || 0;
          const newVal = Math.max(0, curr - r.qty * item.qty);
          ingSheet.getRange(i + 1, 5).setValue(newVal);  // E = สต็อกคงเหลือ
          // อัปเดตสถานะ
          const minStock = Number(rows[i][5]) || 0;
          const status   = newVal === 0 ? 'หมด' : newVal <= minStock ? 'ใกล้หมด' : 'มี';
          ingSheet.getRange(i + 1, 8).setValue(status);  // H = สถานะ
          // อัปเดตเมนูถ้าวัตถุดิบหมด
          if (newVal === 0) updateMenuStatus(ss, r.ing, 'หมด');
          break;
        }
      }
    });
  });
}

/** อัปเดตสถานะเมนูเมื่อวัตถุดิบหมด */
function updateMenuStatus(ss, ingId, newStatus) {
  const menuSheet = ss.getSheetByName('เมนู');
  if (!menuSheet) return;
  const ING_TO_MENU = {
    'ING-01': ['MN-01'],
    'ING-03': ['MN-02', 'MN-06'],
    'ING-14': ['MN-03'],
  };
  const affected = ING_TO_MENU[ingId] || [];
  const rows = menuSheet.getDataRange().getValues();
  affected.forEach(menuId => {
    for (let i = 2; i < rows.length; i++) {
      if (rows[i][0] === menuId) {
        menuSheet.getRange(i + 1, 7).setValue(newStatus);  // G = สถานะ
        break;
      }
    }
  });
}


// ════════════════════════════════════════════════════════════════
//  4. LINE MESSAGING API — แจ้งเตือนเจ้าของร้าน (Push Message)
// ════════════════════════════════════════════════════════════════

function sendOwnerNotify(orderNum, data, itemsSummary, total, timeStr, hasCustom) {
  const token  = CONFIG.LINE_CHANNEL_ACCESS_TOKEN;
  const userId = CONFIG.LINE_OWNER_USER_ID;
  if (!token  || token.startsWith('วาง'))  return;
  if (!userId || userId.startsWith('วาง')) return;
  const adminIds = CONFIG.LINE_ADMIN_IDS || [userId];

  const typeEmoji   = data.deliveryType === 'deliver' ? '🛵 ส่งถึงบ้าน' : '🏪 รับหน้าร้าน';
  const payEmoji    = data.paymentMethod === 'transfer' ? '📲 เงินโอน' : '💵 เงินสด';
  const addressLine = data.address ? `\n📍 ${data.address}` : '';
  const noteLine    = data.note    ? `\n📝 ${data.note}` : '';
  const customLine  = hasCustom    ? `\n⭐ มีรายการพิเศษ กรุณาระบุราคาก่อนแจ้งลูกค้า` : '';

  const msgText =
`🔔 ออเดอร์ใหม่! ${orderNum}
⏰ ${timeStr} น.  ${typeEmoji}${addressLine}
👤 ${data.customerName || 'ลูกค้า LINE'}

🍽 ${itemsSummary}${noteLine}

💰 ยอดปัจจุบัน ฿${total}  ${payEmoji}${customLine}
──────────────────
⚡ กรุณารับออเดอร์ในแอดมิน`;

  // ส่งไปกลุ่มก่อน (ถ้าตั้งค่าไว้) — แชทกลุ่มจะขึ้นล่าสุดในลิสต์ + มีเสียง
  const groupId = CONFIG.LINE_GROUP_ID;
  if (groupId && !groupId.startsWith('วาง') && groupId !== '') {
    linePush(groupId, msgText);
  } else {
    // ส่งหาแอดมินทุกคน
    adminIds.forEach(id => { if (id) linePush(id, msgText); });
  }
}


// ════════════════════════════════════════════════════════════════
//  แจ้งลูกค้าทันทีเมื่อออเดอร์เข้าระบบ
// ════════════════════════════════════════════════════════════════

function sendOrderConfirm(lineUserId, orderNum, itemsSummary, total, data, hasCustom) {
  const typeText   = data.deliveryType === 'deliver' ? '🛵 ส่งถึงบ้าน' : '🏪 รับหน้าร้าน';
  const payText    = data.paymentMethod === 'transfer' ? '📲 ชำระด้วยเงินโอน' : '💵 ชำระด้วยเงินสด';
  const addrLine   = data.address ? `\n📍 ${data.address}` : '';
  const noteLine   = data.note    ? `\n📝 ${data.note}` : '';
  const msgText =
`🛒 รับออเดอร์แล้ว! ${orderNum}
${typeText}${addrLine}

🍽 ${itemsSummary}${noteLine}

⏳ รอร้านยืนยันออเดอร์และแจ้งราคาสักครู่นะครับ 🙏`;

  linePush(lineUserId, msgText);
}


// ── แจ้งลูกค้าหลังแอดมินยืนยันออเดอร์แล้ว ──
function sendCustomPriceNotify(lineUserId, orderNum, itemsSummary, total) {
  const msgText =
`✅ ร้านยืนยันออเดอร์แล้ว! ${orderNum}

🍽 ${itemsSummary}

💰 ยอดรวม ฿${total}
──────────────────
กำลังเตรียมอาหาร รอสักครู่นะครับ 🙏`;
  linePush(lineUserId, msgText);
}

function getOwnerUserId() {
  Logger.log('ℹ️ วิธีง่ายกว่า: เปิด LINE OA Manager → Chats → เลือกแชทของตัวเอง');
  Logger.log('   แล้วดู URL: https://manager.line.biz/account/.../chat/U[xxxxxxxx]');
}


// ════════════════════════════════════════════════════════════════
//  5. AUTO RESET STOCK (ตั้ง Trigger ทุกวันอาทิตย์ 20:00)
// ════════════════════════════════════════════════════════════════

function weeklyAutoReset() {
  const ss       = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const ingSheet = ss.getSheetByName('วัตถุดิบ');
  if (!ingSheet) return;

  const rows = ingSheet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const defaultStock = Number(rows[i][3]) || 0;
    ingSheet.getRange(i + 1, 5).setValue(defaultStock);
    ingSheet.getRange(i + 1, 8).setValue('มี');
  }

  const menuSheet = ss.getSheetByName('เมนู');
  if (menuSheet) {
    const mRows = menuSheet.getDataRange().getValues();
    for (let i = 2; i < mRows.length; i++) {
      if (mRows[i][0]) menuSheet.getRange(i + 1, 7).setValue('มี');
    }
  }

  sendOwnerNotify('AUTO', { deliveryType: 'none' }, '♻️ รีเซ็ตสต็อกสำเร็จ', 0,
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm'));
}

/** ทดสอบ Gemini API — รันจาก Editor แล้วดู Logs */
function testGemini() {
  const result = callGemini('กะเพราหมู 1 ส่งบ้าน — นี่คือออเดอร์อาหารไหม? ตอบแค่ใช่/ไม่ใช่');
  Logger.log('Gemini response: ' + result);
  if (!result) Logger.log('❌ Gemini ไม่ตอบ — ตรวจสอบ API Key หรือ quota');
  else Logger.log('✅ Gemini ทำงานได้!');
}

/** สร้าง Trigger อัตโนมัติ — รันครั้งเดียวตอนตั้งค่า */
function createWeeklyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'weeklyAutoReset') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyAutoReset')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(20)
    .create();
  Logger.log('✅ Trigger สร้างเรียบร้อย: ทุกวันอาทิตย์ 20:00');
}


// ════════════════════════════════════════════════════════════════
//  6. GET ORDERS — ดึงรายการออเดอร์ทั้งหมด (สำหรับหน้า Admin)
// ════════════════════════════════════════════════════════════════

function getOrders() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('ออเดอร์');
  if (!sheet || sheet.getLastRow() < 2) return { status: 'success', orders: [] };

  const rows   = sheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    // แปลง Date object → string ก่อน (Sheets มักแปลง string วันที่เป็น Date อัตโนมัติ)
    const dateVal = r[1] instanceof Date
      ? Utilities.formatDate(r[1], CONFIG.TIMEZONE, 'dd/MM/yyyy')
      : String(r[1] || '');
    const timeVal = r[2] instanceof Date
      ? Utilities.formatDate(r[2], CONFIG.TIMEZONE, 'HH:mm')
      : String(r[2] || '');

    orders.push({
      orderNum:     String(r[0]),
      date:         dateVal,     // B: วันที่  → "13/05/2026"
      time:         timeVal,     // C: เวลา   → "15:49"
      type:         r[4],        // E: ประเภท (ส่ง/รับหน้าร้าน)
      total:        r[5],        // F: ยอดรวม
      items:        r[6],        // G: รายการเมนู
      note:         r[8],        // I: โน้ต
      address:      r[9],        // J: ที่อยู่
      customerName: r[10],       // K: ชื่อลูกค้า
      status:       r[11],       // L: สถานะ
      payment:      r[12] || '', // M: วิธีชำระ
      lineUserId:   r[13] || '', // N: LINE User ID
      rowIndex:     i + 1,
    });
  }
  orders.reverse(); // ล่าสุดขึ้นก่อน
  return { status: 'success', orders: orders };
}


// ════════════════════════════════════════════════════════════════
//  7b. UPDATE ORDER ITEMS — แก้ไขรายการอาหารในออเดอร์
// ════════════════════════════════════════════════════════════════

function updateOrderItems(data) {
  // data: { orderNum, items, addons, total, sendNotify?, lineUserId? }
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('ออเดอร์');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet ออเดอร์' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === String(data.orderNum).trim()) {
      sheet.getRange(i + 1, 6).setValue(Number(data.total) || 0);
      sheet.getRange(i + 1, 7).setValue(data.items  || '');
      sheet.getRange(i + 1, 8).setValue(data.addons || '');
      // ส่ง LINE แจ้งยอดลูกค้า (กรณีมีรายการพิเศษและแอดมินกำหนดราคาแล้ว)
      // ไม่ส่ง sendCustomPriceNotify — แจ้งลูกค้าเฉพาะตอนเสร็จ/ส่งเท่านั้น
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'ไม่พบออเดอร์ ' + data.orderNum };
}


// ════════════════════════════════════════════════════════════════
//  7. UPDATE ORDER STATUS — เปลี่ยนสถานะ + แจ้ง LINE ลูกค้า
// ════════════════════════════════════════════════════════════════

function updateOrderStatus(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('ออเดอร์');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet ออเดอร์' };

  const rows = sheet.getDataRange().getValues();

  // ใช้ rowIndex ตรงๆ เพราะเลขออเดอร์ซ้ำกันข้ามวันได้
  if (data.rowIndex && parseInt(data.rowIndex) > 1) {
    const ri = parseInt(data.rowIndex);
    if (ri <= rows.length && String(rows[ri - 1][0]) === String(data.orderNum)) {
      sheet.getRange(ri, 12).setValue(data.status);
      const lineUserId = rows[ri - 1][13] || '';
      const orderType  = String(rows[ri - 1][4] || '');
      if (lineUserId) sendCustomerNotify(lineUserId, data.orderNum, data.status, String(rows[ri - 1][6]), orderType, Number(rows[ri - 1][5]) || 0);
      return { status: 'success', message: 'อัปเดตสถานะเรียบร้อย' };
    }
  }

  // fallback: ค้นหาจากเลขออเดอร์ (กรณีไม่มี rowIndex)
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] !== data.orderNum) continue;
    sheet.getRange(i + 1, 12).setValue(data.status);
    const lineUserId  = rows[i][13] || '';
    const orderType   = String(rows[i][4] || '');
    if (lineUserId) {
      sendCustomerNotify(lineUserId, data.orderNum, data.status, String(rows[i][6]), orderType, Number(rows[i][5]) || 0);
    }
    return { status: 'success', message: 'อัปเดตสถานะเรียบร้อย' };
  }
  return { status: 'error', message: 'ไม่พบออเดอร์ ' + data.orderNum };
}


// ════════════════════════════════════════════════════════════════
//  8. SEND CUSTOMER NOTIFY — แจ้งลูกค้าเมื่อสถานะเปลี่ยน
// ════════════════════════════════════════════════════════════════

function sendCustomerNotify(lineUserId, orderNum, status, items, orderType, total) {
  // ส่งแจ้งลูกค้าเฉพาะตอนเสร็จ / ส่ง / ยกเลิก เท่านั้น (ประหยัดโควต้า LINE)
  if (!['ส่งแล้ว','ยกเลิก'].includes(status)) return;

  const shopName  = CONFIG.SHOP_NAME;
  const isDeliver = orderType === 'ส่ง';
  const totalLine = total ? `\n💰 ยอดรวม ฿${total}` : '';

  const msgs = {
    'เสร็จแล้ว': isDeliver
      ?
`🎉 อาหารพร้อมแล้ว! ${orderNum}

🍽 ${items}
${totalLine}

🛵 กำลังจัดส่งให้คุณนะครับ
ขอบคุณที่อุดหนุน ${shopName} 🙏`
      :
`🎉 อาหารพร้อมแล้ว! ${orderNum}

🍽 ${items}
${totalLine}

🏃 มารับได้เลยครับ
ขอบคุณที่อุดหนุน ${shopName} 🙏`,

    'ส่งแล้ว':
`✅ ส่งอาหารแล้ว! ${orderNum}

🍽 ${items}
${totalLine}

🙏 ขอบคุณที่อุดหนุน ${shopName} นะครับ`,

    'ยกเลิก':
`❌ ออเดอร์ถูกยกเลิก ${orderNum}

😔 ขออภัยในความไม่สะดวกครับ`,
  };

  const msgText = msgs[status] || `📋 ${orderNum} — อัปเดตสถานะ: ${status}`;
  linePush(lineUserId, msgText);
}


// ════════════════════════════════════════════════════════════════
//  LINE Push Message — helper ใช้ร่วมกัน
// ════════════════════════════════════════════════════════════════

function linePush(to, text) {
  const token = CONFIG.LINE_CHANNEL_ACCESS_TOKEN;
  if (!token || token.startsWith('วาง')) {
    logError('linePush', 'TOKEN ยังไม่ได้ตั้งค่า'); return;
  }
  if (!to || to.startsWith('วาง') || to === '') {
    logError('linePush', 'USER ID ยังไม่ได้ตั้งค่า (to=' + to + ')'); return;
  }
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method:  'post',
      headers: {
        'Content-Type':  'application/json',
        'Authorization': 'Bearer ' + token,
      },
      payload: JSON.stringify({ to, messages: [{ type: 'text', text }] }),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code !== 200) {
      logError('linePush', 'HTTP ' + code + ' | to=' + to.substring(0,8) + '... | ' + body);
    }
  } catch(e) {
    logError('linePush', e.toString());
  }
}

// ── ทดสอบ LINE Push — รันใน Apps Script Editor แล้วดูผลใน Execution Log ──
function testLinePush() {
  const token  = CONFIG.LINE_CHANNEL_ACCESS_TOKEN;
  const userId = CONFIG.LINE_OWNER_USER_ID;
  Logger.log('Token  (10 ตัวแรก): ' + token.substring(0, 10) + '...');
  Logger.log('UserId (10 ตัวแรก): ' + userId.substring(0, 10) + '...');

  const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method:  'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
    payload: JSON.stringify({ to: userId, messages: [{ type: 'text', text: '🧪 ทดสอบ LINE ครัวพอใจ — ถ้าเห็นข้อความนี้ระบบปกติ ✅' }] }),
    muteHttpExceptions: true,
  });
  Logger.log('HTTP Status: ' + res.getResponseCode());
  Logger.log('Response: ' + res.getContentText());
}


// ════════════════════════════════════════════════════════════════
//  9. ADMIN MENU — ดึงเมนูทั้งหมด (สำหรับหน้าจัดการเมนู Admin)
// ════════════════════════════════════════════════════════════════

function getAdminMenu() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet เมนู' };

  const rows  = sheet.getDataRange().getValues();
  const items = [];

  for (let i = 2; i < rows.length; i++) {
    const row = rows[i];
    if (!row[0]) continue;
    items.push({
      id:       row[0],                      // A: รหัสเมนู
      name:     row[1],                      // B: ชื่อเมนู
      category: row[2],                      // C: หมวดหมู่
      price:    Number(row[3]) || 0,         // D: ราคา
      cost:     Number(row[4]) || 0,         // E: ต้นทุน
      status:   row[6] || 'มี',              // G: สถานะ
      addons:   row[7] ? String(row[7]) : '', // H: Add-ons (string คั่นด้วย ,)
      image:    row[10] || '',               // K: รูปภาพ
      rowIndex: i + 1,                       // สำหรับ update/delete
    });
  }
  return { status: 'success', items: items };
}


// ════════════════════════════════════════════════════════════════
//  10. TOGGLE MENU STATUS — สลับสถานะ มี/หมด
// ════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════
//  INGREDIENT MANAGEMENT — วัตถุดิบ มี/หมด
// ════════════════════════════════════════════════════════════════

function getIngredients() {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('วัตถุดิบ');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet วัตถุดิบ' };
  const rows   = sheet.getDataRange().getValues();
  const items  = [];
  for (let i = 2; i < rows.length; i++) {
    const id   = String(rows[i][0] || '').trim();
    const name = String(rows[i][1] || '').trim();
    if (!id && !name) continue;
    items.push({
      id:     id,
      name:   name,
      unit:   String(rows[i][2] || ''),
      status: String(rows[i][7] || 'มี')   // คอลัมน์ H = สถานะ
    });
  }
  return { status: 'ok', items };
}

function toggleIngredient(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('วัตถุดิบ');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet วัตถุดิบ' };
  const rows = sheet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== String(data.id).trim()) continue;
    const newStatus = data.status || (rows[i][7] === 'มี' ? 'หมด' : 'มี');
    sheet.getRange(i + 1, 8).setValue(newStatus);  // คอลัมน์ H
    return { status: 'success', id: data.id, newStatus };
  }
  return { status: 'error', message: 'ไม่พบวัตถุดิบ ' + data.id };
}

function toggleMenu(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet เมนู' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (rows[i][0] !== data.menuId) continue;
    const newStatus = data.status || (rows[i][6] === 'มี' ? 'หมด' : 'มี');
    sheet.getRange(i + 1, 7).setValue(newStatus);  // G = สถานะ
    return { status: 'success', menuId: data.menuId, newStatus: newStatus };
  }
  return { status: 'error', message: 'ไม่พบเมนู ' + data.menuId };
}


// ════════════════════════════════════════════════════════════════
//  11. ADD MENU — เพิ่มเมนูใหม่
// ════════════════════════════════════════════════════════════════

function addMenu(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet เมนู' };

  // สร้าง ID อัตโนมัติ MN-XX
  const rows   = sheet.getDataRange().getValues();
  let maxNum   = 0;
  for (let i = 2; i < rows.length; i++) {
    const id = String(rows[i][0] || '');
    const match = id.match(/MN-(\d+)/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1]));
  }
  const newId = 'MN-' + String(maxNum + 1).padStart(2, '0');

  // เพิ่มแถวใหม่ (ตรงกับโครงสร้าง A-K)
  sheet.appendRow([
    newId,                         // A: รหัสเมนู
    data.name     || '',           // B: ชื่อเมนู
    data.category || '',           // C: หมวดหมู่
    Number(data.price) || 0,       // D: ราคา
    Number(data.cost)  || 0,       // E: ต้นทุน
    '',                            // F: (สำรอง)
    data.status   || 'มี',         // G: สถานะ
    data.addons   || '',           // H: Add-ons
    '', '', '',                    // I, J, K (รูปภาพว่างไว้ก่อน)
  ]);

  return { status: 'success', menuId: newId, message: 'เพิ่มเมนูเรียบร้อย' };
}


// ════════════════════════════════════════════════════════════════
//  12. UPDATE MENU — แก้ไขข้อมูลเมนู
// ════════════════════════════════════════════════════════════════

function updateMenu(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet เมนู' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (rows[i][0] !== data.menuId) continue;
    const row = i + 1;  // 1-based row number
    if (data.name     !== undefined) sheet.getRange(row, 2).setValue(data.name);
    if (data.category !== undefined) sheet.getRange(row, 3).setValue(data.category);
    if (data.price    !== undefined) sheet.getRange(row, 4).setValue(Number(data.price) || 0);
    if (data.cost     !== undefined) sheet.getRange(row, 5).setValue(Number(data.cost) || 0);
    if (data.status   !== undefined) sheet.getRange(row, 7).setValue(data.status);
    if (data.addons   !== undefined) sheet.getRange(row, 8).setValue(data.addons);
    return { status: 'success', message: 'แก้ไขเมนูเรียบร้อย' };
  }
  return { status: 'error', message: 'ไม่พบเมนู ' + data.menuId };
}


// ════════════════════════════════════════════════════════════════
//  13. DELETE MENU — ลบเมนู
// ════════════════════════════════════════════════════════════════

function deleteMenu(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName('เมนู');
  if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet เมนู' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 2; i < rows.length; i++) {
    if (rows[i][0] !== data.menuId) continue;
    sheet.deleteRow(i + 1);
    return { status: 'success', message: 'ลบเมนูเรียบร้อย' };
  }
  return { status: 'error', message: 'ไม่พบเมนู ' + data.menuId };
}


// ════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#2563EB')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold');
  }
  return sheet;
}

function logError(fn, err) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = getOrCreateSheet(ss, 'Error Log', ['เวลา','Function','Error']);
    sheet.appendRow([new Date(), fn, err.toString()]);
  } catch(e) { /* silently fail */ }
}


// ════════════════════════════════════════════════════════════════
//  AI ORDER BOT — LINE OA + Gemini
// ════════════════════════════════════════════════════════════════

function handleLineMessage(ev) {
  const userId     = ev.source.userId;
  const text       = ev.message.text.trim();
  const replyToken = ev.replyToken;
  const msgId      = ev.message.id;

  // ป้องกัน LINE ส่ง webhook ซ้ำ (Webhook redelivery)
  const cache    = CacheService.getScriptCache();
  const dedupKey = 'msg_' + msgId;
  if (cache.get(dedupKey)) return; // ประมวลผลแล้ว — ข้ามเลย
  cache.put(dedupKey, '1', 300);   // lock 5 นาที

  const isAdmin = CONFIG.LINE_ADMIN_IDS.includes(userId);

  // ══ โหมดหลังครัว (แอดมินเท่านั้น) ══
  if (text === 'หลังครัว') {
    if (!isAdmin) return; // ลูกค้าทั่วไปพิมพ์มา → เพิกเฉย
    setConvState(userId, { step: 'admin' });
    lineReply(replyToken, '🔑 โหมดหลังครัว\n\nจัดการอะไรดีครับ?\nพิมพ์ เช่น\n• ปิดหมูกรอบ\n• เปิดกุ้งสด\n• ดูวัตถุดิบ\n• ออก (จบโหมด)');
    return;
  }

  // ── อยู่ในโหมดหลังครัว ──
  const adminState = getConvState(userId);
  if (isAdmin && adminState && adminState.step === 'admin') {
    if (['ออก','exit','จบ','เสร็จ'].includes(text)) {
      clearConvState(userId);
      lineReply(replyToken, 'ออกจากโหมดหลังครัวแล้วครับ');
      return;
    }
    // เปิดวัตถุดิบทั้งหมด
    if (text === 'เปิดทั้งหมด') {
      const ing = getIngredients();
      (ing.items || []).forEach(i => toggleIngredient({ id: i.id, status: 'มี' }));
      lineReply(replyToken, 'เปิดวัตถุดิบทั้งหมดแล้วครับ ✅');
      return;
    }
    if (text === 'ดูวัตถุดิบ') {
      const ing  = getIngredients();
      const list = (ing.items || []).map(i =>
        (i.status === 'หมด' ? '❌ ' : '✅ ') + i.name).join('\n');
      lineReply(replyToken, '📋 สถานะวัตถุดิบ\n\n' + list);
      return;
    }
    // ปิด/เปิด X
    const closeMatch = text.match(/^ปิด(.+)/);
    const openMatch  = text.match(/^เปิด(.+)/);
    if (closeMatch || openMatch) {
      const query     = (closeMatch || openMatch)[1].trim();
      const newStatus = closeMatch ? 'หมด' : 'มี';
      const found     = getIngredientStatus(query);
      if (!found.found) {
        lineReply(replyToken, `ไม่พบ "${query}" ในลิสต์วัตถุดิบครับ`);
        return;
      }
      toggleIngredient({ id: found.id, status: newStatus });
      const action = newStatus === 'หมด' ? 'ปิด' : 'เปิด';
      lineReply(replyToken, `${action}${found.name}แล้วครับ ✅`);
      return;
    }
    // ── บันทึกรายจ่าย (แอดมินพิมพ์ "จ่าย ..." หรือ "รายจ่าย ...") ──
    if (/^(จ่าย|รายจ่าย|ค่าใช้จ่าย)\s+/i.test(text) || /^\d/.test(text) && isAdmin) {
      if (/^(จ่าย|รายจ่าย|ค่าใช้จ่าย)\s+/i.test(text)) {
        try {
          const expenses = parseExpenseWithGemini(text);
          if (expenses && expenses.length > 0) {
            const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
            expenses.forEach(exp => addExpenseRow(ss, exp));
            const summary = expenses.map(e => `• ${e.category}: ${e.name} ฿${e.amount}`).join('\n');
            const total   = expenses.reduce((s, e) => s + (e.amount || 0), 0);
            lineReply(replyToken, `✅ บันทึกรายจ่ายแล้ว!\n\n${summary}\n\n💸 รวม ฿${total}`);
          } else {
            lineReply(replyToken, '❓ ไม่เข้าใจรายจ่าย ลองพิมพ์เช่น\n"จ่าย หมูกรอบ 200 น้ำมัน 80"');
          }
        } catch(err) {
          logError('expense', err.toString());
          lineReply(replyToken, '❌ บันทึกไม่สำเร็จ: ' + err.toString());
        }
        return;
      }
    }

    // ── แอดมินสั่งออเดอร์แทนลูกค้า ──
    const adminOrder = parseAdminOrderWithGemini(text);
    if (adminOrder && adminOrder.isOrder) {
      try {
        const result = submitOrderFromAI({
          items:         adminOrder.items        || [],
          customItems:   adminOrder.customItems  || [],
          deliveryType:  adminOrder.deliveryType || 'pickup',
          address:       adminOrder.address      || '',
          paymentMethod: 'cash',
          customerName:  adminOrder.customerName || 'ลูกค้า (แอดมินสั่ง)',
          userId:        null,
          note:          adminOrder.time ? 'เวลา: ' + adminOrder.time : ''
        });
        if (result.status === 'success') {
          let msg = '✅ สร้างออเดอร์ ' + result.orderNum + ' แล้วครับ\n\n';
          msg += '👤 ' + (adminOrder.customerName || 'ลูกค้า') + '\n';
          msg += formatOrderSummary(adminOrder.items, adminOrder.customItems);
          if (adminOrder.address) msg += '\n📍 ' + adminOrder.address;
          if (adminOrder.time)    msg += '\n⏰ ' + adminOrder.time;
            lineReply(replyToken, msg);
        } else {
          lineReply(replyToken, '❌ สร้างออเดอร์ไม่สำเร็จ กรุณาลองใหม่');
        }
      } catch(err) {
        logError('adminOrder', err);
        lineReply(replyToken, '❌ เกิดข้อผิดพลาด: ' + err.toString());
      }
      return;
    }

    lineReply(replyToken, '🔑 โหมดหลังครัว\n\nพิมพ์ได้เลย:\n• ปิด/เปิดวัตถุดิบ เช่น "ปิดหมูกรอบ"\n• สั่งอาหาร เช่น "สมชาย กะเพราหมู 1 ส่งห้อง 201"\n• ดูวัตถุดิบ\n• ออก');
    return;
  }

  // คำสั่งพิเศษ
  if (['ยกเลิก','cancel','เริ่มใหม่','reset'].includes(text.toLowerCase())) {
    clearConvState(userId);
    lineReply(replyToken, '❌ ยกเลิกออเดอร์แล้วครับ\nพิมพ์ได้เลยถ้าอยากสั่งใหม่ 😊');
    return;
  }
  if (['เมนู','menu','ดูเมนู'].includes(text.toLowerCase())) {
    lineReply(replyToken, '🍽 เมนูวันนี้\n\n' + getMenuSummaryText() + '\n\nพิมพ์สั่งได้เลยครับ');
    return;
  }

  let state = getConvState(userId) || { step: 'idle' };

  // ── STEP: Gemini ดึงทุกอย่างจากข้อความเดียว → ส่งออเดอร์ทันที ──
  if (state.step === 'idle') {
    const parsed = parseOrderWithGemini(text);

    // ── ถามเรื่องวัตถุดิบ ──
    if (parsed && parsed.isIngredientQuery) {
      const queries = parsed.ingredientQueries || [];
      if (!queries.length) return; // ไม่มีในลิสต์ → แอดมินตอบเอง

      const results = queries.map(q => getIngredientStatus(q)).filter(r => r.found);
      if (!results.length) return; // ไม่พบเลย → แอดมินตอบเอง

      const available = results.filter(r => r.status !== 'หมด').map(r => r.name);
      const outOfStock = results.filter(r => r.status === 'หมด').map(r => r.name);

      let msg = '';
      if (outOfStock.length === 0) {
        // ทุกอย่างมี
        msg = available.join('กับ') + 'มีค่ะ';
      } else if (available.length === 0) {
        // ทุกอย่างหมด
        msg = outOfStock.join('กับ') + 'หมดค่ะ';
      } else {
        // บางอย่างหมด บางอย่างมี
        msg = outOfStock.join('กับ') + 'หมด มีแต่' + available.join('กับ') + 'ค่ะ';
      }
      lineReply(replyToken, msg);
      return;
    }

    if (!parsed || !parsed.isOrder || (!parsed.items.length && !parsed.customItems.length)) {
      return; // ไม่ใช่ออเดอร์ — ไม่ตอบ
    }
    clearConvState(userId);
    try {
      const profile  = lineGetProfile(userId);
      const custName = profile ? profile.displayName : 'ลูกค้า LINE';
      const result   = submitOrderFromAI({
        items:         parsed.items,
        customItems:   parsed.customItems || [],
        deliveryType:  parsed.deliveryType || 'pickup',
        address:       parsed.address      || '',
        paymentMethod: 'cash',
        customerName:  custName,
        userId:        userId,
        skipConfirm:   true  // bot ตอบด้วย lineReply แล้ว ไม่ต้องส่ง push ซ้อน
      });
      if (result.status === 'success') {
        // ไม่ส่งข้อความยืนยันให้ลูกค้าทันที — ร้านจะแจ้งตอนอาหารพร้อม/ส่งเท่านั้น (ประหยัดโควต้า LINE)
      } else {
        lineReply(replyToken, '❌ เกิดข้อผิดพลาด กรุณาลองใหม่หรือติดต่อร้านโดยตรงครับ');
      }
    } catch(err) {
      logError('handleLineMessage/submit', err);
      lineReply(replyToken, '❌ เกิดข้อผิดพลาด กรุณาลองใหม่ครับ');
    }
    return;
  }
}

// ── Conversation State (CacheService 30 นาที) ──
function getConvState(userId) {
  try { const d = CacheService.getScriptCache().get('ai_' + userId); return d ? JSON.parse(d) : null; }
  catch(e) { return null; }
}
function setConvState(userId, state) {
  try { CacheService.getScriptCache().put('ai_' + userId, JSON.stringify(state), 1800); } catch(e) {}
}
function clearConvState(userId) {
  try { CacheService.getScriptCache().remove('ai_' + userId); } catch(e) {}
}

// ── Gemini API (รองรับทั้ง AIzaSy... และ AQ. key) ──
function callGemini(prompt) {
  const key = CONFIG.GEMINI_API_KEY;
  // AQ. key ใช้ header / AIzaSy key ใช้ query param
  const isAQKey = key.startsWith('AQ.');
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent'
    + (isAQKey ? '' : '?key=' + key);
  const headers = { 'Content-Type': 'application/json' };
  if (isAQKey) {
    headers['Authorization'] = 'Bearer ' + key;
  }

  try {
    const res  = UrlFetchApp.fetch(url, {
      method: 'POST', headers, muteHttpExceptions: true,
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 800 }
      })
    });
    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code !== 200) {
      logError('callGemini', 'HTTP ' + code + ' | ' + body.slice(0, 300));
      return '';
    }
    const json = JSON.parse(body);
    return json.candidates?.[0]?.content?.parts?.[0]?.text || '';
  } catch(err) { logError('callGemini', err); return ''; }
}

// ── แปลงข้อความลูกค้า → JSON ออเดอร์ ──
function parseOrderWithGemini(text) {
  const menuItems = getMenuItemsForAI();
  const menuStr   = menuItems.map(m => m.name + ' (฿' + m.price + ')').join(', ');

  const ingItems = getIngredientNamesForAI();
  const ingStr   = ingItems.join(', ');

  const prompt =
`คุณเป็นผู้ช่วยร้านอาหารไทย ชื่อ "ครัวพอใจ"

เมนูในร้าน: ${menuStr}
วัตถุดิบในร้าน: ${ingStr}

ข้อความจากลูกค้า: "${text}"

ตอบเป็น JSON เท่านั้น ห้ามอธิบายเพิ่ม:
{
  "isOrder": true/false,
  "isIngredientQuery": true/false,
  "ingredientQueries": ["วัตถุดิบที่เกี่ยวข้องทั้งหมด"],
  "items": [{"name":"ชื่อเมนูในร้าน","qty":จำนวน,"addon":"add-on ถ้ามี","price":ราคา}],
  "customItems": [{"name":"ชื่อรายการที่ไม่มีในเมนู","qty":จำนวน}],
  "deliveryType": "delivery หรือ pickup หรือ dine-in",
  "address": "ที่อยู่จัดส่ง ถ้ามี"
}

กฎ:
- ถ้าลูกค้าถามว่า "มี...ไหม" "...หมดไหม" "...ยังมีอยู่ไหม" → isIngredientQuery: true
- ingredientQueries ให้ map กับวัตถุดิบในลิสต์ที่เกี่ยวข้องทั้งหมด เช่น "ทะเล" → ["หมึกสด","กุ้งสด","ทะเลรวม"] (เฉพาะที่มีในลิสต์)
- ถ้าถามวัตถุดิบเดี่ยว เช่น "หมูกรอบ" → ["หมูกรอบ"]
- ถ้าไม่มีวัตถุดิบในลิสต์เลย → ingredientQueries: []
- ถ้าชื่อใกล้เคียงเมนูในร้าน ให้ map ชื่อเมนูจริง + ราคาจริง
- "พิเศษ" = ขนาดใหญ่ ใส่ใน addon
- "ไข่ดาว" "ไข่เจียว" "ไข่ข้น" = addon (สะกด "ไข่ข้น" ตรงตามชื่อใน Add-ons เสมอ) ถ้ามีหลาย addon ให้คั่นด้วย ", " เช่น "ไข่ข้น, พิเศษ"
- ของที่ไม่มีในเมนู → customItems
- "ส่ง..." หรือ "จัดส่ง" → deliveryType: "delivery", ดึงที่อยู่ใส่ address
- "รับเอง" "รับที่ร้าน" → deliveryType: "pickup"
- "ทานที่ร้าน" → deliveryType: "dine-in"
- ถ้าไม่ระบุ → deliveryType: "pickup"
- ถ้าไม่ใช่ออเดอร์และไม่ใช่ถามวัตถุดิบ → isOrder: false, isIngredientQuery: false`;

  const resp = callGemini(prompt);
  try {
    const m = resp.match(/\{[\s\S]*\}/);
    return m ? JSON.parse(m[0]) : null;
  } catch(e) { return null; }
}

// ── แอดมินสั่งอาหารแทนลูกค้า ──
function parseAdminOrderWithGemini(text) {
  const menuItems = getMenuItemsForAI();
  const menuStr   = menuItems.map(m => m.name + ' (฿' + m.price + ')').join(', ');

  const prompt =
`คุณเป็นผู้ช่วยรับออเดอร์ร้านอาหาร แอดมินจะพิมพ์ข้อมูลออเดอร์ให้ลูกค้า

เมนูในร้าน: ${menuStr}

ข้อความจากแอดมิน: "${text}"

ตอบเป็น JSON เท่านั้น:
{
  "isOrder": true/false,
  "customerName": "ชื่อลูกค้า",
  "items": [{"name":"ชื่อเมนู","qty":จำนวน,"addon":"","price":ราคา}],
  "customItems": [{"name":"รายการนอกเมนู","qty":จำนวน}],
  "deliveryType": "delivery หรือ pickup หรือ dine-in",
  "address": "ที่อยู่หรือสถานที่",
  "time": "เวลาที่ระบุ ถ้ามี"
}

กฎ:
- ชื่อคนแรกที่ขึ้นต้นหรือต่อจากคำว่า "ชื่อ/ลูกค้า" = customerName
- ถ้าไม่มีชื่อ ให้ customerName: ""
- "ส่ง..." = deliveryType: "delivery"
- "รับที่ร้าน/รับเอง" = deliveryType: "pickup"
- "จอง/กิน/นั่งทาน" = deliveryType: "dine-in"
- เวลาที่ระบุเช่น "12:00" "บ่าย 2" = time
- ถ้าไม่ใช่ออเดอร์อาหาร = isOrder: false`;

  const resp = callGemini(prompt);
  try {
    const m = resp.match(/\{[\s\S]*\}/);
    return m ? JSON.parse(m[0]) : null;
  } catch(e) { return null; }
}

function getIngredientNamesForAI() {
  try {
    const rows  = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('วัตถุดิบ').getDataRange().getValues();
    const names = [];
    for (let i = 2; i < rows.length; i++) {
      if (rows[i][1]) names.push(String(rows[i][1]));
    }
    return names;
  } catch(e) { return []; }
}

function getIngredientStatus(query) {
  // คืน { found: true/false, name, status }
  try {
    const rows = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('วัตถุดิบ').getDataRange().getValues();
    const q    = query.toLowerCase().trim();
    for (let i = 2; i < rows.length; i++) {
      const name = String(rows[i][1] || '').toLowerCase().trim();
      if (!name) continue;
      if (name.includes(q) || q.includes(name)) {
        return { found: true, id: String(rows[i][0]), name: String(rows[i][1]), status: String(rows[i][7] || 'มี') };
      }
    }
    return { found: false };
  } catch(e) { return { found: false }; }
}

function getMenuItemsForAI() {
  try {
    const rows = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('เมนู').getDataRange().getValues();
    const items = [];
    for (let i = 2; i < rows.length; i++) {
      if (rows[i][1] && rows[i][6] !== 'หมด') items.push({ name: String(rows[i][1]), price: Number(rows[i][3]) || 0 });
    }
    return items;
  } catch(e) { return []; }
}

function getMenuSummaryText() {
  return getMenuItemsForAI().map(m => '• ' + m.name + ' ฿' + m.price).join('\n') || 'ยังไม่มีเมนู';
}

function formatOrderSummary(items, customItems) {
  let lines = [], total = 0;
  (items || []).forEach(i => {
    const addon = i.addon ? ' (+' + i.addon + ')' : '';
    const sub   = (i.price || 0) * (i.qty || 1);
    total += sub;
    lines.push('• ' + i.name + addon + ' x' + i.qty);
  });
  (customItems || []).forEach(i => lines.push('⭐ ' + i.name + ' x' + i.qty));
  return lines.join('\n');
}

// ── สร้างออเดอร์เข้าระบบ ──
function submitOrderFromAI(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  // ส่ง items เป็น array of objects ตามที่ submitOrder() ต้องการ
  const rawItems = (data.items || []).map(i => ({
    name:  i.name,
    qty:   Number(i.qty)   || 1,
    price: Number(i.price) || 0,
    addon: i.addon || ''
  }));
  // บวกราคา Add-ons จากตาราง ตั้งค่า (ไม่ใช้แค่ base price จาก Gemini)
  const items = rawItems.map(i=>{try{const m={};const sh=ss.getSheets().find(s=>s.getName().includes('ตั้งค่า'));if(sh&&sh.getLastRow()>=15)sh.getRange(15,2,sh.getLastRow()-14,4).getValues().forEach(r=>{if(r[1]&&String(r[3]).trim()==='เปิด')m[String(r[1]).trim()]=parseFloat(r[2])||0;});const extra=String(i.addon||'').split(/[,\s+]+/).reduce((s,n)=>s+(m[n.trim()]||0),0);return Object.assign({},i,{price:(i.price||0)+extra});}catch(e){return i;}});

  return submitOrder({
    action:        'submitOrder',
    items:         items,
    customItems:   data.customItems || [],
    deliveryType:  data.deliveryType,
    address:       data.address     || '',
    paymentMethod: data.paymentMethod,
    customerName:  data.customerName,
    userId:        data.userId,
    channel:       'LINE OA (AI)',
    note:          '',
    skipConfirm:   true   // AI bot ส่ง reply เองแล้ว ไม่ต้องส่งซ้ำ
  });
}

// ════════════════════════════════════════════════════════════════
//  EXPENSE — ระบบรายจ่าย
// ════════════════════════════════════════════════════════════════

const EXPENSE_CATS = ['วัตถุดิบ','แก๊ส/เชื้อเพลิง','บรรจุภัณฑ์','ค่าจ้าง','ค่าเช่า','อื่นๆ'];

function getExpenseSheet(ss) {
  return getOrCreateSheet(ss, 'รายจ่าย',
    ['วันที่','เวลา','หมวดหมู่','รายการ','จำนวนเงิน (฿)','หมายเหตุ','บันทึกโดย']);
}

function addExpenseRow(ss, exp) {
  const sheet   = getExpenseSheet(ss);
  const now     = new Date();
  const dateStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd/MM/yyyy');
  const timeStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm');
  sheet.appendRow([dateStr, timeStr, exp.category || 'อื่นๆ', exp.name || '', Number(exp.amount) || 0, exp.note || '', exp.by || 'แอดมิน']);
}

function addExpenseFromAdmin(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  addExpenseRow(ss, data);
  return { status: 'success' };
}

function deleteExpense(data) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getExpenseSheet(ss);
  const rows  = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(data.date) && String(rows[i][3]) === String(data.name) && String(rows[i][4]) === String(data.amount)) {
      sheet.deleteRow(i + 1);
      return { status: 'success' };
    }
  }
  return { status: 'error', message: 'ไม่พบรายการ' };
}

function getExpenses(params) {
  const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getExpenseSheet(ss);
  if (sheet.getLastRow() < 2) return { status: 'success', expenses: [] };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  let expenses = rows.map(r => ({
    date: String(r[0]), time: String(r[1]), category: String(r[2]),
    name: String(r[3]), amount: Number(r[4]) || 0, note: String(r[5]), by: String(r[6])
  })).filter(e => e.name);
  // filter by date range
  if (params && params.from) {
    expenses = expenses.filter(e => {
      try {
        const m = e.date.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
        if (!m) return false;
        const d = new Date(`${m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}T12:00:00+07:00`);
        const from = new Date(params.from + 'T00:00:00+07:00');
        const to   = new Date((params.to || params.from) + 'T23:59:59+07:00');
        return d >= from && d <= to;
      } catch(e) { return true; }
    });
  }
  return { status: 'success', expenses, categories: EXPENSE_CATS };
}

// ── Gemini parse รายจ่าย ──
function parseExpenseWithGemini(text) {
  const prompt =
`แยกรายจ่ายจากข้อความนี้: "${text}"

หมวดหมู่ที่มี: วัตถุดิบ, แก๊ส/เชื้อเพลิง, บรรจุภัณฑ์, ค่าจ้าง, ค่าเช่า, อื่นๆ

ตอบเป็น JSON array เท่านั้น:
[{"name":"ชื่อรายการ","amount":จำนวนเงิน,"category":"หมวดหมู่"}]

ถ้ามีหลายรายการให้แยกแต่ละรายการ เช่น "หมูกรอบ 200 น้ำมัน 80" → 2 รายการ`;

  const resp = callGemini(prompt);
  try {
    const m = resp.match(/\[[\s\S]*\]/);
    return m ? JSON.parse(m[0]) : null;
  } catch(e) { return null; }
}

// ── LINE Reply (ใช้ replyToken แทน push — ฟรี ไม่จำกัด) ──
function lineReply(replyToken, message) {
  const token = CONFIG.LINE_CHANNEL_ACCESS_TOKEN;
  if (!token || token.startsWith('วาง')) return;
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'POST', muteHttpExceptions: true,
      headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({ replyToken, messages: [{ type: 'text', text: message }] })
    });
  } catch(err) { logError('lineReply', err); }
}

// ── ดึงชื่อลูกค้าจาก LINE Profile ──
function lineGetProfile(userId) {
  try {
    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + userId, {
      headers: { 'Authorization': 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN },
      muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText());
  } catch(e) { return null; }
}
