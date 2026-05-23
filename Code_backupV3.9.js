/**
 * POS Noir Backend - v6.8 (Smart Binding + Column Protection + Turbo Auth + Original Logic)
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ POS Setup')
    .addItem('🔌 1. Привязать базу к кассе', 'initDatabase')
    .addToUi();
}

function initDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('DATABASE_ID', ss.getId());
  SpreadsheetApp.getUi().alert('✅ Успешно! Касса намертво привязана к этой таблице.');
}

function getDB() {
  const id = PropertiesService.getScriptProperties().getProperty('DATABASE_ID');
  if (!id) throw new Error("База не привязана! Нажмите POS Setup -> Привязать базу.");
  return SpreadsheetApp.openById(id);
}

// --- УСКОРЕННОЕ ПОЛУЧЕНИЕ НАСТРОЕК (С КЭШЕМ И ПРАВИЛЬНЫМИ КОЛОНКАМИ) ---
function getAppConfig() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get('APP_CONFIG_V4');
  
  if (cachedData) return JSON.parse(cachedData);

  const dash = getDB().getSheetByName("Dashboard");
  let cfg = { staff: [], folderId: "", invoiceFolderId: "", defaultVat: 0.16 };
  if (dash) {
    cfg.folderId = dash.getRange("N1").getValue().toString().trim();
    cfg.invoiceFolderId = dash.getRange("N2").getValue().toString().trim();
    let vatVal = dash.getRange("N3").getValue();
    cfg.defaultVat = typeof vatVal === 'number' ? vatVal : parseFloat(vatVal) || 0.16;
    
    const data = dash.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      let row = data[i];
      // Читаем из O=14 (UID), P=15 (Имя), Q=16 (ПИН)
      if(row[14] && String(row[14]).trim() !== "") { 
        const uidVal = String(row[14]).trim();
        cfg.staff.push({
          uid: uidVal, 
          name: row[15] ? String(row[15]).trim() : "", 
          pin: row[16] ? String(row[16]).trim() : "",
          role: uidVal.toUpperCase().startsWith('M') ? 'manager' : 'seller'
        });
      }
    }
  }
  
  cache.put('APP_CONFIG_V4', JSON.stringify(cfg), 900);
  // Кэш на 15 минут
  return cfg;
}

function dateStrToTs(dStr) {
  if (!dStr) return 0;
  let p = dStr.split('.');
  if (p.length !== 3) return 0;
  return new Date(p[2], p[1]-1, p[0]).getTime();
}

function getTrueLastRow(sheet) {
  const values = sheet.getRange("A:A").getValues();
  let lastRow = values.length;
  while (lastRow > 0 && !values[lastRow - 1][0]) {
    lastRow--;
  }
  return lastRow;
}

function doGet(e) {
  try {
    const config = getAppConfig();
    // ==========================================
    // ВАШ ОРИГИНАЛЬНЫЙ БЛОК ОТЧЕТОВ (НЕТРОНУТЫЙ)
    // ==========================================
    if (e && e.parameter && e.parameter.action === 'report') {
      const sheetTx = getDB().getSheetByName("Transactions");
      const dateStartReq = e.parameter.date ? e.parameter.date.replace(/[^\d\.]/g, '') : "";
      const dateEndReq = e.parameter.endDate ? e.parameter.endDate.replace(/[^\d\.]/g, '') : dateStartReq;
      const tsStart = dateStrToTs(dateStartReq);
      const tsEnd = dateStrToTs(dateEndReq);
      const sellerIdReq = e.parameter.seller_id;
      
      const data = sheetTx.getDataRange().getValues();
      let reportData = [];
      if (data.length >= 3) {
        let receiptsMap = {};
        for (let i = 2; i < data.length; i++) {
          let row = data[i];
          if (!row[0]) continue;
          
          let strDate = "", timeStr = "";
          if (row[0] instanceof Date) {
            strDate = Utilities.formatDate(row[0], "Asia/Almaty", "dd.MM.yyyy");
            timeStr = Utilities.formatDate(row[0], "Asia/Almaty", "HH:mm");
          } else {
            let cleanRaw = String(row[0]).replace(/[^\d\.\:\s]/g, '').trim();
            strDate = cleanRaw.split(' ')[0] || ""; 
            timeStr = cleanRaw.split(' ')[1] ? cleanRaw.split(' ')[1].substring(0, 5) : "";
          }

          let rowTs = dateStrToTs(strDate);
          let rowSellerId = row[16] ? row[16].toString() : ""; 

          if (rowTs >= tsStart && rowTs <= tsEnd && rowTs !== 0) {
            if (sellerIdReq && rowSellerId !== sellerIdReq) continue;
            let tid = row[1];
            if (!receiptsMap[tid]) {
             receiptsMap[tid] = { 
                 type: row[2], 
                 total: 0, 
                 cart: [], 
                 methodCode: row[11] || "cash", 
                 time: timeStr, 
                 date: strDate,
                 seller: rowSellerId || "Анон" // <--- ДОБАВЛЯЕМ ЭТУ СТРОКУ
             };
          }
            let qty = Number(row[5]) || 0;
            let price = Number(row[6]) || 0;
            receiptsMap[tid].cart.push({ name: row[4], qty: qty, price: price });
            receiptsMap[tid].total += (qty * price);
          }
        }
        reportData = Object.keys(receiptsMap).map(key => receiptsMap[key]);
      }
      // Возвращаем только отчет, без лишних данных
      return ContentService.createTextOutput(JSON.stringify({ report: reportData })).setMimeType(ContentService.MimeType.JSON);
    }

    // ==========================================
    // СБОРКА ТОВАРОВ И ПЕРСОНАЛЬНЫХ ИТОГОВ
    // ==========================================
    const sheetItems = getDB().getSheetByName("Items");
    const itemsData = sheetItems.getDataRange().getValues();
    let items = [];
    for (let i = 2; i < itemsData.length; i++) {
      let row = itemsData[i];
      if (row[0]) {
        items.push({
          id: row[0].toString(), 
          name: row[1] ? row[1].toString() : "Без названия", 
          category: row[2] ? row[2].toString() : "",
          barcode: row[3] ? row[3].toString() : "", 
          stock: parseFloat(row[4]) || 0, 
          cost: parseFloat(row[5]) || 0,
          price: parseFloat(row[6]) || 0, 
          img: row[10] ? row[10].toString() : ""
        });
      }
    }

    // --- ПЕРСОНАЛЬНАЯ КАССА (ДИНАМИЧЕСКИЙ ПОДСЧЕТ) ---
    const sheetTx = getDB().getSheetByName("Transactions");
    const txData = sheetTx.getDataRange().getValues();
    const reqUid = e.parameter.uid || ""; 
    const reqRole = e.parameter.role || "";
    const isManager = reqRole === 'manager';
    
    // Получаем начало и конец сегодняшнего дня (по времени Алматы)
    const now = new Date();
    const todayStr = Utilities.formatDate(now, "Asia/Almaty", "dd.MM.yyyy");
    const tsStartToday = dateStrToTs(todayStr);

    let personalTotals = { cash: 0, qr: 0, pos: 0, transfer: 0, red: 0 };

    if (txData.length >= 3) {
      let processedTids = new Set(); // Чтобы не дублировать товары из одного чека
      
      for (let i = 2; i < txData.length; i++) {
        let row = txData[i];
        if (!row[0]) continue;
        
        let rowStrDate = "";
        if (row[0] instanceof Date) {
          rowStrDate = Utilities.formatDate(row[0], "Asia/Almaty", "dd.MM.yyyy");
        } else {
          rowStrDate = String(row[0]).replace(/[^\d\.\:\s]/g, '').split(' ')[0] || ""; 
        }

        let rowTs = dateStrToTs(rowStrDate);
        let rowSellerId = row[16] ? row[16].toString().trim() : ""; 
        
        // Считаем только сегодняшние чеки
        if (rowTs === tsStartToday) {
          // Если это не менеджер, и чек не этого продавца - пропускаем
          if (!isManager && rowSellerId !== reqUid) continue;

          let tid = row[1];
          let txType = row[2];
          let qty = Number(row[5]) || 0;
          let price = Number(row[6]) || 0;
          let methodCode = row[11] ? row[11].toString().trim().toLowerCase() : "cash";
          
          let sum = qty * price;
          if (txType === 'return') sum = -sum; // Возвраты отнимаем из кассы

          // Маппинг методов оплаты
          if (methodCode.includes('cash')) personalTotals.cash += sum;
          else if (methodCode.includes('qr')) personalTotals.qr += sum;
          else if (methodCode.includes('pos_terminal') || methodCode.includes('card')) personalTotals.pos += sum;
          else if (methodCode.includes('transfer')) personalTotals.transfer += sum;
          else if (methodCode.includes('installment') || methodCode.includes('red')) personalTotals.red += sum;
        }
      }
    }

    // --- БЕЗОПАСНОСТЬ: ИСКЛЮЧАЕМ ПИН-КОДЫ ПЕРЕД ОТПРАВКОЙ ---
    const safeStaffList = config.staff.map(person => ({
      uid: person.uid,
      name: person.name,
      role: person.role
    }));

    // ФИНАЛЬНЫЙ ВЫВОД ДЛЯ ЗАГРУЗКИ КАССЫ
    return ContentService.createTextOutput(JSON.stringify({ 
      items: items, 
      staff: safeStaffList,
      totals: personalTotals, // Отдаем динамические итоги
      synonyms: getInvoiceSynonyms() 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ error: e.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const p = JSON.parse(e.postData.contents);

  // ==========================================
  // БЫСТРЫЙ ПУТЬ: АВТОРИЗАЦИЯ (ДО ВКЛЮЧЕНИЯ ОЧЕРЕДИ)
  // ==========================================
  if (p.action === 'verifyPin') {
    const cfg = getAppConfig();
    const user = cfg.staff.find(person => String(person.pin).trim() === String(p.pin).trim());
    
    if (user) {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: true, 
        user: { uid: user.uid, name: user.name, role: user.role } 
      })).setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ 
        success: false, 
        error: "Неверный ПИН-код" 
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ==========================================
  // МЕДЛЕННЫЙ ПУТЬ: ТРАНЗАКЦИИ (ВКЛЮЧАЕМ ОЧЕРЕДЬ)
  // ==========================================
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    if (p.action === 'income') {
      return ContentService.createTextOutput(JSON.stringify(processIncomes(p))).setMimeType(ContentService.MimeType.JSON);
    }
    
    if (p.action === 'upload_photo') {
      const folder = DriveApp.getFolderById(getAppConfig().folderId);
      const contentType = p.base64.split(',')[0].split(':')[1].split(';')[0];
      const blob = Utilities.newBlob(Utilities.base64Decode(p.base64.split(',')[1]), contentType, p.filename);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
      
      const sheetItems = getDB().getSheetByName("Items");
      const ids = sheetItems.getRange(3, 1, Math.max(1, getTrueLastRow(sheetItems)-2), 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0].toString() === p.item_id.toString()) {
          sheetItems.getRange(i + 3, 11).setValue(url);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true, url: url })).setMimeType(ContentService.MimeType.JSON);
    }

    const sheetTx = getDB().getSheetByName("Transactions");
    const ts = Utilities.formatDate(new Date(), "Asia/Almaty", "dd.MM.yyyy HH:mm:ss");
    let rows = p.cart.map(c => {
      let row = new Array(17).fill(""); 
      row[0] = ts; row[1] = p.tx_id; row[2] = p.tx_type; row[3] = c.item_id; row[4] = c.item_name;
      row[5] = c.qty; row[6] = c.price; row[8] = c.cost_price; row[10] = "Клиент"; row[11] = p.payment_method;
      row[13] = p.source; row[16] = p.seller_id; 
      return row;
    });

    if (rows.length > 0) {
      const lr = Math.max(1, getTrueLastRow(sheetTx));
      sheetTx.getRange(lr + 1, 1, rows.length, 17).setValues(rows);
    }
    return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: e.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Чтение синонимов шапок инвойса из листа Synonyms
 */
function getInvoiceSynonyms() {
  try {
    const sheet = getDB().getSheetByName("Synonyms");
    let synonyms = {};
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) {
          const key = data[i][0].toString().trim().toLowerCase();
          const words = data[i][1] ? data[i][1].toString().split(',').map(w => w.trim().toLowerCase()) : [];
          synonyms[key] = words;
        }
      }
    }
    return synonyms;
  } catch(e) {
    return {};
  }
}
