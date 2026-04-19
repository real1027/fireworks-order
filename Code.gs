// ============================================================
//  煙火訂購系統 — Google Apps Script 後端（多 Tab 版）
//  管理密碼存在「設定」Tab B2，可直接在 Sheets 修改
//  部署：貼上存檔 → 部署 → 管理部署作業 → 編輯 → 版本「新版本」→ 部署
// ============================================================

const WRITE_KEY = 'fw2025secret'; // ← 要跟 index.html 裡的 WRITE_KEY 一樣
const SKIP_SHEETS = ['設定', '_meta', 'Sheet1', '工作表1'];

// ── GET：讀取品項 + 管理密碼 ───────────────────────────────
function doGet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const products = [];

    ss.getSheets().forEach(sheet => {
      const catName = sheet.getName();
      if (SKIP_SHEETS.includes(catName)) return;
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return;
      const headers = data[0];
      data.slice(1).forEach(row => {
        if (!row[0] && !row[1]) return;
        const p = {};
        headers.forEach((h, i) => p[String(h).trim()] = row[i]);
        p.cat = catName;
        p.on = (p.on === true || String(p.on).toUpperCase() === 'TRUE');
        if (p.id || p.name) products.push(p);
      });
    });

    const adminPin = getAdminPin(ss);
    return jsonOut({ products, adminPin });
  } catch (e) {
    return jsonOut({ error: e.message });
  }
}

// ── POST：寫入品項 或 更新密碼 ─────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.key !== WRITE_KEY) return jsonOut({ error: 'unauthorized' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSettingsSheet(ss);

    // ── 上傳檔案到 Google Drive ──
    if (payload.action === 'uploadFile') {
      const { fileName, mimeType, data } = payload;
      const folderName = '煙火訂購系統_附件';
      let folder;
      const it = DriveApp.getFoldersByName(folderName);
      folder = it.hasNext() ? it.next() : DriveApp.createFolder(folderName);
      const blob = Utilities.newBlob(Utilities.base64Decode(data), mimeType, fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      const fileId = file.getId();
      // 圖片和影片都用 /preview，透過 iframe 嵌入顯示
      const url = `https://drive.google.com/file/d/${fileId}/preview`;
      return jsonOut({ success: true, url });
    }

    // ── 更新管理密碼 ──
    if (payload.action === 'updatePin') {
      const newPin = String(payload.newPin);
      if (!/^\d{4}$/.test(newPin)) return jsonOut({ error: '密碼必須是4位數字' });
      ss.getSheetByName('設定').getRange('B2').setValue(newPin);
      return jsonOut({ success: true });
    }

    // ── 同步品項 ──
    const products = payload.products;
    const byCat = {};
    products.forEach(p => (byCat[p.cat] = byCat[p.cat] || []).push(p));
    const catNames = Object.keys(byCat);

    const existingSheets = {};
    ss.getSheets().forEach(s => existingSheets[s.getName()] = s);

    catNames.forEach((cat, idx) => {
      let sheet = existingSheets[cat];
      if (!sheet) {
        sheet = ss.insertSheet(cat, idx);
      } else {
        ss.setActiveSheet(sheet);
        ss.moveActiveSheet(idx + 1);
      }
      sheet.clearContents();
      sheet.appendRow(['id', 'name', 'status', 'on', 'productImg', 'effectImg', 'effectVideo']);
      const rows = byCat[cat].map(p => [p.id, p.name, p.status, p.on, p.productImg||'', p.effectImg||'', p.effectVideo||'']);
      if (rows.length > 0) sheet.getRange(2, 1, rows.length, 7).setValues(rows);
      const hr = sheet.getRange(1, 1, 1, 4);
      hr.setFontWeight('bold');
      hr.setBackground('#302b63');
      hr.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, 7);
      if (rows.length > 0)
        sheet.getRange(2, 3, rows.length, 1).setHorizontalAlignment('center');
    });

    // 刪除已不存在的分類 Tab
    ss.getSheets().forEach(sheet => {
      const name = sheet.getName();
      if (!SKIP_SHEETS.includes(name) && !catNames.includes(name))
        if (ss.getSheets().length > 1) ss.deleteSheet(sheet);
    });

    // 把「設定」Tab 移到最後
    const settingsSheet = ss.getSheetByName('設定');
    if (settingsSheet) {
      ss.setActiveSheet(settingsSheet);
      ss.moveActiveSheet(ss.getSheets().length);
    }

    return jsonOut({ success: true, categories: catNames.length, total: products.length });
  } catch (e) {
    return jsonOut({ error: e.message });
  }
}

// ── 工具 ───────────────────────────────────────────────────
function getAdminPin(ss) {
  try {
    const sheet = ss.getSheetByName('設定');
    if (!sheet) return '1234';
    const val = sheet.getRange('B2').getValue();
    return val ? String(val) : '1234';
  } catch (_) { return '1234'; }
}

function ensureSettingsSheet(ss) {
  if (ss.getSheetByName('設定')) return;
  const sheet = ss.insertSheet('設定');
  sheet.getRange('A1').setValue('設定項目');
  sheet.getRange('B1').setValue('值');
  sheet.getRange('A2').setValue('管理密碼');
  sheet.getRange('B2').setValue('1234');
  const hr = sheet.getRange('A1:B1');
  hr.setFontWeight('bold');
  hr.setBackground('#e8eaf6');
  sheet.getRange('A2').setFontWeight('bold');
  sheet.getRange('B2').setFontSize(14).setFontWeight('bold').setFontColor('#302b63');
  sheet.autoResizeColumns(1, 2);
  sheet.setColumnWidth(2, 120);
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
