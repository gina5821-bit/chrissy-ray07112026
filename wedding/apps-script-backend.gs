// ═══════════════════════════════════════════════════
//  婚禮後端 — Google Apps Script
//  GET  ?action=search&name=姓名        → 查詢座位
//  GET  ?action=gallery&type=wedding    → 婚禮照片（新資料夾）
//  GET  ?action=gallery&type=portrait   → 自助婚紗/婚紗（舊資料夾）
//  POST { data, mimeType, name }        → 上傳照片至婚禮資料夾
// ═══════════════════════════════════════════════════

const WEDDING_FOLDER_ID  = '1yGSE1B8nnkSkyPcHEm-aot4unbo4reKX'; // 婚禮當天 + 賓客上傳
const PORTRAIT_FOLDER_ID = '1e14Ethn3l0edZ2wp0R2ZjrQWO3Mt7EDb'; // 自助婚紗 / 婚紗
const SHEET_ID           = '10AGSak3ZFZ_ZfXWgzSpkNCEtwYkBfQ1xyC_m8QfuS-A';

// ── GET ───────────────────────────────────────────
function doGet(e) {

  // 座位查詢
  if (e.parameter.action === 'search') {
    const query = (e.parameter.name || '').trim().toLowerCase();
    if (!query) return respond({ success: false, error: 'no query' });

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const rows  = sheet.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      const [name, english, table, seat, meal] = rows[i];
      if (
        String(name).toLowerCase().includes(query) ||
        String(english).toLowerCase().includes(query)
      ) {
        return respond({ success: true, guest: { name, english, table: String(table), seat: String(seat), meal } });
      }
    }
    return respond({ success: true, guest: null });
  }

  // 祝福留言清單（存在 Drive，不需試算表授權）
  if (e.parameter.action === 'messages') {
    const folder = getMessagesFolder();
    const files  = folder.getFiles();
    const list   = [];
    while (files.hasNext()) {
      const f = files.next();
      try {
        const obj = JSON.parse(f.getBlob().getDataAsString());
        list.push({ name: obj.name, text: obj.text, ts: obj.ts });
      } catch (err) {}
    }
    list.sort((a, b) => b.ts - a.ts); // 最新在前
    return respond({ success: true, messages: list.slice(0, 300) });
  }

  // 即時動態清單（照片 + 文字，存在 Drive）
  if (e.parameter.action === 'moments') {
    const folder = getMomentsFolder();
    const files  = folder.getFiles();
    const list   = [];
    while (files.hasNext()) {
      const f = files.next();
      if (f.getName().slice(-5) === '.json') {
        try {
          const o = JSON.parse(f.getBlob().getDataAsString());
          list.push({ name: o.name, text: o.text, photoId: o.photoId, ts: o.ts });
        } catch (err) {}
      }
    }
    list.sort((a, b) => b.ts - a.ts);
    return respond({ success: true, moments: list.slice(0, 200) });
  }

  // 圖庫照片清單
  if (e.parameter.action === 'gallery') {
    const type     = e.parameter.type || 'wedding';
    const folderId = type === 'portrait' ? PORTRAIT_FOLDER_ID : WEDDING_FOLDER_ID;
    const folder   = DriveApp.getFolderById(folderId);
    const files    = folder.getFiles();
    const photos   = [];

    while (files.hasNext()) {
      const file = files.next();
      const fname = file.getName();
      // 排除即時動態照片（檔名以「動態_」開頭），不混進相簿
      if (file.getMimeType().startsWith('image/') && fname.indexOf('動態_') !== 0) {
        photos.push({ id: file.getId(), name: fname });
      }
    }

    // 依檔名排序
    photos.sort((a, b) => a.name.localeCompare(b.name, 'zh-TW'));
    return respond({ success: true, photos });
  }

  return respond({ success: true, message: 'Wedding API is running ✓' });
}

// ── POST：照片上傳 或 祝福留言 ────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // 祝福留言（存成 Drive 上的 JSON 檔，沿用 Drive 授權）
    if (payload.type === 'message') {
      const name = String(payload.name || '匿名').slice(0, 40);
      const text = String(payload.text || '').slice(0, 300);
      if (!text.trim()) return respond({ success: false, error: 'empty' });
      const obj  = { name: name, text: text, ts: Date.now() };
      const folder = getMessagesFolder();
      folder.createFile(
        Utilities.newBlob(JSON.stringify(obj), 'application/json', 'msg_' + obj.ts + '.json')
      );
      return respond({ success: true });
    }

    // 即時動態（照片 + 文字）
    if (payload.type === 'moment') {
      const name = String(payload.name || '匿名').slice(0, 40);
      const text = String(payload.text || '').slice(0, 200);
      let photoId = '';
      if (payload.data) {
        const bytes = Utilities.base64Decode(payload.data);
        const blob  = Utilities.newBlob(bytes, payload.mimeType || 'image/jpeg', '動態_' + Date.now() + '.jpg');
        // 照片統一存進婚禮資料夾（與其他婚禮照片集中在一起）
        photoId = DriveApp.getFolderById(WEDDING_FOLDER_ID).createFile(blob).getId();
      }
      if (!photoId && !text.trim()) return respond({ success: false, error: 'empty' });
      // 動態索引（名字/文字/照片ID）仍存在即時動態資料夾
      const obj = { name: name, text: text, photoId: photoId, ts: Date.now() };
      getMomentsFolder().createFile(Utilities.newBlob(JSON.stringify(obj), 'application/json', 'moment_' + obj.ts + '.json'));
      return respond({ success: true });
    }

    // 照片上傳至婚禮資料夾
    const folder  = DriveApp.getFolderById(WEDDING_FOLDER_ID);
    const bytes   = Utilities.base64Decode(payload.data);
    const blob    = Utilities.newBlob(bytes, payload.mimeType, payload.name);
    const file    = folder.createFile(blob);
    return respond({ success: true, fileId: file.getId(), name: file.getName() });
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

// 取得（或建立）留言儲存資料夾——不需另外設定 ID
function getMessagesFolder() {
  const NAME = '婚禮祝福留言';
  const it = DriveApp.getFoldersByName(NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(NAME);
}

// 取得（或建立）即時動態資料夾——設為「知道連結者可檢視」，照片才能顯示在網頁上
function getMomentsFolder() {
  const NAME = '婚禮即時動態';
  const it = DriveApp.getFoldersByName(NAME);
  if (it.hasNext()) return it.next();
  const folder = DriveApp.createFolder(NAME);
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (err) {}
  return folder;
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 授權專用：執行此函式以觸發試算表權限授權 ──────
function authorize() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  const rows  = sheet.getDataRange().getValues();
  Logger.log('讀取成功，共 ' + rows.length + ' 列');
  Logger.log('第一筆：' + JSON.stringify(rows[1]));
}
