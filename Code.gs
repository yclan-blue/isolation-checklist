/**
 * 115年隔離醫院緊急應變檢核表 — Google 試算表後端（含 Google 登入白名單驗證）
 * 使用方式：
 * 1. 將下方 GOOGLE_CLIENT_ID 填入 Google Cloud OAuth 2.0 網頁用戶端 ID。
 * 2. 在試算表中「擴充功能 → Apps Script」，貼上本程式碼並儲存。
 * 3. 執行一次 setup()（會建立「授權名單」工作表並自動加入試算表擁有者 Email）。
 * 4. 在「授權名單」工作表新增允許存取的 Google 帳號（Email、權限：編輯／唯讀、狀態：啟用）。
 * 5. 「部署 → 管理部署作業 → 編輯 → 版本：新版本 → 部署」（執行身分：我；誰可以存取：所有人）。
 */
const GOOGLE_CLIENT_ID = '617987658234-tb5q7dqalnbmdji0cu4okdf21isjl53m.apps.googleusercontent.com';

const SHEETS = {
  store: '資料儲存',
  log: '異動紀錄',
  kpi: 'KPI紀錄',
  needs: '需求追蹤',
  items: '查核結果',
  auth: '授權名單',
};
const PARTS = ['field', 'director', 'static'];
const PART_NAME = { field: '現場主管版', director: '院長版', static: '靜態檢核表' };
const RES = { full: '完全符合', part: '部分符合', none: '無規劃', na: '不適用' };

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mk = (name, header) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (header && sh.getLastRow() === 0) {
      sh.appendRow(header);
      sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#dcfce7');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const st = mk(SHEETS.store, ['部分', 'JSON 資料（請勿手動修改）', '更新時間']);
  PARTS.forEach((p, i) => { if (!st.getRange(i + 2, 1).getValue()) st.getRange(i + 2, 1).setValue(p); });
  mk(SHEETS.log, ['時間', '頁面', '醫院', '填報／簽核人', '摘要']);
  mk(SHEETS.kpi, ['時間', '醫院', '警示燈號', '湧入分級', '已驗收可用床', '已收治', '可再收', '佔床率', '急診等床(時)', '呼吸器佔用', 'ICU佔用', '預估滿床(天)', '最晚決策(天)', '最先不足資源', 'B1–B12符合度']);
  mk(SHEETS.needs, ['需求編號', '區域', '類型', '需求內容', '數量', '提出日', '需到位期限', '進度', '院長核定措施', '經費', '執行人', '完成期限', '到位與驗收結果']);
  mk(SHEETS.items, ['項目', '主題', '結果', '佐證', '缺口／說明', '更新時間']);
  const au = mk(SHEETS.auth, ['Google Email（小寫）', '姓名／職稱', '權限（編輯／唯讀）', '狀態（啟用／停用）', '備註']);
  if (au.getLastRow() === 1) {
    const owner = (Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
    if (owner) au.appendRow([owner, '管理員（試算表擁有者）', '編輯', '啟用', 'setup() 自動加入']);
  }
  const def = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

function authErr_(code, msg) {
  const e = new Error(msg);
  e.code = code;
  return e;
}

function verifyUser_(token) {
  if (!GOOGLE_CLIENT_ID) throw authErr_('CONFIG_ERROR', '後端尚未設定 GOOGLE_CLIENT_ID');
  if (!token) throw authErr_('UNAUTHORIZED', '請先使用 Google 帳號登入');

  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token));
  const cacheKey = 'gtok_' + Utilities.base64EncodeWebSafe(digest);
  const cache = CacheService.getScriptCache();
  let info = null;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { info = JSON.parse(cached); } catch (e) {}
  }
  if (!info) {
    const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(String(token)), { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) throw authErr_('TOKEN_EXPIRED', 'Google 登入憑證無效或已過期，請重新登入');
    const j = JSON.parse(resp.getContentText());
    if (j.aud !== GOOGLE_CLIENT_ID) throw authErr_('UNAUTHORIZED', 'Google 登入來源（Client ID）不符');
    if (j.email_verified !== 'true' && j.email_verified !== true) throw authErr_('UNAUTHORIZED', 'Google 帳號 Email 尚未驗證');
    const expMs = Number(j.exp || 0) * 1000;
    if (!expMs || expMs <= Date.now()) throw authErr_('TOKEN_EXPIRED', 'Google 登入已過期，請重新登入');
    info = { email: String(j.email || '').trim().toLowerCase(), name: String(j.name || ''), expMs: expMs };
    const ttl = Math.min(300, Math.max(10, Math.floor((expMs - Date.now()) / 1000) - 5));
    cache.put(cacheKey, JSON.stringify(info), ttl);
  } else if (info.expMs <= Date.now()) {
    cache.remove(cacheKey);
    throw authErr_('TOKEN_EXPIRED', 'Google 登入已過期，請重新登入');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.auth);
  if (!sh) throw authErr_('CONFIG_ERROR', '尚未建立「授權名單」工作表，請執行 setup()');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw authErr_('FORBIDDEN', '「授權名單」尚無任何帳號，請聯繫管理員加入：' + info.email);
  const rows = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  for (let i = 0; i < rows.length; i++) {
    const em = String(rows[i][0] || '').trim().toLowerCase();
    const status = String(rows[i][3] || '').trim();
    if (em && em === info.email) {
      if (status === '停用') throw authErr_('FORBIDDEN', '此 Google 帳號（' + info.email + '）已被停用，請聯繫管理員');
      const role = String(rows[i][2] || '').trim() === '唯讀' ? 'viewer' : 'editor';
      return { email: info.email, name: String(rows[i][1] || info.name || info.email).trim(), role: role };
    }
  }
  throw authErr_('FORBIDDEN', '此 Google 帳號（' + info.email + '）尚未獲得授權，請聯繫管理員加入試算表「授權名單」');
}

function readAll_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.store);
  if (!sh) throw new Error('尚未執行 setup()');
  const out = {};
  PARTS.forEach((p, i) => {
    const v = sh.getRange(i + 2, 2).getValue();
    if (v) { try { out[p] = JSON.parse(v); } catch (e) {} }
  });
  return out;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const token = e && e.parameter && e.parameter.token;
    const user = verifyUser_(token);
    return json_({ ok: true, user: user, data: readAll_(), serverTime: Date.now() });
  } catch (err) {
    return json_({ ok: false, code: err.code || 'ERROR', error: String(err.message || err) });
  }
}

function doPost(e) {
  let body, user;
  try {
    body = JSON.parse(e.postData.contents);
    user = verifyUser_(body.token);
    if (body.action === 'get') {
      return json_({ ok: true, user: user, data: readAll_(), serverTime: Date.now() });
    }
    if (user.role !== 'editor') {
      throw authErr_('FORBIDDEN', '此帳號（' + user.email + '）僅有「唯讀」權限，無法修改雲端資料');
    }
  } catch (err) {
    return json_({ ok: false, code: err.code || 'ERROR', error: String(err.message || err) });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const st = ss.getSheetByName(SHEETS.store);
    if (!st) throw new Error('尚未執行 setup()');
    const now = new Date();
    const saved = [];
    PARTS.forEach((p, i) => {
      const part = body.parts && body.parts[p];
      if (!part) return;
      const cur = st.getRange(i + 2, 2).getValue();
      let curTs = 0; try { curTs = cur ? (JSON.parse(cur)._ts || 0) : 0; } catch (x) {}
      if ((part._ts || 0) < curTs) return; // 雲端較新，不覆蓋
      const s = JSON.stringify(part);
      if (s.length > 49000) throw new Error(p + ' 資料過大（超過試算表儲存格上限）');
      st.getRange(i + 2, 1, 1, 3).setValues([[p, s, now]]);
      saved.push(p);
    });
    const all = readAll_();
    const f = all.field || {}, meta = f.meta || {};
    const actor = body.by ? (body.by + ' [' + user.email + ']') : user.email;
    if (saved.length) {
      ss.getSheetByName(SHEETS.log).appendRow([now, saved.map(p => PART_NAME[p]).join('、'), meta.hospital || '', actor, body.summary || '']);
    }
    if (body.kpi && saved.indexOf('field') > -1) {
      const k = body.kpi;
      ss.getSheetByName(SHEETS.kpi).appendRow([now, meta.hospital || ''].concat(['level', 'surge', 'accAll', 'admAll', 'remAll', 'occAll', 'edWait', 'ventOcc', 'icuOcc', 'daysToFull', 'latestDays', 'first', 'comp'].map(x => k[x] == null ? '' : k[x])));
    }
    if (saved.indexOf('field') > -1 || saved.indexOf('director') > -1) writeNeeds_(ss, all);
    if (saved.indexOf('field') > -1) writeItems_(ss, f, now);
    return json_({ ok: true, user: user, saved: saved, data: all, serverTime: Date.now() });
  } catch (err) {
    return json_({ ok: false, code: err.code || 'ERROR', error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function writeNeeds_(ss, all) {
  const sh = ss.getSheetByName(SHEETS.needs);
  const needs = (all.field && all.field.needs) || [];
  const dec = (all.director && all.director.decisions) || {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 13).clearContent();
  if (!needs.length) return;
  const rows = needs.map(x => { const d = dec[x.id] || {}; return [x.id || '', x.zone || '', x.type || '', x.desc || '', x.qty || '', x.raised || '', x.due || '', x.status || '提出', d.measure || '', d.budget || '', d.owner || '', d.due || '', d.result || '']; });
  sh.getRange(2, 1, rows.length, 13).setValues(rows);
}

function writeItems_(ss, f, now) {
  const sh = ss.getSheetByName(SHEETS.items);
  const items = f.items || {};
  const titles = { B1: '急診接收與交接', B2: '三區實際量能', B3: '第二區擴充整備', B4: '第三區預備支援', B5: '各班人力與長期支持', B6: '跨科及外援能力', B7: '設備與供氧能力', B8: '物資與藥品供應', B9: '分區感控與聯合驗收', B10: '醫療服務與後勤銜接', B11: '量能不足與轉院', B12: '資源到位與異常回報' };
  const rows = Object.keys(titles).map(id => { const v = items[id] || {}; return [id, titles[id], RES[v.r] || '', v.ev || '', v.gap || '', now]; });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 6).clearContent();
  sh.getRange(2, 1, rows.length, 6).setValues(rows);
}
