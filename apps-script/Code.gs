/** بک‌اند طرح قطره / ۳۱۳ — Google Apps Script
 *  ۱) یک Google Sheet بسازید → Extensions > Apps Script → این کد را بچسبانید
 *  ۲) تابع setup را یک بار اجرا کنید (ادمین پیش‌فرض: admin / ChangeMe123)
 *  ۳) Deploy > New deployment > Web app > Execute as: Me | Access: Anyone
 *  ۴) آدرس /exec را در index.html داخل API بگذارید
 */
const SALT = 'change-this-salt';
const HEAD = {
  Members: ['id','plan','name','family','phone','pledge','parentId','createdAt','createdBy'],
  Users: ['username','passHash','role','memberId'],
  Devices: ['tokenHash','username','label','createdAt','lastUsed']
};

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(n => {
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    sh.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]);
  });
  sh_('Users').appendRow(['admin', hash_('ChangeMe123'), 'admin', '']);
}

function doPost(e) {
  try { return out_(handle_(JSON.parse(e.postData.contents))); }
  catch (err) { return out_({ ok: false, error: err.message }); }
}

function handle_(r) {
  if (r.action === 'login') return login_(r);
  if (r.action === 'deviceLogin') return deviceLogin_(r);
  const u = auth_(r.token);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    if (r.action === 'registerDevice') {
      const token = Utilities.getUuid() + Utilities.getUuid();
      dev_().appendRow([hash_(token), u.username, String(r.label || '').slice(0, 80), new Date(), new Date()]);
      return { ok: true, deviceToken: token };
    }
    if (r.action === 'revokeDevice') {
      const sh = dev_(), v = sh.getDataRange().getValues(), h = hash_(String(r.deviceToken || ''));
      for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === h && v[i][1] === u.username) sh.deleteRow(i + 1);
      return { ok: true };
    }
    if (r.action === 'list') {
      let rows = rows_('Members');
      if (u.role !== 'admin') rows = rows.filter(m => String(m.parentId) === String(u.memberId));
      return { ok: true, data: rows, role: u.role };
    }
    if (r.action === 'add') {
      const m = r.member || {};
      if (!m.name || !m.family) throw new Error('نام و نام خانوادگی الزامی است');
      if (!/^09\d{9}$/.test(String(m.phone))) throw new Error('شماره موبایل نامعتبر است');
      const parent = u.role === 'admin' ? (m.parentId || '') : u.memberId;
      const id = Utilities.getUuid().slice(0, 8);
      sh_('Members').appendRow([id, m.plan === '313' ? '313' : 'ghatre', m.name, m.family,
        "'" + m.phone, Number(m.pledge) || 0, parent, new Date(), u.username]);
      return { ok: true, id };
    }
    if (r.action === 'addMany') {
      const list = r.members || [], parent = u.role === 'admin' ? (r.parentId || '') : u.memberId;
      if (!list.length) throw new Error('موردی برای ثبت نیست');
      const now = new Date(), out = list.map(m => {
        if (!m.name || !m.family) throw new Error('نام و نام خانوادگی الزامی است');
        if (!/^09\d{9}$/.test(String(m.phone))) throw new Error('شماره موبایل نامعتبر: ' + m.phone);
        return [Utilities.getUuid().slice(0, 8), 'ghatre', m.name, m.family, "'" + m.phone, Number(m.pledge) || 0, parent, now, u.username];
      });
      const sh = sh_('Members');
      sh.getRange(sh.getLastRow() + 1, 1, out.length, 9).setValues(out);
      return { ok: true, count: out.length };
    }
    if (r.action === 'update') {
      const m = r.member || {}, sh = sh_('Members');
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
      const i = ids.findIndex(x => x[0] === r.id);
      if (i < 0) throw new Error('یافت نشد');
      const row = i + 2, old = sh.getRange(row, 1, 1, 9).getValues()[0];
      if (u.role !== 'admin' && String(old[6]) !== String(u.memberId)) throw new Error('دسترسی ندارید');
      if (!m.name || !m.family) throw new Error('نام و نام خانوادگی الزامی است');
      if (!/^09\d{9}$/.test(String(m.phone))) throw new Error('شماره موبایل نامعتبر است');
      const parent = u.role === 'admin' ? (m.parentId || '') : old[6];
      sh.getRange(row, 2, 1, 6).setValues([[m.plan === '313' ? '313' : 'ghatre', m.name, m.family, "'" + m.phone, Number(m.pledge) || 0, parent]]);
      return { ok: true };
    }
    if (r.action === 'remove') {
      if (u.role !== 'admin') throw new Error('دسترسی ندارید');
      const sh = sh_('Members'), ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
      const i = ids.findIndex(x => x[0] === r.id);
      if (i < 0) throw new Error('یافت نشد');
      sh.deleteRow(i + 2); return { ok: true };
    }
    throw new Error('عملیات نامعتبر');
  } finally { lock.releaseLock(); }
}

function login_(r) {
  const u = rows_('Users').find(x => x.username === r.username && x.passHash === hash_(r.password));
  if (!u) throw new Error('نام کاربری یا رمز اشتباه است');
  return session_(u);
}
function session_(u) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put(token, JSON.stringify(u), 21600);
  return { ok: true, token, role: u.role };
}
/** ورود با اثرانگشت: توکن دستگاه (۹۰ روز اعتبار) */
function deviceLogin_(r) {
  const sh = dev_(), v = sh.getDataRange().getValues(), h = hash_(String(r.deviceToken || ''));
  const i = v.findIndex((x, k) => k > 0 && x[0] === h);
  if (i < 0 || Date.now() - new Date(v[i][4]).getTime() > 90 * 864e5) throw new Error('دستگاه شناسایی نشد؛ با رمز وارد شوید');
  const u = rows_('Users').find(x => x.username === v[i][1]);
  if (!u) throw new Error('کاربر یافت نشد');
  sh.getRange(i + 1, 5).setValue(new Date());
  return session_(u);
}
function dev_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Devices');
  if (!sh) { sh = ss.insertSheet('Devices'); sh.getRange(1, 1, 1, 5).setValues([HEAD.Devices]); }
  return sh;
}
function auth_(t) {
  const s = t && CacheService.getScriptCache().get(t);
  if (!s) throw new Error('نشست منقضی شده؛ دوباره وارد شوید');
  return JSON.parse(s);
}
function hash_(p) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p + SALT)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function sh_(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function rows_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i] instanceof Date ? r[i].toISOString() : r[i]])));
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
