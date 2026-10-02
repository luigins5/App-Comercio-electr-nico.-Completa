/**
 * Developed by Mohammad Rameez Imdad (Rameez Scripts)
 * WhatsApp: https://whatsapp.rameezscripts.com/ (For Custom Projects)
 * YouTube: https://www.youtube.com/@rameezimdad (Subscribe for more!)
 */

// ============== Configuration ==============
var APP_NAME = 'Tienda de Bebidas';
var USERS_SHEET = 'USERS', LOGS_SHEET = 'LOGS', ROLES_SHEET = 'ROLES', SETTINGS_SHEET = 'SETTINGS', RESETS_SHEET = 'PASSWORD_RESETS',
    ASSETS_FOLDER_NAME = 'ASSETS', RECEIPTS_FOLDER_NAME = 'Shop Receipts (private)';
var CAT_SHEET = 'CATEGORIES', PROD_SHEET = 'PRODUCTS', ADDON_SHEET = 'ADDONS', PM_SHEET = 'PAYMENT_METHODS', ORDER_SHEET = 'ORDERS';
// cols 12-14 = project additions (full name, phone, delivery address)
var USER_HEAD = ['Username','Email','Password','Role','Status','ProfileImage','ThemeMode','CustomColors','CreatedAt','CreatedBy','UpdatedAt','UpdatedBy','FullName','Phone','Address'];
var LOG_HEAD  = ['Timestamp','User','Action','Details'];
var ROLE_HEAD = ['role_key','label','color','sort_order','is_super','hidden_signup','permissions'];
var SET_HEAD  = ['Key','Value'];
// cols: 0=token, 1=username, 2=expires, 3=used, 4=created
var RESET_HEAD = ['token','username','expires_at','used_at','created_at'];

// ============== RBAC Config ==============
// cols: 0=role_key, 1=label, 2=color, 3=sort, 4=is_super, 5=hidden_signup, 6=permissions(JSON)
var ROLE_C = { KEY:0, LABEL:1, COLOR:2, SORT:3, SUPER:4, HIDDEN:5, PERMS:6 };

// pages the matrix governs (account = always-on, permissions = editor-only — not listed here)
var RBAC_PAGES = [
  { key:'dashboard',  label:'Panel de control',    group:'General' },
  { key:'orders',     label:'Pedidos',             group:'Pedidos' },
  { key:'products',   label:'Productos',           group:'Menú' },
  { key:'categories', label:'Categorías',          group:'Menú' },
  { key:'addons',     label:'Complementos',        group:'Menú' },
  { key:'customers',  label:'Clientes',            group:'Ventas' },
  { key:'reports',    label:'Reportes',            group:'Ventas' },
  { key:'paymethods', label:'Métodos de pago',     group:'Configuración' },
  { key:'settings',   label:'Configuración',       group:'Configuración' },
  { key:'users',      label:'Usuarios',            group:'Sistema' },
  { key:'logs',       label:'Registro de actividad', group:'Sistema' },
  { key:'mysettings', label:'Mis ajustes',         group:'Sistema' },
  { key:'about',      label:'Acerca de',           group:'Sistema' }
];

// roles — keys MATCH USERS.Role values; Customer is the self-signup role (Active on signup, storefront only)
var RBAC_ROLE_DEFS = [
  { key:'Admin',    label:'Administrador', color:'#6a1b9a', is_super:1, hidden_signup:1 },
  { key:'Customer', label:'Cliente',       color:'#0074D9', is_super:0, hidden_signup:0 }
];

var RBAC_EDIT_ROLES = ['Admin']; // who may open/edit the matrix

// default grants per role: v=view a=add e=edit d=delete — Admin (owner) gets everything, Customer no admin page
var ROLE_GRANTS = { Customer: {} };

function rbacDefaultPerms_(roleKey) {
  var g = roleKey === 'Admin' ? null : (ROLE_GRANTS[roleKey] || {}), perms = {};
  RBAC_PAGES.forEach(function(p){
    var s = g ? (g[p.key] || '') : 'vaed', has = function(c){ return s.indexOf(c) !== -1 ? 1 : 0; };
    perms[p.key] = { v:has('v'), a:has('a'), e:has('e'), d:has('d') };
  });
  return perms;
}

// ============== Tiny helpers (DRY spine) ==============
var DEFAULT_LOGO = 'https://blogger.googleusercontent.com/img/b/R29vZ2xl/AVvXsEiGXxCe0WNNedmFqSWeF761f7Kshhc-NP5ChRQKz9fr97cO8VaarvD0KlCwqHojJVBWv-RAxfOqMI5rD4H78KnARyOc6QgwL1nRRFWf5xNQ1d9F9HfAoLPPGlTyP0GwNl4n-INMEsWLQ4Y7zJtz5bOdAnc2ePH9-uCRgshlo6BsS6gJEz6fhrxL-5U5O3sX/s160/channels4_profile.jpg';
// USERS cols: 0=name 1=email 2=pwd 3=role 4=status 5=img 6=theme 7=colors 8=createdAt 9=createdBy 10=updatedAt 11=updatedBy 12=fullName 13=phone 14=address
var U = { NAME:0, EMAIL:1, PWD:2, ROLE:3, STATUS:4, IMG:5, THEME:6, COLORS:7, CREATED:8, CREATED_BY:9, UPDATED:10, UPDATED_BY:11, FULL:12, PHONE:13, ADDR:14 };
var EMAIL_RX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/, COLOR_RX = /^#[0-9a-f]{6}$/i, PHONE_RX = /^\+?[0-9 ]{7,20}$/;
var ok_  = function(o){ o = o || {}; o.success = true; return o; };
var err_ = function(m){ return { success:false, message:m }; };
var _m = {};   // per-request memo — every google.script.run call is a fresh V8 context, so it can never go stale across users
var ss_  = function(){ return _m.ss || (_m.ss = SpreadsheetApp.getActiveSpreadsheet()); };
var sh_  = function(n){ var c = _m.sh || (_m.sh = {}); return (n in c) ? c[n] : (c[n] = ss_().getSheetByName(n)); };
var nowIso_ = function(){ return new Date().toISOString(); };
var iso_ = function(v){ if (v && !(typeof v === 'string' && v.indexOf('T') !== -1)) { try { return new Date(v).toISOString(); } catch (e) {} } return v; };
var safeParse_ = function(s, d){ try { return JSON.parse(s); } catch (e) { return d; } };   // a corrupt cell must not nuke a read
// only a css-var map of plain values may reach the raw pre-paint injection in index.html
var themeVarsSafe_ = function(json){ var o = safeParse_(json, null), out = {};
  if (!o || typeof o !== 'object') return '';
  Object.keys(o).forEach(function(k){ var v = String(o[k]); if (/^--[\w-]{1,40}$/.test(k) && /^[#\w(),.% -]{1,60}$/.test(v)) out[k] = v; });
  return JSON.stringify(out); };
var themeIdSafe_ = function(id){ return String(id || '').replace(/[^\w -]/g, '').slice(0, 20); };
var head_ = function(sh, cols){ sh.getRange(1, 1, 1, cols.length).setValues([cols]).setBackground('#001f3f').setFontColor('white').setFontWeight('bold'); return sh; };

// cache: roles + identity index ride along with every session lookup (one getAll per request)
var CACHE_ = CacheService.getScriptCache(), ROLES_KEY = 'rbac_roles_v1', UIX_KEY = 'users_ix_v1';
var CACHE_TTL = 21600, UIX_TTL = 1800, SESSION_TTL = 3600;   // roles 6h · identity 30m safety net · session 1h absolute
var cachePut_ = function(k, v, ttl){ var s = JSON.stringify(v); if (s.length < 95000) CACHE_.put(k, s, ttl); };   // 100 KB/value cap

var userRows_ = function(){ return _m.users || (_m.users = (function(){
  var s = sh_(USERS_SHEET); if (!s) throw 'Hoja de usuarios no encontrada';
  var d = s.getDataRange().getValues(), ix = {};
  for (var i = 1; i < d.length; i++) ix[String(d[i][U.NAME])] = i;   // decorate once: name -> row idx
  return { sh:s, data:d, ix:ix }; })()); };
var uRow_ = function(u, name){ var i = u.ix[String(name)]; return i === undefined ? -1 : i; };
var saveRow_ = function(u, i, r, by){ r[U.UPDATED] = nowIso_(); r[U.UPDATED_BY] = by; while (r.length < USER_HEAD.length) r.push(''); putText_(u.sh, i + 1, [r]); };
// keep the identity cache in step with a write instead of dropping it (inside the lock, so it re-reads the live copy)
var ixPatch_ = function(fn){ _m.users = null; var ix = safeParse_(CACHE_.get(UIX_KEY), null); if (!ix) return; fn(ix); cachePut_(UIX_KEY, (_m.uix = ix), UIX_TTL); };
var bustRoles_ = function(){ _m.roles = null; CACHE_.remove(ROLES_KEY); };
var userOut_ = function(r){ return { Username:r[U.NAME] || '', Email:r[U.EMAIL] || '', Role:r[U.ROLE] || '', Status:r[U.STATUS] || '', CreatedAt:iso_(r[U.CREATED]) || '',
  FullName:String(r[U.FULL] || r[U.NAME] || ''), Phone:String(r[U.PHONE] || ''), Address:String(r[U.ADDR] || '') }; };
// one row shape for every writer — 15 cols, sheet-safe
var userRow_ = function(o){ var ts = o.ts || nowIso_();
  return [o.name, o.email, o.pwd, o.role, o.status, DEFAULT_LOGO, 'light', '', ts, o.by, ts, o.by, o.full || o.name, o.phone || '', o.addr || '']; };
// email is the customer login -> unique across every account (case-insensitive)
var emailTaken_ = function(u, email, self){ email = String(email).trim().toLowerCase();
  return u.data.slice(1).some(function(r){ return String(r[U.NAME]) !== self && String(r[U.EMAIL]).trim().toLowerCase() === email; }); };
// full name / phone / address — shared by add, edit, import, signup, profile
function userExtra_(d, cur) {
  var full = String(d.FullName || '').trim().slice(0, 80), phone = String(d.Phone == null ? '' : d.Phone).trim().slice(0, 20),
      addr = String(d.Address == null ? '' : d.Address).trim().slice(0, 300);
  if (phone && !PHONE_RX.test(phone)) fail_('Teléfono: 7-20 dígitos (se permiten espacios y un + al inicio)');
  return { full:full || (cur ? String(cur[U.FULL] || '') : ''), phone:phone, addr:addr };
}

// read-modify-write under one lock — the read happens INSIDE it, so row indexes can't shift underneath
// re-entrant: a nested call runs inside the outer lock instead of releasing it early
function withLock_(fn) {
  if (_m.locked) return fn();
  var l = LockService.getScriptLock(); l.waitLock(20000);
  try { _m.users = null; _m.jdb = {}; _m.jk = {}; _m.jr = {}; _m.memo = {}; _m.locked = true; return fn(); } finally { _m.locked = false; l.releaseLock(); }
}

// user-facing refusal — api_ turns it into a clean message (no "Error:" prefix)
var fail_ = function(m){ throw { user:m }; };
var tz_  = function(){ return _m.tz || (_m.tz = Session.getScriptTimeZone()); };
// any date-ish -> 'yyyy-MM-dd' in script tz (plain day keys pass through untouched)
var ymd_ = function(v){
  if (!v) return '';
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}(T[\d:.]+)?$/.test(v)) return v.slice(0, 10);   // plain day / local datetime: the typed day
  var d = v instanceof Date ? v : new Date(v);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
};
var escHtml_ = function(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, function(c){ return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]; }); };
// classic tab write: every cell plain text, format set in the SAME write (no 1%->0.01, no lost leading zeros)
var putText_ = function(sh, row, rows){ if (rows.length) sh.getRange(row, 1, rows.length, rows[0].length).setNumberFormat('@').setValues(rows); };

// ============== Session ==============
var AUTH = 'AUTH';
// name -> [role, status]; cached so a gate costs zero sheet reads
function uix_() {
  if (_m.uix) return _m.uix;
  var d = userRows_().data, ix = {};
  for (var i = 1; i < d.length; i++) ix[String(d[i][U.NAME])] = [d[i][U.ROLE], d[i][U.STATUS]];
  cachePut_(UIX_KEY, ix, UIX_TTL);
  return (_m.uix = ix);
}
// token -> {u, role}; inactive / deleted users are cut off on their very next call
function who_(tok) {
  if (_m.me) return _m.me;
  if (!tok || typeof tok !== 'string') throw AUTH;
  var hit = CACHE_.getAll(['s_' + tok, UIX_KEY, ROLES_KEY]), name = hit['s_' + tok];
  if (!name) throw AUTH;
  if (hit[UIX_KEY]) _m.uix = safeParse_(hit[UIX_KEY], null);
  if (hit[ROLES_KEY]) _m.roles = safeParse_(hit[ROLES_KEY], null);
  var row = uix_()[name];
  if (!row || row[1] !== 'Active') { CACHE_.remove('s_' + tok); throw AUTH; }
  return (_m.me = { u:name, role:row[0] });
}
// optional session for the public endpoints: a live token binds the caller, a dead / missing one = guest (never an AUTH reply)
var whoMaybe_ = function(tok){ try { return tok ? who_(tok) : null; } catch (e) { if (e === AUTH) return null; throw e; } };
// endpoint funnel: session -> gate -> body. page null = any signed-in user, 'rbac' = matrix editors
function api_(tok, page, perm, fn) {
  try {
    var me = who_(tok);
    if (page && !(page === 'rbac' ? canEditRbac_(me.role) : hasPerm_(me.role, page, perm))) return err_('Acceso denegado');
    return fn(me);
  } catch (e) {
    if (e === AUTH) return { success:false, code:AUTH, message:'La sesión ha expirado. Por favor, inicia sesión nuevamente.' };
    return e && e.user ? err_(e.user) : err_('Error: ' + ((e && e.message) || e));
  }
}
// public funnel: no session needed, the same clean error shape
function pub_(fn) {
  try { return fn(); }
  catch (e) { return e && e.user ? err_(e.user) : err_('Algo salió mal — por favor intenta de nuevo'); }
}

// ============== Main Web App Entry Point ==============
var bootTok_ = function(v){ return String(v || '').replace(/[^A-Za-z0-9-]/g, '').slice(0, 64); };
function doGet(e) {
  var t;
  try { t = HtmlService.createTemplateFromFile('Index'); } catch (err) { t = HtmlService.createTemplateFromFile('index'); }
  var p = PropertiesService.getScriptProperties().getProperties(), q = (e && e.parameter) || {};
  t.defaultThemeVars = themeVarsSafe_(p.DEFAULT_THEME_VARS);   // zero-flash first paint (printed raw, so re-checked)
  t.defaultThemeId = themeIdSafe_(p.DEFAULT_THEME_ID);
  // ?reset= / ?track= / ?page= — the app runs in a sandbox iframe, so it can't read the top URL itself
  t.resetToken = bootTok_(q.reset); t.trackToken = bootTok_(q.track); t.startPage = bootTok_(q.page);
  var name = APP_NAME; try { name = settings_().shop_name || APP_NAME; } catch (x) {}
  return t.evaluate()
    .setTitle(name)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// server-side splice — 4 editor files, ONE document in the browser
function include(f) { return HtmlService.createHtmlOutputFromFile(f).getContent(); }

// simple trigger: hand edits in the sheet drop the matching cache (row deletes by hand fall back to UIX_TTL)
function onEdit(e) {
  var n = e && e.range && e.range.getSheet().getName();
  if (n === USERS_SHEET) CACHE_.remove(UIX_KEY);
  if (n === ROLES_SHEET) CACHE_.remove(ROLES_KEY);
  if (n === SETTINGS_SHEET) CACHE_.remove(SET_KEY);
  if ([SETTINGS_SHEET, CAT_SHEET, PROD_SHEET, ADDON_SHEET, PM_SHEET].indexOf(n) !== -1) bustFront_();
}

// ============== Authentication ==============
var LOGIN_MAX = 5, LOGIN_LOCK = 900;   // 5 failures -> 15 min lockout, per login
// username first, else a unique email match (customers sign in with their email)
var findLogin_ = function(u, ident){
  var i = uRow_(u, ident); if (i !== -1) return i;
  var key = String(ident).trim().toLowerCase(), hits = [];
  if (!key) return -1;
  for (var j = 1; j < u.data.length; j++) if (String(u.data[j][U.EMAIL]).trim().toLowerCase() === key) hits.push(j);
  return hits.length === 1 ? hits[0] : -1;
};
// the one session shape login + signup hand back
function sessionOut_(r) {
  var tok = Utilities.getUuid();
  CACHE_.put('s_' + tok, String(r[U.NAME]), SESSION_TTL);
  return ok_({ token:tok, username:r[U.NAME], email:r[U.EMAIL], role:r[U.ROLE], fullName:String(r[U.FULL] || r[U.NAME]), phone:String(r[U.PHONE] || ''),
               address:String(r[U.ADDR] || ''), profileImage:r[U.IMG] || '', themeMode:r[U.THEME] || 'light', customColors:r[U.COLORS] || '',
               permissions:rbacPermsFor_(r[U.ROLE]), canEditRbac:canEditRbac_(r[U.ROLE]),
               has_open:(rbacPermsFor_(r[U.ROLE]).orders || {}).v ? (openDays_().length ? 1 : 0) : undefined });   // where the admin lands: Board while orders are open
}
function authenticateUser(username, password) {
  try {
    var name = String(username || '').trim().slice(0, 100), lk = 'lf_' + name.toLowerCase(), fails = Number(CACHE_.get(lk)) || 0;
    if (fails >= LOGIN_MAX) { addLog_(name, 'Inicio Bloqueado', 'Demasiados intentos fallidos'); return err_('Demasiados intentos fallidos. Intenta de nuevo en 15 minutos.'); }
    var bad = function(msg, why){ CACHE_.put(lk, String(fails + 1), LOGIN_LOCK); addLog_(name, 'Inicio Fallido', why); return err_(msg); };
    var u = userRows_(), i = findLogin_(u, name);
    if (i === -1) return bad('Correo/usuario o contraseña no válidos', 'Cuenta no encontrada');   // mismo mensaje para ambos
    var r = u.data[i];
    if (r[U.STATUS] !== 'Active') return bad('La cuenta está inactiva. Comunícate con la tienda.', 'Cuenta inactiva');
    if (String(password) !== String(r[U.PWD])) return bad('Correo/usuario o contraseña no válidos', 'Contraseña incorrecta');
    CACHE_.remove(lk); uix_();                                                              // calentar caché de identidad
    addLog_(r[U.NAME], 'Inicio Exitoso', 'Sesión iniciada exitosamente');
    return sessionOut_(r);
  } catch (e) { return err_('Error: ' + e); }
}

function logout(tok) { if (tok && typeof tok === 'string') CACHE_.remove('s_' + tok); return ok_(); }

// ============== Signup + Forgot / Reset password (public) ==============
var RESET_TTL = 3600000, PUB_MAX = 5, PUB_WINDOW = 3600;   // enlace de restablecimiento 1h · 5 solicitudes por hora por correo
var TOKEN_RX = /^[0-9a-f-]{36}$/i;
// cheap per-key throttle for the public endpoints (signup / forgot / track)
var throttled_ = function(k, max, win){ var n = Number(CACHE_.get(k)) || 0; if (n >= (max || PUB_MAX)) return true; CACHE_.put(k, String(n + 1), win || PUB_WINDOW); return false; };
var signupRole_ = function(){ return (readRoles_().filter(function(r){ return !r.is_super && !r.hidden_signup; })[0] || {}).key || ''; };
function resetsSheet_() { return sh_(RESETS_SHEET) || (_m.sh[RESETS_SHEET] = head_(ss_().insertSheet(RESETS_SHEET), RESET_HEAD)); }

// a shopper signs up and is signed in at once (Username = the email, Active) — the storefront needs no approval step
function signup(d) {
  try {
    d = d || {};
    var email = String(d.email || '').trim().toLowerCase(), pwd = String(d.password || ''), full = String(d.name || '').trim(), phone = String(d.phone || '').trim();
    if (full.length < 2) return err_('Ingresa tu nombre');
    if (!EMAIL_RX.test(email) || email.length > 100) return err_('Ingresa un correo electrónico válido');
    if (!PHONE_RX.test(phone)) return err_('Teléfono: 7-20 dígitos (se permiten espacios y un + al inicio)');
    if (pwd.length < 8) return err_('La contraseña debe tener al menos 8 caracteres');
    var role = signupRole_();
    if (!role) return err_('El registro está cerrado — por favor realiza tu pedido como invitado');
    if (throttled_('su_' + email)) return err_('Demasiados intentos. Intenta más tarde.');
    var out = withLock_(function(){
      var u = userRows_();
      if (uRow_(u, email) !== -1 || emailTaken_(u, email, '')) return err_('Ese correo ya tiene una cuenta registrada — inicia sesión');
      var row = userRow_({ name:email, email:email, pwd:pwd, role:role, status:'Active', by:'signup', full:full.slice(0, 80), phone:phone, addr:String(d.address || '').trim().slice(0, 300) });
      putText_(u.sh, u.data.length + 1, [row]);
      ixPatch_(function(ix){ ix[email] = [role, 'Active']; });
      addLog_(email, 'REGISTRO', 'Cuenta de cliente creada');
      var c = d.claim || {}, o = c.order_no ? findOrder_(c.order_no) : null;                // vincular pedido de invitado
      if (o && o.track_token === String(c.track_token || '') && !o.customer_username)
        JDB.update(ORDER_SHEET, o.id, { customer_username:email }, ymd_(o.placed_at));
      return sessionOut_(row);
    });
    if (out && out.success) { var st = settings_(), who = { name:full, email:email, phone:phone };
      mailSend_('customer.welcome', null, st, who); mailSend_('shop.signup', null, st, who); }
    return out;
  } catch (e) { return err_('Error al registrarse — por favor intenta de nuevo'); }
}

// always the same answer — never reveals whether an account exists
function forgotPassword(ident) {
  var done = ok_({ message:'Si la cuenta existe, se ha enviado un enlace de restablecimiento a su correo.' });
  try {
    var key = String(ident || '').trim().toLowerCase().slice(0, 100);
    if (!key || throttled_('fp_' + key)) return done;
    var u = userRows_(), i = findLogin_(u, key), r = i === -1 ? null : u.data[i];
    if (!r || r[U.STATUS] !== 'Active' || !EMAIL_RX.test(String(r[U.EMAIL]))) return done;
    var tok = Utilities.getUuid(), sh = resetsSheet_(), shop = settings_().shop_name || APP_NAME;
    withLock_(function(){ putText_(sh, sh.getLastRow() + 1, [[tok, String(r[U.NAME]), new Date(Date.now() + RESET_TTL).toISOString(), '', nowIso_()]]); });
    var url = ScriptApp.getService().getUrl() + '?reset=' + tok, nm = escHtml_(r[U.FULL] || r[U.NAME]);
    MailApp.sendEmail({ to:String(r[U.EMAIL]), subject:shop + ' — restablecimiento de contraseña', htmlBody:
      '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:20px">' +
      '<h2 style="color:#001f3f">Restablecer contraseña</h2><p>Hola ' + nm + ',</p><p>Haz clic en el botón de abajo para establecer una nueva contraseña:</p>' +
      '<p><a href="' + url + '" style="display:inline-block;padding:12px 24px;background:#001f3f;color:#fff;text-decoration:none;border-radius:4px;font-weight:600">Restablecer contraseña</a></p>' +
      '<p style="color:#666;font-size:13px">Este enlace funciona una sola vez y vence en 1 hora. Si no solicitaste este cambio, puedes ignorar este correo.</p></div>' });
    addLog_(r[U.NAME], 'SOLICITUD_RESTABLECER_PASS', 'Enlace de restablecimiento enviado por correo');
  } catch (e) { Logger.log('forgotPassword: ' + e); }
  return done;
}

// single-use token -> new password; every other open link for that user dies with it
function resetPassword(token, pwd) {
  try {
    token = String(token || ''); pwd = String(pwd || '');
    if (!TOKEN_RX.test(token)) return err_('Este enlace de restablecimiento no es válido');
    if (pwd.length < 8) return err_('La contraseña debe tener al menos 8 caracteres');
    return withLock_(function(){
      var sh = resetsSheet_(), rows = sh.getDataRange().getValues(), i = -1;
      for (var j = 1; j < rows.length; j++) if (rows[j][0] === token) { i = j; break; }
      if (i === -1) return err_('Este enlace de restablecimiento no es válido');
      if (rows[i][3]) return err_('Este enlace ya fue utilizado — solicita uno nuevo');
      if (new Date(rows[i][2]).getTime() < Date.now()) return err_('Este enlace ha expirado — solicita uno nuevo');
      var name = String(rows[i][1]), u = userRows_(), k = uRow_(u, name);
      if (k === -1) return err_('Cuenta no encontrada');
      var r = u.data[k]; r[U.PWD] = pwd; r[U.UPDATED] = nowIso_(); r[U.UPDATED_BY] = 'reset';
      while (r.length < USER_HEAD.length) r.push('');
      putText_(u.sh, k + 1, [r]);
      var now = nowIso_(), used = rows.slice(1).map(function(x){ return [String(x[1]) === name && !x[3] ? now : x[3]]; });
      sh.getRange(2, 4, used.length, 1).setNumberFormat('@').setValues(used);
      CACHE_.remove('lf_' + name.toLowerCase());                                          // limpiar bloqueo de login
      addLog_(name, 'PASS_RESTABLECIDA', 'Contraseña cambiada desde enlace de restablecimiento');
      return ok_({ message:'Contraseña cambiada — ya puedes iniciar sesión.' });
    });
  } catch (e) { return err_('Error al restablecer — por favor intenta de nuevo'); }
}

// ============== Users ==============
// role must exist; only an owner (super role) may hand out or touch an owner account
var isSuper_ = function(role){ return !!(roleByKey_(role) || {}).is_super; };
function roleBad_(key, me) {
  if (!roleByKey_(key)) return 'Rol desconocido: ' + key;
  return isSuper_(key) && !isSuper_(me.role) ? 'Solo un propietario puede asignar el rol ' + key : '';
}
var guarded_ = function(r, me){ return isSuper_(r[U.ROLE]) && !isSuper_(me.role); };   // non-owner vs owner account

function getAllUsers(tok) {
  return api_(tok, 'users', 'v', function(){ return ok_({ data: userRows_().data.slice(1).map(userOut_) }); });
}

function addUser(tok, d) {
  return api_(tok, 'users', 'a', function(me){
    d = d || {};
    var name = String(d.Username || '').trim(), role = d.Role || 'Customer', email = String(d.Email || '').trim();
    if (!name || !d.Password) return err_('El usuario y la contraseña son obligatorios');
    if (!EMAIL_RX.test(email)) return err_('Ingresa un correo electrónico válido');
    var bad = roleBad_(role, me); if (bad) return err_(bad);
    return withLock_(function(){
      var u = userRows_();
      if (uRow_(u, name) !== -1) return err_('El nombre de usuario ya existe');
      if (emailTaken_(u, email, name)) return err_('Ese correo ya está registrado en otra cuenta');
      var x = userExtra_(d, null), row = userRow_({ name:name, email:email, pwd:String(d.Password), role:role,
        status:d.Status === 'Inactive' ? 'Inactive' : 'Active', by:me.u, full:x.full, phone:x.phone, addr:x.addr });
      putText_(u.sh, u.data.length + 1, [row]);
      ixPatch_(function(ix){ ix[name] = [row[U.ROLE], row[U.STATUS]]; }); addLog_(me.u, 'Usuario Agregado', 'Usuario agregado: ' + name);
      return ok_({ message:'¡Usuario agregado exitosamente!', user:userOut_(row) });
    });
  });
}

// skip-and-collect: bad rows -> errors[], good rows still land; ONE setValues, ONE log
function bulkImportUsers(tok, rows) {
  return api_(tok, 'users', 'a', function(me){
    if (!rows || !rows.length) return err_('No hay filas para importar');
    return withLock_(function(){
      var u = userRows_(), existing = {}, mails = {}, ts = nowIso_(), out = [], errors = [];
      Object.keys(u.ix).forEach(function(k){ existing[k.trim().toLowerCase()] = 1; });
      u.data.slice(1).forEach(function(r){ mails[String(r[U.EMAIL]).trim().toLowerCase()] = 1; });
      rows.forEach(function(r, i) {
        var name = String(r.Username || '').trim(), key = name.toLowerCase(), mail = String(r.Email || '').trim(), role = roleByKey_(r.Role) ? r.Role : 'Customer', n = 'Fila ' + (i + 1) + ': ';
        if (!name)                                     return errors.push(n + 'falta el nombre de usuario');
        if (!r.Password || !EMAIL_RX.test(mail))       return errors.push(n + 'falta la contraseña o el correo no es válido');
        if (existing[key])                             return errors.push(n + 'nombre de usuario duplicado: ' + name);
        if (mails[mail.toLowerCase()])                 return errors.push(n + 'correo ya registrado: ' + mail);
        if (roleBad_(role, me))                        return errors.push(n + roleBad_(role, me));
        var x; try { x = userExtra_(r, null); } catch (e) { return errors.push(n + ((e && e.user) || e)); }
        existing[key] = 1; mails[mail.toLowerCase()] = 1;                              // dups inside the same csv
        out.push(userRow_({ name:name, email:mail, pwd:String(r.Password), role:role, status:r.Status === 'Inactive' ? 'Inactive' : 'Active',
                            by:me.u, ts:ts, full:x.full, phone:x.phone, addr:x.addr }));
      });
      putText_(u.sh, u.data.length + 1, out);
      ixPatch_(function(ix){ out.forEach(function(r){ ix[r[U.NAME]] = [r[U.ROLE], r[U.STATUS]]; }); }); addLog_(me.u, 'Importación Masiva', 'Usuarios: ' + out.length + ' importados, ' + errors.length + ' omitidos');
      return ok_({ count:out.length, errors:errors });
    });
  });
}

// bulk status flip — one read, two column writes; skips self + owner accounts a non-owner can't touch
function bulkSetStatus(tok, usernames, status) {
  return api_(tok, 'users', 'e', function(me){
    if (!usernames || !usernames.length) return err_('No hay usuarios seleccionados');
    if (status !== 'Active' && status !== 'Inactive') return err_('Estado no válido');
    return withLock_(function(){
      var u = userRows_(), d = u.data, want = {}, ts = nowIso_(), changed = [], skipped = 0;
      if (d.length < 2) return ok_({ count:0 });
      usernames.forEach(function(x){ want[String(x)] = 1; });
      d.slice(1).forEach(function(r){
        if (!want[String(r[U.NAME])] || r[U.STATUS] === status) return;
        if (String(r[U.NAME]) === me.u || guarded_(r, me)) return skipped++;
        r[U.STATUS] = status; r[U.UPDATED] = ts; r[U.UPDATED_BY] = me.u; changed.push(String(r[U.NAME]));
      });
      if (changed.length) {
        u.sh.getRange(2, U.STATUS + 1, d.length - 1, 1).setNumberFormat('@').setValues(d.slice(1).map(function(r){ return [r[U.STATUS]]; }));
        u.sh.getRange(2, U.UPDATED + 1, d.length - 1, 2).setNumberFormat('@').setValues(d.slice(1).map(function(r){ return [r[U.UPDATED], r[U.UPDATED_BY]]; }));
        ixPatch_(function(ix){ changed.forEach(function(x){ if (ix[x]) ix[x][1] = status; }); });
      }
      addLog_(me.u, 'Estado Masivo', status + ': ' + changed.length + ' usuario(s)' + (skipped ? ', ' + skipped + ' omitidos' : ''));
      return ok_({ count:changed.length, skipped:skipped, changed:changed });
    });
  });
}

// bulk delete — compact survivors up, then ONE block delete; never the caller, never an owner for a non-owner
function bulkDeleteUsers(tok, usernames) {
  return api_(tok, 'users', 'd', function(me){
    if (!usernames || !usernames.length) return err_('No hay usuarios seleccionados');
    return withLock_(function(){
      var u = userRows_(), want = {}, keep = [], gone = [], skipped = 0;
      usernames.forEach(function(x){ want[String(x)] = 1; });
      u.data.slice(1).forEach(function(r){
        var nm = String(r[U.NAME]);
        if (want[nm] && nm !== me.u && !guarded_(r, me)) return gone.push(nm);
        if (want[nm]) skipped++;
        keep.push(r);
      });
      if (gone.length) {
        if (keep.length) putText_(u.sh, 2, keep.map(function(r){ while (r.length < USER_HEAD.length) r.push(''); return r; }));
        u.sh.deleteRows(keep.length + 2, gone.length);
        ixPatch_(function(ix){ gone.forEach(function(x){ delete ix[x]; }); });
      }
      addLog_(me.u, 'Eliminación Masiva', 'Usuarios eliminados: ' + gone.length + (skipped ? ' (' + skipped + ' omitidos)' : ''));
      return ok_({ count:gone.length, skipped:skipped, deleted:gone });
    });
  });
}

function updateUser(tok, username, d) {
  return api_(tok, 'users', 'e', function(me){
    d = d || {};
    var bad = roleBad_(d.Role, me); if (bad) return err_(bad);
    if (d.Status !== 'Active' && d.Status !== 'Inactive') return err_('Estado no válido');
    var email = String(d.Email || '').trim();
    if (!EMAIL_RX.test(email)) return err_('Ingresa un correo electrónico válido');
    if (username === me.u && (d.Role !== me.role || d.Status !== 'Active')) return err_('No puedes cambiar tu propio rol o estado');
    return withLock_(function(){
      var u = userRows_(), i = uRow_(u, username);
      if (i === -1) return err_('Usuario no encontrado');
      var r = u.data[i];
      if (guarded_(r, me)) return err_('Solo un propietario puede editar una cuenta de propietario');
      if (emailTaken_(u, email, username)) return err_('Ese correo ya está registrado en otra cuenta');
      var x = userExtra_(d, r);
      r[U.EMAIL] = email; r[U.ROLE] = d.Role; r[U.STATUS] = d.Status;
      r[U.FULL] = x.full || username; r[U.PHONE] = x.phone; r[U.ADDR] = x.addr;
      if (d.Password && String(d.Password).trim()) r[U.PWD] = String(d.Password);
      saveRow_(u, i, r, me.u);
      ixPatch_(function(ix){ ix[username] = [r[U.ROLE], r[U.STATUS]]; }); addLog_(me.u, 'Usuario Actualizado', 'Usuario actualizado: ' + username);
      return ok_({ message:'¡Usuario actualizado exitosamente!', user:userOut_(r) });
    });
  });
}

function deleteUser(tok, username) {
  return api_(tok, 'users', 'd', function(me){
    if (username === me.u) return err_('No puedes eliminar tu propia cuenta');
    return withLock_(function(){
      var u = userRows_(), i = uRow_(u, username);
      if (i === -1) return err_('Usuario no encontrado');
      if (guarded_(u.data[i], me)) return err_('Solo un propietario puede eliminar una cuenta de propietario');
      u.sh.deleteRow(i + 1);
      ixPatch_(function(ix){ delete ix[username]; }); addLog_(me.u, 'Usuario Eliminado', 'Usuario eliminado: ' + username);
      return ok_({ message:'¡Usuario eliminado exitosamente!' });
    });
  });
}

// own login details — CurrentPassword is re-checked on top of the session
function updateMyAccount(tok, f) {
  return api_(tok, null, null, function(me){
    f = f || {};
    var email = String(f.Email || '').trim();
    if (!EMAIL_RX.test(email)) return err_('Ingresa un correo electrónico válido');
    if (f.NewPassword && String(f.NewPassword).length < 8) return err_('La nueva contraseña debe tener al menos 8 caracteres');
    return withLock_(function(){
      var u = userRows_(), i = uRow_(u, me.u), r = u.data[i];
      if (i === -1) throw AUTH;                                         // deleted since the identity cache filled
      if (String(f.CurrentPassword) !== String(r[U.PWD])) return err_('La contraseña actual es incorrecta');
      if (emailTaken_(u, email, me.u)) return err_('Ese correo ya está registrado en otra cuenta');
      r[U.EMAIL] = email;
      if (f.NewPassword && String(f.NewPassword).trim()) r[U.PWD] = String(f.NewPassword);
      saveRow_(u, i, r, me.u);
      addLog_(me.u, 'Perfil Actualizado', 'Actualizó sus propios datos de acceso');
      return ok_({ message:'¡Cuenta actualizada exitosamente!' });
    });
  });
}

// customer profile (name, phone, default delivery address) — pre-fills checkout
function updateMyProfile(tok, f) {
  return api_(tok, null, null, function(me){
    f = f || {};
    if (String(f.FullName || '').trim().length < 2) return err_('Ingresa tu nombre');
    if (!PHONE_RX.test(String(f.Phone || '').trim())) return err_('Teléfono: 7-20 dígitos (se permiten espacios y un + al inicio)');
    return withLock_(function(){
      var u = userRows_(), i = uRow_(u, me.u), r = u.data[i];
      if (i === -1) throw AUTH;
      var x = userExtra_(f, r);
      r[U.FULL] = x.full; r[U.PHONE] = x.phone; r[U.ADDR] = x.addr;
      saveRow_(u, i, r, me.u);
      addLog_(me.u, 'Perfil Actualizado', 'Nombre / teléfono / dirección actualizados');
      return ok_({ message:'Perfil guardado', profile:{ fullName:x.full, phone:x.phone, address:x.addr } });
    });
  });
}

// ============== Profile image + settings ==============
function getAssetsFolder_() {
  var it = DriveApp.getFoldersByName(ASSETS_FOLDER_NAME);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(ASSETS_FOLDER_NAME).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
}

// the one settings shape every call hands back — callers apply it directly instead of re-reading
var settingsOf_ = function(r){ return { profileImage:r[U.IMG] || '', themeMode:r[U.THEME] || 'light', customColors:r[U.COLORS] || '' }; };
// patch own row under the lock, return fresh settings
var patchMe_ = function(me, fn){ return withLock_(function(){
  var u = userRows_(), i = uRow_(u, me.u), r = u.data[i];
  if (i === -1) throw AUTH;
  fn(r); saveRow_(u, i, r, me.u); return settingsOf_(r); }); };

function uploadProfileImage(tok, base64Data, filename) {
  return api_(tok, null, null, function(me){
    var img = saveImage_(base64Data, filename, me.u, getAssetsFolder_()), file = img.file, url = img.url;
    var st = patchMe_(me, function(r){ r[U.IMG] = url; });
    addLog_(me.u, 'Foto de Perfil Subida', 'Foto de perfil subida: ' + file.getName());
    return ok_({ fileId:file.getId(), fileUrl:url, fileName:file.getName(), settings:st });
  });
}

function updateUserSettings(tok, s) {
  return api_(tok, null, null, function(me){
    s = s || {};
    var st = patchMe_(me, function(r){
      if (s.profileImage !== undefined) r[U.IMG] = s.profileImage;
      if (s.themeMode === 'light' || s.themeMode === 'dark') r[U.THEME] = s.themeMode;
      if (s.customColors !== undefined) r[U.COLORS] = String(s.customColors).slice(0, 5000);
    });
    addLog_(me.u, 'Ajustes Actualizados', 'Ajustes de usuario actualizados');
    return ok_({ message:'¡Ajustes actualizados exitosamente!', settings:st });
  });
}

// owner-only: the theme every visitor gets on first load (injected by doGet)
function setDefaultTheme(tok, themeId, varsJson) {
  return api_(tok, 'rbac', null, function(me){
    var id = themeIdSafe_(themeId);
    PropertiesService.getScriptProperties().setProperties({ DEFAULT_THEME_ID:id, DEFAULT_THEME_VARS:themeVarsSafe_(varsJson) });
    addLog_(me.u, 'Tema Predeterminado Fijado', 'Tema predeterminado → ' + id);
    return ok_({ message:'¡Tema predeterminado guardado para todos los usuarios!' });
  });
}

function addLog_(user, action, details) {
  try { var s = sh_(LOGS_SHEET); if (s) s.appendRow([nowIso_(), String(user || '').slice(0, 80), action, details]); }
  catch (e) { Logger.log('Error adding log: ' + e); }
}

// ============== RBAC Backend ==============
// create + seed Roles sheet if missing (idempotent / self-healing)
function ensureRbac_() {
  var sh = sh_(ROLES_SHEET);
  if (sh) return sh;
  sh = _m.sh[ROLES_SHEET] = head_(ss_().insertSheet(ROLES_SHEET), ROLE_HEAD);
  sh.getRange(2, 1, RBAC_ROLE_DEFS.length, ROLE_HEAD.length).setValues(RBAC_ROLE_DEFS.map(function(r, i){
    return [r.key, r.label, r.color, i, r.is_super, r.hidden_signup, JSON.stringify(rbacDefaultPerms_(r.key))]; }));
  return sh;
}

// read-through cache — roles drive every permission check but change rarely
function readRoles_() {
  if (_m.roles) return _m.roles;
  var hit = safeParse_(CACHE_.get(ROLES_KEY), null);
  if (hit) return (_m.roles = hit);
  var out = ensureRbac_().getDataRange().getValues().slice(1).filter(function(r){ return r[ROLE_C.KEY]; }).map(function(r){
    return { key:r[ROLE_C.KEY], label:r[ROLE_C.LABEL], color:r[ROLE_C.COLOR], sort:Number(r[ROLE_C.SORT]) || 0,
             is_super:Number(r[ROLE_C.SUPER]) ? 1 : 0, hidden_signup:Number(r[ROLE_C.HIDDEN]) ? 1 : 0, perms:safeParse_(r[ROLE_C.PERMS], {}) };
  }).sort(function(a, b){ return a.sort - b.sort; });
  cachePut_(ROLES_KEY, out, CACHE_TTL);
  return (_m.roles = out);
}
function roleByKey_(key){ return readRoles_().filter(function(r){ return r.key === key; })[0] || null; }
// role-row lookup inside the sheet — keys unique case-insensitively (addRole's dup rule)
var roleRowIx_ = function(rows, key){ key = String(key).toLowerCase();
  for (var i = 1; i < rows.length; i++) if (String(rows[i][ROLE_C.KEY]).toLowerCase() === key) return i;
  return -1; };
function canEditRbac_(role){ return RBAC_EDIT_ROLES.indexOf(role) !== -1; }
function rbacPermsFor_(role){ var r = roleByKey_(role); return r ? r.perms : rbacDefaultPerms_(role); }
function hasPerm_(role, page, perm){ var r = roleByKey_(role); return !!(r && r.perms && r.perms[page] && r.perms[page][perm || 'v']); }
var rolesOut_ = function(){ return readRoles_().map(function(r){ return { key:r.key, label:r.label, color:r.color, is_super:r.is_super, hidden_signup:r.hidden_signup }; }); };

// caller's own perms — refresh menus on restored sessions (AUTH here = session gone -> client logs out)
function getMyPermissions(tok) {
  return api_(tok, null, null, function(me){ return ok_({ perms:rbacPermsFor_(me.role), canEdit:canEditRbac_(me.role) }); });
}

// full matrix for the editor page; roles double as the picker list
function getRbacMatrix(tok) {
  return api_(tok, 'rbac', null, function(){
    var perms = {};
    readRoles_().forEach(function(r){ perms[r.key] = r.perms; });
    return ok_({ pages:RBAC_PAGES, roles:rolesOut_(), perms:perms });
  });
}

// toggle one cell (v/a/e/d) — implied-view logic + owner-row lock
function toggleRbac(tok, roleKey, pageKey, perm, value) {
  return api_(tok, 'rbac', null, function(me){
    if (['v','a','e','d'].indexOf(perm) === -1) return err_('Permiso no válido');
    if (!RBAC_PAGES.some(function(p){ return p.key === pageKey; })) return err_('Página no válida');
    return withLock_(function(){
      var sh = ensureRbac_(), data = sh.getDataRange().getValues(), i = roleRowIx_(data, roleKey);
      if (i === -1) return err_('Rol no encontrado');
      if (Number(data[i][ROLE_C.SUPER]) === 1) return err_('Los permisos de Administrador (propietario) están bloqueados');
      var p = safeParse_(data[i][ROLE_C.PERMS], {}) || {}, c = p[pageKey] || (p[pageKey] = { v:0, a:0, e:0, d:0 }), on = value ? 1 : 0;
      c[perm] = on;
      if (perm === 'v' && !on) c.a = c.e = c.d = 0;                     // no view -> nothing else
      if (perm !== 'v' && on) c.v = 1;                                  // any grant implies view
      sh.getRange(i + 1, ROLE_C.PERMS + 1).setValue(JSON.stringify(p));
      bustRoles_(); addLog_(me.u, 'Permisos Actualizados', roleKey + ' · ' + pageKey + ' · ' + perm + '=' + on);
      return ok_({ message:'Guardado' });
    });
  });
}

// role list for pickers (user modal, filters) — anyone who can see users
function getRoleOptions(tok) {
  return api_(tok, 'users', 'v', function(){ return ok_({ data:rolesOut_() }); });
}

// add a role — key is what Users.Role stores, so it must be unique and stable
function addRole(tok, role) {
  return api_(tok, 'rbac', null, function(me){
    role = role || {};
    var key = String(role.key || '').trim();
    if (!/^[A-Za-z0-9 _-]{2,30}$/.test(key)) return err_('Clave de rol: 2-30 caracteres — solo letras, números, espacios, _ o -');
    return withLock_(function(){
      var sh = ensureRbac_(), data = sh.getDataRange().getValues();
      if (roleRowIx_(data, key) !== -1) return err_('Ese rol ya existe');
      var src = role.copyFrom ? roleByKey_(role.copyFrom) : null;                         // start from a template role, else deny-all
      sh.appendRow([key, String(role.label || key).trim().slice(0, 40), COLOR_RX.test(role.color) ? role.color : '#0074D9', data.length,
                    0, role.hidden_signup ? 1 : 0, JSON.stringify(src ? src.perms : rbacDefaultPerms_(key))]);   // is_super always 0 — owner is seeded only
      bustRoles_(); addLog_(me.u, 'Rol Agregado', key + (src ? ' (copiado de ' + src.key + ')' : ''));
      return ok_({ message:'Rol agregado' });
    });
  });
}

// rename / recolor — the KEY never changes, Users.Role points at it
function updateRole(tok, roleKey, role) {
  return api_(tok, 'rbac', null, function(me){
    role = role || {};
    return withLock_(function(){
      var sh = ensureRbac_(), data = sh.getDataRange().getValues(), i = roleRowIx_(data, roleKey);
      if (i === -1) return err_('Rol no encontrado');
      var r = data[i];
      r[ROLE_C.LABEL] = String(role.label || roleKey).trim().slice(0, 40);
      r[ROLE_C.COLOR] = COLOR_RX.test(role.color) ? role.color : '#0074D9';
      r[ROLE_C.HIDDEN] = role.hidden_signup ? 1 : 0;
      sh.getRange(i + 1, ROLE_C.LABEL + 1, 1, ROLE_C.HIDDEN - ROLE_C.LABEL + 1).setValues([r.slice(ROLE_C.LABEL, ROLE_C.HIDDEN + 1)]);  // one write
      bustRoles_(); addLog_(me.u, 'Rol Actualizado', roleKey);
      return ok_({ message:'Rol actualizado' });
    });
  });
}

// delete — refused while any user still holds it, so nobody is orphaned without perms
function deleteRole(tok, roleKey) {
  return api_(tok, 'rbac', null, function(me){
    var r = roleByKey_(roleKey);
    if (!r) return err_('Rol no encontrado');
    if (r.is_super) return err_('El rol de propietario no se puede eliminar');
    if (canEditRbac_(roleKey)) return err_('Este rol administra permisos y no se puede eliminar');
    return withLock_(function(){
      var n = userRows_().data.filter(function(x){ return x[U.ROLE] === roleKey; }).length;
      if (n) return err_(n + ' usuario(s) todavía tienen este rol — asígnales otro primero');
      var sh = ensureRbac_(), j = roleRowIx_(sh.getDataRange().getValues(), roleKey);
      if (j !== -1) sh.deleteRow(j + 1);
      bustRoles_(); addLog_(me.u, 'Rol Eliminado', roleKey);
      return ok_({ message:'Rol eliminado' });
    });
  });
}

// ============== App Settings (key | value) ==============
var SET_KEY = 'settings_v1';
var DEFAULT_HOURS = [0, 1, 2, 3, 4, 5, 6].map(function(d){ return { day:d, open:'11:00', close:'21:30', closed:0 }; });   // 0 = Sunday
var SET_DEFAULTS = {
  shop_name:'Bebidas Demo', shop_logo:'', shop_address:'', shop_phone:'', shop_email:'', map_url:'', about_text:'',
  hero_title:'Tu bebida, a tu gusto', hero_subtitle:'', hero_image:'',
  currency_symbol:'$', currency_decimals:0,
  hours:JSON.stringify(DEFAULT_HOURS), always_open:0, ordering_enabled:1, pause_message:'', last_order_minutes:30, prep_minutes:15, slot_minutes:15,
  pickup_enabled:1, delivery_enabled:0, delivery_fee:0, delivery_note:'', min_order_amount:0,
  sugar_levels:JSON.stringify(['Normal', 'Menos', 'Mitad', 'Poca', 'Sin azúcar']), ice_levels:JSON.stringify(['Hielo normal', 'Poco hielo', 'Sin hielo', 'Caliente']),
  max_receipt_mb:5, reupload_max:3, notify_shop:0, notify_shop_email:'', notify_customer:0, notify_events:'shop.new_order,shop.reupload,shop.cust_cancel,customer.received,customer.approved,customer.rejected,customer.ready,customer.cancelled,customer.refunded,customer.welcome',
  footer_note:'', receipt_size:'80mm', date_format:'dd-MMM-yyyy', time_format:'hh:mm a', backup_keep_days:30,
  assets_folder_id:'', receipts_folder_id:'', backup_folder_id:''
};
var SET_JSON = { hours:DEFAULT_HOURS, sugar_levels:['Normal'], ice_levels:['Hielo normal'] };   // stored as JSON text, read as arrays
var DATE_FORMATS = ['dd-MMM-yyyy', 'dd-MM-yyyy', 'yyyy-MM-dd'], TIME_FORMATS = ['hh:mm a', 'HH:mm'];
var HM_RX = /^([01]\d|2[0-3]):[0-5]\d$/;
// the keys a shopper may see — never folder ids, notification emails or backup settings
var PUBLIC_KEYS = ['shop_name', 'shop_address', 'shop_phone', 'shop_email', 'map_url', 'about_text', 'hero_title', 'hero_subtitle',
  'currency_symbol', 'currency_decimals', 'hours', 'always_open', 'ordering_enabled', 'pause_message', 'last_order_minutes', 'prep_minutes', 'slot_minutes',
  'pickup_enabled', 'delivery_enabled', 'delivery_fee', 'delivery_note', 'min_order_amount', 'sugar_levels', 'ice_levels', 'max_receipt_mb',
  'reupload_max', 'footer_note', 'receipt_size', 'date_format', 'time_format'];

function settingsSheet_() { return sh_(SETTINGS_SHEET) || (_m.sh[SETTINGS_SHEET] = head_(ss_().insertSheet(SETTINGS_SHEET), SET_HEAD)); }
// a Drive file id -> its lh3 link; a full https link (demo photos) passes through untouched
var imgUrl_ = function(id){ return !id ? '' : String(id).indexOf('https://') === 0 ? id : 'https://lh3.google.com/u/0/d/' + id; };

// defaults <- sheet, typed by the default's own type; JSON keys come back parsed; cached (every storefront hit reads it)
function settings_() {
  if (_m.set) return _m.set;
  var hit = safeParse_(CACHE_.get(SET_KEY), null);
  if (hit) return (_m.set = hit);
  var s = sh_(SETTINGS_SHEET), raw = {}, out = {};
  if (s && s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(function(r){ if (r[0]) raw[String(r[0])] = r[1]; });
  Object.keys(SET_DEFAULTS).forEach(function(k){
    var def = SET_DEFAULTS[k], v = raw.hasOwnProperty(k) && raw[k] !== '' ? raw[k] : null;
    out[k] = v === null ? def : typeof def === 'number' ? (isNaN(Number(v)) ? def : Number(v)) : String(v);
    if (SET_JSON[k]) { var a = safeParse_(out[k], null); out[k] = Array.isArray(a) && a.length ? a : SET_JSON[k]; }
  });
  cachePut_(SET_KEY, out, CACHE_TTL);
  return (_m.set = out);
}

// whole tab rewritten in one write, every cell plain text (phone / prices keep their exact text)
function writeSettings_(s) {
  var sh = settingsSheet_(), last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 2).clearContent();
  putText_(sh, 2, Object.keys(SET_DEFAULTS).map(function(k){ var v = s[k]; return [k, SET_JSON[k] && typeof v !== 'string' ? JSON.stringify(v) : String(v == null ? '' : v)]; }));
  CACHE_.remove(SET_KEY); _m.set = null; bustFront_();
  return settings_();
}
// one-key system write (folder ids, logo, banner) — keeps everything else
var settingSet_ = function(k, v){ return withLock_(function(){ var s = JSON.parse(JSON.stringify(settings_())); s[k] = v; return writeSettings_(s); }); };

// 7 day rows, HH:mm, close after open unless the day is closed
function cleanHours_(v) {
  var a = Array.isArray(v) ? v : safeParse_(v, null);
  if (!Array.isArray(a) || a.length !== 7) fail_('El horario de apertura requiere los 7 días');
  var seen = {};
  return a.map(function(h){
    var day = Number(h && h.day), closed = flag_(h.closed, 0), open = String(h.open || '').trim(), close = String(h.close || '').trim();
    if (!(day >= 0 && day <= 6) || seen[day]) fail_('Horario de apertura: cada día una sola vez'); seen[day] = 1;
    if (!closed && (!HM_RX.test(open) || !HM_RX.test(close))) fail_('El horario debe tener el formato 11:00');
    if (!closed && close <= open) fail_('La hora de cierre debe ser posterior a la de apertura');
    return { day:day, open:HM_RX.test(open) ? open : '11:00', close:HM_RX.test(close) ? close : '21:30', closed:closed };
  }).sort(function(x, y){ return x.day - y.day; });
}
// sugar / ice pick-lists: 1-8 unique labels, 1-20 chars
function cleanLevels_(v, what) {
  var a = (Array.isArray(v) ? v : safeParse_(v, [])).map(function(x){ return String(x == null ? '' : x).trim().slice(0, 20); }).filter(Boolean), seen = {};
  if (!a.length || a.length > 8) fail_(what + ': 1-8 opciones');
  a.forEach(function(x){ var k = x.toLowerCase(); if (seen[k]) fail_(what + ': "' + x + '" está repetido'); seen[k] = 1; });
  return a;
}
var money_ = function(v, msg){ return num_(v, 0, 100000, msg, 2); };

// the settings form, validated against the stored copy — system keys (folder ids, images) stay as they were
function cleanSettings_(d, cur) {
  var s = {}, str = function(k, max){ return String(d[k] == null ? '' : d[k]).trim().slice(0, max); };
  var int = function(k, lo, hi, msg){ var n = Number(d[k]); if (d[k] === '' || d[k] == null || isNaN(n) || n < lo || n > hi || Math.round(n) !== n) fail_(msg); return n; };
  Object.keys(SET_DEFAULTS).forEach(function(k){ s[k] = cur[k]; });
  s.shop_name = str('shop_name', 80);                 if (!s.shop_name) fail_('El nombre de la tienda es obligatorio');
  s.shop_address = str('shop_address', 300);
  s.shop_phone = str('shop_phone', 20);               if (s.shop_phone && !PHONE_RX.test(s.shop_phone)) fail_('Teléfono de la tienda: 7-20 dígitos');
  s.shop_email = str('shop_email', 100);              if (s.shop_email && !EMAIL_RX.test(s.shop_email)) fail_('Ingresa un correo de tienda válido');
  s.map_url = str('map_url', 255);                    if (s.map_url && !/^https:\/\/\S+$/.test(s.map_url)) fail_('El enlace de mapa debe comenzar con https://');
  s.about_text = str('about_text', 600);
  s.hero_title = str('hero_title', 80);               if (!s.hero_title) fail_('El título principal (Hero) es obligatorio');
  s.hero_subtitle = str('hero_subtitle', 160);
  if (d.shop_logo === '') s.shop_logo = '';           // cleared from the form; a new image only arrives via upload
  if (d.hero_image === '') s.hero_image = '';
  s.currency_symbol = str('currency_symbol', 5);      if (!s.currency_symbol) fail_('El símbolo de moneda es obligatorio');
  s.currency_decimals = Number(d.currency_decimals) === 2 ? 2 : 0;
  s.hours = cleanHours_(d.hours);                     // kept even while always open — they come back when it is switched off
  s.always_open = flag_(d.always_open, 0);
  s.ordering_enabled = flag_(d.ordering_enabled, 1);
  s.pause_message = str('pause_message', 160);
  s.last_order_minutes = int('last_order_minutes', 0, 240, 'Límite de último pedido: 0-240 minutos');
  s.prep_minutes = int('prep_minutes', 1, 240, 'Tiempo de preparación: 1-240 minutos');
  s.slot_minutes = Number(d.slot_minutes) === 30 ? 30 : 15;
  s.pickup_enabled = flag_(d.pickup_enabled, 1); s.delivery_enabled = flag_(d.delivery_enabled, 0);
  if (!s.pickup_enabled && !s.delivery_enabled) fail_('Habilita recogida, entrega a domicilio o ambas');
  s.delivery_fee = money_(d.delivery_fee, 'El costo de envío debe ser 0 o más');
  s.delivery_note = str('delivery_note', 160);
  s.min_order_amount = money_(d.min_order_amount, 'El pedido mínimo debe ser 0 o más');
  s.sugar_levels = cleanLevels_(d.sugar_levels, 'Niveles de azúcar');
  s.ice_levels = cleanLevels_(d.ice_levels, 'Niveles de hielo');
  s.max_receipt_mb = int('max_receipt_mb', 1, 10, 'Límite de tamaño de comprobante: 1-10 MB');
  s.reupload_max = int('reupload_max', 1, 5, 'Intentos de comprobante: 1-5');
  s.notify_shop = flag_(d.notify_shop, 0); s.notify_customer = flag_(d.notify_customer, 0);
  s.notify_shop_email = mailList_(str('notify_shop_email', 300));
  if (s.notify_shop && !s.notify_shop_email) fail_('Ingresa el correo que recibe las alertas de nuevos pedidos');
  s.notify_events = String(d.notify_events == null ? cur.notify_events : d.notify_events).split(',').map(function(x){ return x.trim(); })
    .filter(function(x, i, a){ return MAIL_EVENTS.indexOf(x) !== -1 && a.indexOf(x) === i; }).join(',');
  s.footer_note = str('footer_note', 200);
  if (['80mm', '58mm'].indexOf(d.receipt_size) === -1) fail_('El tamaño del recibo debe ser 80mm o 58mm');
  s.receipt_size = d.receipt_size;
  s.date_format = oneOf_(d.date_format, DATE_FORMATS, 'dd-MMM-yyyy', 'Selecciona un formato de fecha');
  s.time_format = oneOf_(d.time_format, TIME_FORMATS, 'hh:mm a', 'Selecciona un formato de hora');
  s.backup_keep_days = int('backup_keep_days', 1, 365, 'Guardar respaldos: 1-365 días');
  return s;
}
// what the admin screens get: every key + image urls (settings hold ids)
var setOut_ = function(s){ var o = JSON.parse(JSON.stringify(s)); o.shop_logo_url = imgUrl_(s.shop_logo); o.hero_image_url = imgUrl_(s.hero_image); return o; };

// every signed-in user reads it — prints need the letterhead (no secrets live here)
// staff with Settings view get every key; anyone else (a signed-in customer) only the public shop keys
function getAppSettings(tok) {
  return api_(tok, null, null, function(me){ var s = settings_(); return ok_({ settings:hasPerm_(me.role, 'settings', 'v') ? setOut_(s) : publicShop_(s), tz:tz_() }); });
}

function saveSettings(tok, d) {
  return api_(tok, 'settings', 'e', function(me){
    return withLock_(function(){
      var cur = settings_(), s = writeSettings_(cleanSettings_(d || {}, cur)),
          diff = Object.keys(SET_DEFAULTS).filter(function(k){ return JSON.stringify(cur[k]) !== JSON.stringify(s[k]); });
      addLog_(me.u, 'AJUSTES_GUARDADOS', diff.length ? 'Cambiado: ' + diff.join(', ') : 'Sin cambios');
      return ok_({ message:'Ajustes guardados', settings:setOut_(s) });
    });
  });
}

// logo / hero banner -> ASSETS/Shop (link-viewable public assets); the setting keeps the file id
function uploadShopImage(tok, which, base64Data, filename) {
  return api_(tok, 'settings', 'e', function(me){
    if (['shop_logo', 'hero_image'].indexOf(which) === -1) return err_('Imagen desconocida');
    var img = saveImage_(base64Data, filename, which, shopFolder_()), s = settingSet_(which, img.file.getId());
    addLog_(me.u, 'AJUSTES_GUARDADOS', (which === 'shop_logo' ? 'Logo de tienda' : 'Banner principal') + ' subido');
    return ok_({ message:'Imagen subida', settings:setOut_(s) });
  });
}

// the one pause switch — also on the Order Board header
function setOrderingEnabled(tok, on) {
  return api_(tok, 'settings', 'e', function(me){
    var s = settingSet_('ordering_enabled', on ? 1 : 0);
    addLog_(me.u, 'AJUSTES_GUARDADOS', 'Pedidos en línea ' + (on ? 'reanudados' : 'pausados'));
    return ok_({ message:on ? 'Los pedidos en línea están abiertos nuevamente' : 'Pedidos en línea pausados', settings:setOut_(s) });
  });
}

// ============== Field readers (shared by every cleaner — a bad value throws a clean message) ==============
var str_ = function(d, k, max){ return String(d[k] == null ? '' : d[k]).trim().slice(0, max); };
var num_ = function(v, lo, hi, msg, dp){
  var n = v === '' || v == null ? 0 : Number(String(v).replace(/,/g, ''));
  if (isNaN(n) || n < lo || n > hi) fail_(msg);
  return dp == null ? n : Math.round(n * Math.pow(10, dp)) / Math.pow(10, dp); };
var int_ = function(v, lo, hi, msg){ var n = num_(v, lo, hi, msg); if (Math.round(n) !== n) fail_(msg); return n; };
var flag_ = function(v, def){ return v === undefined || v === null || v === '' ? def : (v === true || v === 1 || /^(1|yes|y|true|active|on)$/i.test(String(v).trim()) ? 1 : 0); };
var oneOf_ = function(v, list, def, msg){ var s = String(v == null || v === '' ? def : v).trim().toLowerCase();
  var hit = list.filter(function(x){ return String(x).toLowerCase() === s; })[0]; if (hit === undefined) fail_(msg); return hit; };
var strip_ = function(r){ var o = JSON.parse(JSON.stringify(r)); delete o.deleted; return o; };
var todayYmd_ = function(){ return ymd_(new Date()); };
var addDays_ = function(day, n){ var d = new Date(day + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
// money in the settings' precision (0 or 2 decimals)
var round_ = function(v){ var dp = Number(settings_().currency_decimals) === 2 ? 2 : 0, f = Math.pow(10, dp); return Math.round((Number(v) || 0) * f + 1e-9) / f; };

// ============== JSON-row store (Date | Data) ==============
// one row per day, that day's records as a JSON array; a full cell (45k chars) spills into another row of the same date.
// reads go through the day index: column A first, then ONE contiguous block of column B for the dates asked for
var JDB_MAX = 45000;
function jdbSheet_(n) {
  var s = sh_(n);
  if (s) return s;
  s = _m.sh[n] = head_(ss_().insertSheet(n), ['Date', 'Data']);
  s.getRange('A:A').setNumberFormat('@');
  return s;
}
// a write drops every memo of that sheet — the next read (inside the same lock) sees the sheet as written
var jdbDrop_ = function(n){ [_m.jdb, _m.jk].forEach(function(c){ if (c) delete c[n]; });
  if (_m.jr) Object.keys(_m.jr).forEach(function(k){ if (k.indexOf(n + '|') === 0) delete _m.jr[k]; }); };
var inDays_ = function(d, from, to){ return (!from || d >= from) && (!to || d <= to); };
var JDB = {
  // [{row, date, arr}] — the whole sheet, once per request
  rows: function(n){
    var c = _m.jdb || (_m.jdb = {});
    if (c[n]) return c[n];
    var s = jdbSheet_(n), last = s.getLastRow();
    return (c[n] = last < 2 ? [] : s.getRange(2, 1, last - 1, 2).getValues().map(function(v, i){
      return { row:i + 2, date:ymd_(v[0]), arr:safeParse_(v[1], []) || [] }; }));
  },
  // day keys only (column A) — cheap, the index every ranged read starts from
  keys: function(n){
    var c = _m.jk || (_m.jk = {});
    if (c[n]) return c[n];
    var s = jdbSheet_(n), last = s.getLastRow();
    return (c[n] = last < 2 ? [] : s.getRange(2, 1, last - 1, 1).getValues().map(function(v){ return ymd_(v[0]); }));
  },
  // rows whose day is inside from..to — one block read spanning the first..last match (days are appended in order)
  range: function(n, from, to){
    var c = _m.jr || (_m.jr = {}), k = n + '|' + (from || '') + '|' + (to || '');
    if (c[k]) return c[k];
    if (_m.jdb && _m.jdb[n]) return (c[k] = _m.jdb[n].filter(function(r){ return inDays_(r.date, from, to); }));
    var ks = JDB.keys(n), lo = -1, hi = -1;
    ks.forEach(function(d, i){ if (inDays_(d, from, to)) { if (lo === -1) lo = i; hi = i; } });
    if (lo === -1) return (c[k] = []);
    var vals = jdbSheet_(n).getRange(lo + 2, 2, hi - lo + 1, 1).getValues(), out = [];
    vals.forEach(function(v, j){ var d = ks[lo + j]; if (inDays_(d, from, to)) out.push({ row:lo + j + 2, date:d, arr:safeParse_(v[0], []) || [] }); });
    return (c[k] = out);
  },
  flat_: function(rows){ var out = []; rows.forEach(function(r){ r.arr.forEach(function(x){ if (!x.deleted) out.push(x); }); }); return out; },
  all: function(n){ return JDB.flat_(JDB.rows(n)); },
  between: function(n, from, to){ return JDB.flat_(JDB.range(n, from, to)); },
  // reserves `count` ids in one property write; returns the first
  nextId: function(n, count){
    var p = PropertiesService.getScriptProperties(), k = 'JID_' + n, id = Number(p.getProperty(k)) ||
      JDB.rows(n).reduce(function(m, r){ return r.arr.reduce(function(mm, x){ return Math.max(mm, x.id || 0); }, m); }, 0);
    p.setProperty(k, String(id + (count || 1)));
    return id + 1;
  },
  // bulk insert on today's row(s) — existing day rows rewritten once, new rows appended in ONE setValues
  insertMany: function(n, recs, by){
    if (!_m.locked) fail_('Write outside the lock');
    if (!recs.length) return [];
    var s = jdbSheet_(n), now = nowIso_(), date = ymd_(now), day = JDB.range(n, date, date), first = JDB.nextId(n, recs.length), touched = [], fresh = [];
    var cur = day[day.length - 1] || null;
    if (cur) cur.len = JSON.stringify(cur.arr).length;
    recs.forEach(function(rec, i){
      rec.id = first + i; rec.created = now; rec.updated = now; rec.created_by = rec.created_by || by || ''; rec.deleted = 0;
      var add = JSON.stringify(rec).length + 1;
      if (!cur || cur.len + add > JDB_MAX) { cur = { row:0, date:date, arr:[], len:2 }; fresh.push(cur); }
      cur.arr.push(rec); cur.len += add;
      if (cur.row && touched.indexOf(cur) === -1) touched.push(cur);
    });
    touched.forEach(function(r){ s.getRange(r.row, 2).setValue(JSON.stringify(r.arr)); });
    if (fresh.length) s.getRange(s.getLastRow() + 1, 1, fresh.length, 2).setNumberFormat('@').setValues(fresh.map(function(r){ return [date, JSON.stringify(r.arr)]; }));
    jdbDrop_(n);
    return recs;
  },
  insert: function(n, rec, by){ return JDB.insertMany(n, [rec], by)[0]; },
  // per-record edit: fn(rec) validates first and returns false to skip; one write per touched day row.
  // from/to = the days to search (the record's created day) — omit only for small sheets
  mutate: function(n, ids, fn, from, to){
    if (!_m.locked) fail_('Write outside the lock');
    var want = {}, now = nowIso_(), s = jdbSheet_(n), hit = [];
    ids.forEach(function(id){ want[String(id)] = 1; });
    (from || to ? JDB.range(n, from, to) : JDB.rows(n)).forEach(function(r){
      var t = false;
      r.arr.forEach(function(x){ if (!want[String(x.id)] || x.deleted || fn(x) === false) return; x.updated = now; t = true; hit.push(x); });
      if (t) s.getRange(r.row, 2).setValue(JSON.stringify(r.arr));
    });
    jdbDrop_(n);
    return hit;
  },
  patchMany: function(n, ids, patch, from, to){ return JDB.mutate(n, ids, function(x){ Object.keys(patch).forEach(function(k){ x[k] = patch[k]; }); }, from, to); },
  update: function(n, id, patch, day){ return JDB.patchMany(n, [id], patch, day, day)[0] || null; },
  // one record by id when its day is known (fast) or by a full scan
  get: function(n, id, day){ var a = day ? JDB.between(n, day, day) : JDB.all(n); for (var i = 0; i < a.length; i++) if (a[i].id == id) return a[i]; return null; }
};
// seed writer: records carry their own created stamp; grouped per day in date order, ONE setValues (never insert in a loop)
function jdbSeed_(n, recs, by) {
  var s = jdbSheet_(n), days = {}, first = JDB.nextId(n, recs.length);
  recs.forEach(function(r, i){ r.id = first + i; r.created = r.created || nowIso_(); r.updated = r.updated || r.created; r.created_by = r.created_by || by || 'system'; r.deleted = r.deleted || 0;
    var d = ymd_(r.created); (days[d] = days[d] || []).push(r); });
  var rows = [];
  Object.keys(days).sort().forEach(function(d){ var cur = [], len = 2;
    days[d].sort(function(a, b){ return a.created < b.created ? -1 : 1; }).forEach(function(r){ var add = JSON.stringify(r).length + 1;
      if (len + add > JDB_MAX && cur.length) { rows.push([d, JSON.stringify(cur)]); cur = []; len = 2; }
      cur.push(r); len += add; });
    if (cur.length) rows.push([d, JSON.stringify(cur)]); });
  if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  jdbDrop_(n);
  return recs;
}

// ============== Document numbers — PREFIX-YYYYMMDD-NNNN, one counter per prefix per day, taken inside the write lock ==============
var noDay_ = function(){ return todayYmd_().replace(/-/g, ''); };
function nextNo_(pfx, sheet, field) {
  if (!_m.locked) fail_('Numbering must run inside the write lock');
  var day = noDay_(), p = PropertiesService.getScriptProperties(), k = 'SEQ_' + pfx + '_' + day, n = Number(p.getProperty(k)) || 0, used = {}, no;
  if (sheet && sh_(sheet)) JDB.between(sheet, todayYmd_(), todayYmd_()).forEach(function(x){ used[x[field]] = 1; });   // re-check against today's records
  do { n++; no = pfx + '-' + day + '-' + (n < 10000 ? ('000' + n).slice(-4) : String(n)); } while (used[no]);
  p.setProperty(k, String(n));
  return no;
}

// ============== Drive — public shop assets (link-viewable) vs private receipts ==============
// base64 image -> folder; Drive write stays outside any lock. The client already downscaled it to JPEG / PNG
function saveImage_(base64Data, filename, prefix, folder) {
  var raw = String(base64Data || ''), b64 = raw.split(',').pop(), mime = (/^data:(image\/(jpeg|png|webp));/.exec(raw) || [])[1] || 'image/jpeg';
  if (!b64) fail_('No hay datos de imagen');
  if (b64.length > 4000000) fail_('Imagen demasiado grande — máx. 3 MB');                   // base64 ~4/3 of bytes
  filename = String(filename || 'image').replace(/[^\w.\-]/g, '_').slice(0, 60);   // never trust a client filename
  var file = (folder || getAssetsFolder_()).createFile(Utilities.newBlob(Utilities.base64Decode(b64), mime, filename))
    .setName(prefix + '_' + Date.now() + '_' + filename).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { file:file, url:imgUrl_(file.getId()) };
}
// a child folder by name, created once; the id is remembered in Script Properties
function subFolder_(parent, name, key) {
  var p = PropertiesService.getScriptProperties(), id = p.getProperty(key), hit = null;
  if (id) { try { hit = DriveApp.getFolderById(id); } catch (e) { hit = null; } }
  if (!hit) { var it = parent.getFoldersByName(name); hit = it.hasNext() ? it.next() : parent.createFolder(name); p.setProperty(key, hit.getId()); }
  return hit;
}
// ASSETS/Shop — product photos, logo, banner, payment QR codes (the storefront shows them to anyone)
var shopFolder_ = function(){ return _m.shopF || (_m.shopF = subFolder_(getAssetsFolder_(), 'Shop', 'DRV_SHOP')); };
// receipts live OUTSIDE the shared ASSETS tree (a child of a link-shared folder is link-shared too)
function receiptsRoot_() {
  if (_m.rcpF) return _m.rcpF;
  var p = PropertiesService.getScriptProperties(), id = p.getProperty('DRV_RCP'), hit = null;
  if (id) { try { hit = DriveApp.getFolderById(id); } catch (e) { hit = null; } }
  if (!hit) { var it = DriveApp.getFoldersByName(RECEIPTS_FOLDER_NAME); hit = it.hasNext() ? it.next() : DriveApp.createFolder(RECEIPTS_FOLDER_NAME); p.setProperty('DRV_RCP', hit.getId()); }
  return (_m.rcpF = hit);
}

// ============== Opening hours — the shop's own clock (script time zone) ==============
var hmMin_ = function(hm){ var p = String(hm || '0:0').split(':'); return (+p[0]) * 60 + (+p[1] || 0); };
var minHm_ = function(n){ n = Math.max(0, Math.min(1439, n)); return ('0' + Math.floor(n / 60)).slice(-2) + ':' + ('0' + (n % 60)).slice(-2); };
var dow_ = function(day){ return new Date(day + 'T12:00:00Z').getUTCDay(); };
// is the shop taking online orders right now, and if not, why + when it opens again
function openState_(s, at) {
  var now = at || new Date(), day = ymd_(now), hm = Utilities.formatDate(now, tz_(), 'HH:mm'), hours = s.hours || DEFAULT_HOURS;
  // always open: every day 00:00-23:59, no last-order cut-off; a pause still stops orders
  if (Number(s.always_open)) { var paused = !Number(s.ordering_enabled);
    return { open:!paused, reason:paused ? 'paused' : '', always:1, today:{ date:day, open:'00:00', close:'23:59', closed:0, cutoff:'23:59' }, now_hm:hm, next_open:null }; }
  var of = function(d){ return hours.filter(function(h){ return h.day === dow_(d); })[0] || { closed:1 }; };
  var t = of(day), cut = t.closed ? '' : minHm_(hmMin_(t.close) - (Number(s.last_order_minutes) || 0));
  var reason = !Number(s.ordering_enabled) ? 'paused' : t.closed ? 'closed_today' : hm < t.open ? 'before_open' : hm >= cut ? 'after_cutoff' : '';
  var next = null;
  for (var i = 0; i < 8 && !next; i++) { var d = addDays_(day, i), h = of(d);
    if (!h.closed && (i > 0 || hm < h.open)) next = { date:d, time:h.open }; }
  return { open:!reason, reason:reason, today:{ date:day, open:t.open || '', close:t.close || '', closed:t.closed ? 1 : 0, cutoff:cut }, now_hm:hm, next_open:next };
}

// ============== Storefront (public) — ONE cached payload, the open state computed per call ==============
var FRONT_KEY = 'front_v1';
var bustFront_ = function(){ try { CACHE_.remove(FRONT_KEY); } catch (e) {} _m.front = null; };
// guarded read: a sheet this install hasn't built yet = an empty list (never created by a public read)
var jdbList_ = function(n){ return sh_(n) ? JDB.all(n) : []; };
// the menu + payment methods a shopper may see (filled in by the menu and payment phases)
function front_() {
  if (_m.front) return _m.front;
  var hit = safeParse_(CACHE_.get(FRONT_KEY), null);
  if (hit) return (_m.front = hit);
  var out = menuPublic_();
  cachePut_(FRONT_KEY, out, CACHE_TTL);
  return (_m.front = out);
}
function publicShop_(s) {
  var o = {}; PUBLIC_KEYS.forEach(function(k){ o[k] = s[k]; });
  o.shop_logo_url = imgUrl_(s.shop_logo); o.hero_image_url = imgUrl_(s.hero_image);
  return o;
}
// no token needed — a warm cache reads no sheet at all
function getStorefront(tok) {
  return pub_(function(){
    var s = settings_(), f = front_(), me = whoMaybe_(tok);
    return ok_({ shop:publicShop_(s), open:openState_(s), categories:f.categories, products:f.products, addons:f.addons, methods:f.methods,
                 now:nowIso_(), tz:tz_(), signedIn:!!me });
  });
}

// ============== Menu — categories · add-ons · products on ONE registry (save / archive / import / order) ==============
var CAT_ICONS = ['fa-mug-hot', 'fa-mug-saucer', 'fa-glass-water', 'fa-lemon', 'fa-apple-whole', 'fa-seedling', 'fa-leaf', 'fa-ice-cream', 'fa-snowflake',
  'fa-blender', 'fa-wine-bottle', 'fa-bottle-water', 'fa-whiskey-glass', 'fa-martini-glass', 'fa-champagne-glasses', 'fa-cookie', 'fa-cake-candles',
  'fa-carrot', 'fa-star', 'fa-fire', 'fa-heart', 'fa-bolt', 'fa-droplet', 'fa-cubes-stacked'];
var MAX_SIZES = 4, MAX_ADDON_PICK = 3;
// yes / no from a form flag or a CSV cell
var yes_ = function(v, def){ return flag_(v, def); };

function cleanCategory_(d, cur) {
  var name = str_(d, 'name', 40); if (!name) fail_('Ingresa el nombre de la categoría');
  var icon = str_(d, 'icon', 40) || (cur && cur.icon) || 'fa-mug-hot';
  if (CAT_ICONS.indexOf(icon) === -1) icon = 'fa-mug-hot';
  return { name:name, icon:icon, tagline:str_(d, 'tagline', 60), sort_order:cur ? cur.sort_order : -1, is_active:yes_(d.is_active, 1) };
}
function cleanAddon_(d, cur) {
  var name = str_(d, 'name', 40); if (!name) fail_('Ingresa el nombre del complemento');
  if (d.price === '' || d.price == null) fail_('Ingresa el precio del complemento (se permite 0)');
  return { name:name, price:money_(d.price, 'El precio del complemento debe ser 0 o más'), is_available:yes_(d.is_available, 1), sort_order:cur ? cur.sort_order : -1, is_active:yes_(d.is_active, 1) };
}
// sizes: [{name, price_delta}] from the form, or "Medium:0|Large:0.5" from a CSV cell
function cleanSizes_(v) {
  var a = Array.isArray(v) ? v : String(v || '').split('|').map(function(x){ var p = x.split(':'); return { name:p[0], price_delta:p[1] }; });
  a = a.map(function(s){ return { name:String((s && s.name) || '').trim().slice(0, 20), price_delta:s && s.price_delta }; }).filter(function(s){ return s.name; });
  if (a.length > MAX_SIZES) fail_('Hasta ' + MAX_SIZES + ' tamaños');
  var seen = {};
  return a.map(function(s){ var k = s.name.toLowerCase(); if (seen[k]) fail_('El tamaño "' + s.name + '" está repetido'); seen[k] = 1;
    return { name:s.name, price_delta:num_(s.price_delta, 0, 100000, 'Tamaño "' + s.name + '": la diferencia de precio debe ser 0 o más', 2) }; });
}
// ctx = { cats: byId, catByName, addons: byId, addonByName } built once per save / import
function cleanProduct_(d, cur, ctx) {
  var name = str_(d, 'name', 60); if (!name) fail_('Ingresa el nombre de la bebida');
  var cat = ctx.cats[Number(d.category_id)] || ctx.catByName[String(d.category || d.category_id || '').trim().toLowerCase()];
  if (!cat) fail_('Selecciona una categoría');
  if (!cat.is_active && !(cur && cur.category_id === cat.id)) fail_('La categoría "' + cat.name + '" está archivada — elige otra');
  var base = num_(d.base_price, 0, 100000, 'El precio base debe ser un número', 2); if (!(base > 0)) fail_('El precio base debe ser mayor a 0');
  var ids = Array.isArray(d.addon_ids) ? d.addon_ids : String(d.addons || '').split('|').map(function(x){ x = x.trim().toLowerCase(); var a = ctx.addonByName[x]; if (x && !a) fail_('Complemento desconocido "' + x + '"'); return a ? a.id : 0; });
  var seen = {}, addonIds = [];
  ids.forEach(function(x){ var a = ctx.addons[Number(x)]; if (!Number(x)) return; if (!a) fail_('Complemento desconocido #' + x);
    if (!a.is_active) return; if (!seen[a.id]) { seen[a.id] = 1; addonIds.push(a.id); } });
  var max = d.max_addons === '' || d.max_addons == null ? MAX_ADDON_PICK : int_(d.max_addons, 0, MAX_ADDON_PICK, 'Máx. complementos debe ser de 0 a ' + MAX_ADDON_PICK);
  return { name:name, category_id:cat.id, description:str_(d, 'description', 200), base_price:base, sizes:cleanSizes_(d.sizes),
    has_sugar:yes_(d.has_sugar, 1), has_ice:yes_(d.has_ice, 1), addon_ids:addonIds, max_addons:Math.min(max, addonIds.length),
    is_featured:yes_(d.is_featured, 0), is_available:yes_(d.is_available, 1), image_id:cur ? cur.image_id || '' : '',
    sort_order:cur && cur.category_id === cat.id ? cur.sort_order : -1, is_active:yes_(d.is_active, 1) };
}
var menuCtx_ = function(){
  var cats = {}, catByName = {}, addons = {}, addonByName = {};
  JDB.all(CAT_SHEET).forEach(function(c){ cats[c.id] = c; catByName[String(c.name).toLowerCase()] = c; });
  JDB.all(ADDON_SHEET).forEach(function(a){ addons[a.id] = a; addonByName[String(a.name).toLowerCase()] = a; addonByName[String(a.addon_code).toLowerCase()] = a; });
  return { cats:cats, catByName:catByName, addons:addons, addonByName:addonByName };
};

// category archive guard: active drinks still in it
var catBusy_ = function(r){ var n = JDB.all(PROD_SHEET).filter(function(p){ return p.category_id === r.id && p.is_active; }).length;
  return n ? r.name + ' todavía tiene ' + n + ' bebida' + (n === 1 ? ' activa' : 's activas') + ' — muévelas o archívalas primero' : ''; };

var MENU = {
  category: { sheet:CAT_SHEET, page:'categories', label:'Categoría', clean:cleanCategory_, code:['cat_code', 'CAT-', 2], nm:function(r){ return r.name; },
              uniq:[['name', 'Ya existe una categoría llamada']], guard:function(r, op){ return op === 'off' ? catBusy_(r) : ''; }, log:'CATEGORIA',
              csv:[['Name', 'name'], ['Icon', 'icon'], ['Tagline', 'tagline'], ['Active', 'is_active']] },
  addon:    { sheet:ADDON_SHEET, page:'addons', label:'Complemento', clean:cleanAddon_, code:['addon_code', 'ADD-', 2], nm:function(r){ return r.name; },
              uniq:[['name', 'Ya existe un complemento llamado']], guard:function(){ return ''; }, log:'COMPLEMENTO',
              csv:[['Name', 'name'], ['Price', 'price'], ['Available', 'is_available'], ['Active', 'is_active']] },
  product:  { sheet:PROD_SHEET, page:'products', label:'Bebida', clean:cleanProduct_, code:['product_code', 'DRK-', 3], nm:function(r){ return r.name; },
              uniq:[['name', 'Ya existe una bebida llamada']], guard:function(){ return ''; }, log:'PRODUCTO', order:'category_id',
              csv:[['Name', 'name'], ['Category', 'category'], ['Description', 'description'], ['Base Price', 'base_price'], ['Sizes', 'sizes'],
                   ['Sugar', 'has_sugar'], ['Ice', 'has_ice'], ['Add-ons', 'addons'], ['Max Add-ons', 'max_addons'], ['Featured', 'is_featured'],
                   ['Available', 'is_available'], ['Active', 'is_active']] }
};
// unique fields, case-insensitive, among ACTIVE rows other than itself (an archived name can be reused)
function uniqCheck_(M, rec, id, pool) {
  M.uniq.forEach(function(u){
    var v = String(rec[u[0]] || '').toLowerCase();
    if (v && rec.is_active && pool.some(function(x){ return x.id != id && x.is_active && String(x[u[0]] || '').toLowerCase() === v; })) fail_(u[1] + ' "' + rec[u[0]] + '"');
  });
}
// next running code (CAT-01, ADD-01, DRK-001) — max existing + 1, inside the lock, never reused
var nextCode_ = function(M, pool){ var f = M.code[0], p = M.code[1], n = pool.reduce(function(m, x){ var k = Number(String(x[f] || '').replace(p, '')); return isNaN(k) ? m : Math.max(m, k); }, 0) + 1;
  return p + ('0000' + n).slice(-Math.max(M.code[2], String(n).length)); };
// ordered lists: a new row goes last (products: last within its category)
var nextOrder_ = function(M, pool, rec){ var of = M.order; return pool.filter(function(x){ return !of || x[of] === rec[of]; }).reduce(function(m, x){ return Math.max(m, Number(x.sort_order) || 0); }, 0) + 1; };

// usage + sales per row: menu rows are small, orders are read once for the last 30 days
function menuUsage_() {
  if (_m.usage) return _m.usage;
  var u = { catAll:{}, catSold:{}, addonProds:{}, prodSold:{}, prodLast:{}, addonSold:{} }, from = addDays_(todayYmd_(), -29);
  JDB.all(PROD_SHEET).forEach(function(p){
    if (p.is_active) { u.catAll[p.category_id] = (u.catAll[p.category_id] || 0) + 1; if (!p.is_available) u.catSold[p.category_id] = (u.catSold[p.category_id] || 0) + 1; }
    (p.addon_ids || []).forEach(function(a){ if (p.is_active) u.addonProds[a] = (u.addonProds[a] || 0) + 1; });
  });
  if (sh_(ORDER_SHEET)) JDB.between(ORDER_SHEET, from, null).forEach(function(o){
    if (!PAID_ST[o.status]) return;
    (o.items || []).forEach(function(l){ u.prodSold[l.product_id] = (u.prodSold[l.product_id] || 0) + l.qty;
      if (!u.prodLast[l.product_id] || o.placed_at > u.prodLast[l.product_id]) u.prodLast[l.product_id] = o.placed_at;
      (l.addons || []).forEach(function(a){ u.addonSold[a.addon_id] = (u.addonSold[a.addon_id] || 0) + l.qty; }); });
  });
  return (_m.usage = u);
}
var PAID_ST = { CONFIRMED:1, PREPARING:1, READY:1, COMPLETED:1 };
var fromPrice_ = function(p){ return round_(Number(p.base_price) + (p.sizes && p.sizes.length ? Math.min.apply(null, p.sizes.map(function(s){ return Number(s.price_delta) || 0; })) : 0)); };
function menuOut_(key, r, ctx) {
  var o = strip_(r), u = menuUsage_();
  if (key === 'category') { o.drinks = u.catAll[r.id] || 0; o.sold_out = u.catSold[r.id] || 0; }
  if (key === 'addon') { o.used_by = u.addonProds[r.id] || 0; o.sold_30 = u.addonSold[r.id] || 0; }
  if (key === 'product') {
    var c = ctx.cats[r.category_id] || {};
    o.category_name = c.name || ''; o.category_icon = c.icon || 'fa-mug-hot'; o.image_url = imgUrl_(r.image_id); o.from_price = fromPrice_(r);
    o.addon_names = (r.addon_ids || []).map(function(a){ return (ctx.addons[a] || {}).name; }).filter(Boolean);
    o.sold_30 = u.prodSold[r.id] || 0; o.last_ordered = u.prodLast[r.id] || '';
  }
  return o;
}

// menu photo / payment QR -> ASSETS/Shop (public), BEFORE the lock; the form sends a data url or clears it
function imageNew_(d, key) {
  if (d.image_new) return saveImage_(d.image_new, d.image_name || key + '.jpg', key, shopFolder_()).file.getId();
  return '';
}

function saveMenu_(tok, key, d) {
  d = d || {};
  var M = MENU[key], id = Number(d.id) || 0;
  return api_(tok, M.page, id ? 'e' : 'a', function(me){
    var img = M.img ? imageNew_(d, key) : '';                                // Drive write outside the lock
    return withLock_(function(){
      var pool = JDB.all(M.sheet), cur = id ? pool.filter(function(x){ return x.id == id; })[0] : null, ctx = menuCtx_();
      if (id && !cur) fail_(M.label + ' no encontrado');
      var rec = M.clean(d, cur, ctx);
      if (cur && cur.is_active && !rec.is_active && M.guard(cur, 'off')) fail_(M.guard(cur, 'off'));
      if (M.img) {
        if (img) rec[M.img] = img;
        else if (d.image_clear) rec[M.img] = '';
        else if (!id && d.copy_image_id && pool.some(function(x){ return x[M.img] === d.copy_image_id; })) rec[M.img] = d.copy_image_id;   // duplicate keeps the photo
      }
      if (M.check) M.check(rec);
      if (rec.sort_order === -1) rec.sort_order = nextOrder_(M, pool, rec);
      uniqCheck_(M, rec, id, pool);
      if (!id) rec[M.code[0]] = nextCode_(M, pool);
      var saved = id ? JDB.update(M.sheet, id, rec) : JDB.insert(M.sheet, rec, me.u);
      if (key === 'addon' && cur && cur.is_active && !saved.is_active) dropAddon_([saved.id]);
      bustFront_(); _m.usage = null;
      addLog_(me.u, M.log + (id ? '_EDITADO' : '_AGREGADO'), saved[M.code[0]] + ' · ' + M.nm(saved));
      return ok_({ message:M.label + (id ? ' actualizado' : ' agregado'), row:menuOut_(key, saved, menuCtx_()) });
    });
  });
}
// an archived add-on leaves every drink it was on (max add-ons clamped to what is left) — ONE write per touched day row
function dropAddon_(ids) {
  var gone = {}; ids.forEach(function(x){ gone[x] = 1; });
  var hit = JDB.all(PROD_SHEET).filter(function(p){ return (p.addon_ids || []).some(function(a){ return gone[a]; }); }).map(function(p){ return p.id; });
  if (hit.length) JDB.mutate(PROD_SHEET, hit, function(p){ p.addon_ids = p.addon_ids.filter(function(a){ return !gone[a]; }); p.max_addons = Math.min(p.max_addons, p.addon_ids.length); });
  return hit.length;
}
// archive / restore many — skips (and names) rows the guard protects; never deletes
function setMenuStatus_(tok, key, ids, on) {
  var M = MENU[key];
  return api_(tok, M.page, 'e', function(me){
    if (!ids || !ids.length) return err_('Nada seleccionado');
    return withLock_(function(){
      var want = {}, go = [], skipped = [], pool = JDB.all(M.sheet);
      ids.forEach(function(x){ want[String(x)] = 1; });
      pool.forEach(function(r){
        if (!want[String(r.id)] || r.is_active === (on ? 1 : 0)) return;
        var why = on ? '' : M.guard(r, 'off');
        if (!why && on) { try { uniqCheck_(M, { name:r.name, is_active:1 }, r.id, pool); } catch (e) { why = e.user; } }
        if (why) return skipped.push(M.nm(r) + ' (' + why + ')');
        go.push(r.id);
      });
      if (go.length) JDB.patchMany(M.sheet, go, { is_active:on ? 1 : 0 });
      var moved = key === 'addon' && !on && go.length ? dropAddon_(go) : 0;
      bustFront_(); _m.usage = null;
      if (go.length) addLog_(me.u, M.log + (on ? '_RESTABLECIDO' : '_ARCHIVADO'), M.label + ': ' + go.length + (moved ? ' · removido de ' + moved + ' bebida(s)' : ''));
      return ok_({ count:go.length, ids:go, skipped:skipped.length, skippedNames:skipped, touched:moved });
    });
  });
}
// skip-and-collect: bad rows -> errors[], good rows still land; dedup against the sheet AND the batch; ONE write
function importMenu_(tok, key, rows) {
  var M = MENU[key];
  return api_(tok, M.page, 'a', function(me){
    if (!rows || !rows.length) return err_('No hay filas para importar');
    if (rows.length > 1000) return err_('Importa como máximo 1000 filas a la vez');
    return withLock_(function(){
      var pool = JDB.all(M.sheet), ctx = menuCtx_(), out = [], errors = [];
      rows.forEach(function(raw, i){
        try {
          var d = {}; M.csv.forEach(function(c){ if (raw.hasOwnProperty(c[0])) d[c[1]] = raw[c[0]]; });
          var rec = M.clean(d, null, ctx);
          if (M.check) M.check(rec);
          rec.sort_order = nextOrder_(M, pool.concat(out), rec);
          uniqCheck_(M, rec, 0, pool.concat(out));
          rec[M.code[0]] = nextCode_(M, pool.concat(out));
          out.push(rec);
        } catch (e) { errors.push('Fila ' + (i + 2) + ': ' + ((e && (e.user || e.message)) || e)); }
      });
      JDB.insertMany(M.sheet, out, me.u);
      bustFront_(); _m.usage = null;
      addLog_(me.u, M.log + '_IMPORTACION', 'Importación de ' + M.label.toLowerCase() + ': ' + out.length + ' agregados, ' + errors.length + ' omitidos');
      return ok_({ count:out.length, errors:errors });
    });
  });
}
// ONE admin payload for Products + Categories + Add-ons (and every admin picker)
function getMenuAdmin(tok) {
  return api_(tok, null, null, function(me){
    if (!['products', 'categories', 'addons'].some(function(p){ return hasPerm_(me.role, p, 'v'); })) return err_('Acceso denegado');
    var ctx = menuCtx_(), byOrder = function(a, b){ return (a.sort_order || 0) - (b.sort_order || 0) || String(a.name).localeCompare(String(b.name)); };
    var list = function(key){ return JDB.all(MENU[key].sheet).sort(byOrder).map(function(r){ return menuOut_(key, r, ctx); }); };
    var cats = list('category'), catOrder = {}; cats.forEach(function(c, i){ catOrder[c.id] = i; });
    return ok_({ categories:cats, addons:list('addon'),
                 products:list('product').sort(function(a, b){ return (catOrder[a.category_id] - catOrder[b.category_id]) || byOrder(a, b); }),
                 icons:CAT_ICONS, csv:{ category:MENU.category.csv.map(function(c){ return c[0]; }), addon:MENU.addon.csv.map(function(c){ return c[0]; }), product:MENU.product.csv.map(function(c){ return c[0]; }) } });
  });
}
// sold out / available · featured / not — many rows, one write
function setMenuFlag(tok, key, ids, flag, on) {
  var M = MENU[key];
  if (['product', 'addon'].indexOf(key) === -1 || ['is_available', 'is_featured'].indexOf(flag) === -1 || (flag === 'is_featured' && key !== 'product')) return api_(tok, 'products', 'e', function(){ return err_('No permitido'); });   // gate first, then the argument check
  return api_(tok, M.page, 'e', function(me){
    if (!ids || !ids.length) return err_('Nada seleccionado');
    return withLock_(function(){
      var patch = {}; patch[flag] = on ? 1 : 0;
      var hit = JDB.patchMany(M.sheet, ids, patch);
      bustFront_(); _m.usage = null;
      addLog_(me.u, M.log + '_EDITADO', M.label + ': ' + hit.length + ' → ' + (flag === 'is_available' ? (on ? 'disponible' : 'agotado') : (on ? 'destacado' : 'no destacado')));
      return ok_({ count:hit.length, ids:hit.map(function(x){ return x.id; }), message:hit.length + ' ' + M.label.toLowerCase() + (hit.length === 1 ? '' : 's') + ' actualizado(s)' });
    });
  });
}
// move many drinks to another category (they go last there)
function moveProducts(tok, ids, catId) {
  return api_(tok, 'products', 'e', function(me){
    if (!ids || !ids.length) return err_('Nada seleccionado');
    return withLock_(function(){
      var ctx = menuCtx_(), cat = ctx.cats[Number(catId)];
      if (!cat || !cat.is_active) fail_('Selecciona una categoría activa');
      var pool = JDB.all(PROD_SHEET), n = nextOrder_(MENU.product, pool, { category_id:cat.id });
      var hit = JDB.mutate(PROD_SHEET, ids, function(p){ if (p.category_id === cat.id) return false; p.category_id = cat.id; p.sort_order = n++; });
      bustFront_(); _m.usage = null;
      addLog_(me.u, 'PRODUCTO_EDITADO', hit.length + ' bebida(s) movida(s) a ' + cat.name);
      return ok_({ count:hit.length, ids:hit.map(function(x){ return x.id; }), category_id:cat.id, category_name:cat.name, message:hit.length + ' bebida(s) movida(s) a ' + cat.name });
    });
  });
}
// order: swap with the neighbour above / below (products within their category) -> every row's new order back
function moveMenuItem(tok, key, id, dir) {
  var M = MENU[key];
  if (!M) return api_(tok, 'products', 'e', function(){ return err_('Lista desconocida'); });
  return api_(tok, M.page, 'e', function(me){
    return withLock_(function(){
      var all = JDB.all(M.sheet), me0 = all.filter(function(x){ return x.id == id; })[0];
      if (!me0) fail_(M.label + ' no encontrado');
      var list = all.filter(function(x){ return !M.order || x[M.order] === me0[M.order]; })
        .sort(function(a, b){ return (a.sort_order || 0) - (b.sort_order || 0) || String(a.name).localeCompare(String(b.name)); });
      var i = list.indexOf(me0), j = i + (dir < 0 ? -1 : 1);
      if (j < 0 || j >= list.length) return ok_({ message:'Ya está en el ' + (dir < 0 ? 'principio' : 'final'), order:{} });
      list.splice(j, 0, list.splice(i, 1)[0]);
      var want = {}; list.forEach(function(x, k){ want[x.id] = k + 1; });
      JDB.mutate(M.sheet, list.map(function(x){ return x.id; }), function(x){ if (x.sort_order === want[x.id]) return false; x.sort_order = want[x.id]; });
      bustFront_();
      return ok_({ message:'Orden guardado', order:want });
    });
  });
}

// thin public wrappers (the browser calls these by name)
function saveCategory(tok, d)            { return saveMenu_(tok, 'category', d); }
function setCategoryStatus(tok, ids, on) { return setMenuStatus_(tok, 'category', ids, on); }
function importCategories(tok, rows)     { return importMenu_(tok, 'category', rows); }
function saveAddon(tok, d)               { return saveMenu_(tok, 'addon', d); }
function setAddonStatus(tok, ids, on)    { return setMenuStatus_(tok, 'addon', ids, on); }
function importAddons(tok, rows)         { return importMenu_(tok, 'addon', rows); }
function saveProduct(tok, d)             { return saveMenu_(tok, 'product', d); }
function setProductStatus(tok, ids, on)  { return setMenuStatus_(tok, 'product', ids, on); }
function importProducts(tok, rows)       { return importMenu_(tok, 'product', rows); }

// ============== Storefront menu payload (cached with the storefront) ==============
// active categories · active drinks of active categories (sold-out ones stay, flagged) · active add-ons · active payment methods
function menuPublic_() {
  var cats = jdbList_(CAT_SHEET).filter(function(c){ return c.is_active; }).sort(function(a, b){ return a.sort_order - b.sort_order; }), live = {};
  cats.forEach(function(c){ live[c.id] = 1; });
  var addons = jdbList_(ADDON_SHEET).filter(function(a){ return a.is_active; }).sort(function(a, b){ return a.sort_order - b.sort_order; }), ad = {};
  addons.forEach(function(a){ ad[a.id] = a; });
  var prods = jdbList_(PROD_SHEET).filter(function(p){ return p.is_active && live[p.category_id]; })
    .sort(function(a, b){ return a.sort_order - b.sort_order || String(a.name).localeCompare(String(b.name)); })
    .map(function(p){ var ids = (p.addon_ids || []).filter(function(x){ return ad[x]; });
      return { id:p.id, code:p.product_code, name:p.name, category_id:p.category_id, description:p.description || '', image_url:imgUrl_(p.image_id),
               base_price:p.base_price, from_price:fromPrice_(p), sizes:p.sizes || [], has_sugar:p.has_sugar, has_ice:p.has_ice,
               addon_ids:ids, max_addons:Math.min(p.max_addons, ids.length), is_featured:p.is_featured, is_available:p.is_available }; });
  return { categories:cats.map(function(c){ return { id:c.id, name:c.name, icon:c.icon, tagline:c.tagline || '' }; }), products:prods,
           addons:addons.map(function(a){ return { id:a.id, name:a.name, price:a.price, is_available:a.is_available }; }), methods:paymentsPublic_() };
}

// demo photos: free Unsplash images linked by URL (an upload from the admin replaces them with a Drive file)
var demoPhoto_ = function(id, w, h){ return 'https://images.unsplash.com/photo-' + id + '?w=' + w + '&h=' + h + '&fit=crop&q=75&auto=format'; };
// demo menu — generic drink names, round prices; one drink archived, one drink + one add-on sold out
function seedMenu_() {
  var at = function(n){ return daysAgo_(60 - n, '09:00'); };
  jdbSeed_(CAT_SHEET, [['Té con Leche', 'fa-mug-hot', 'Cremoso. Delicioso. Hecho al momento.'], ['Soda de Frutas', 'fa-lemon', 'Fresca, burbujeante y bien fría.'], ['Jugo Natural', 'fa-glass-water', 'Recién exprimido, 100% natural.'], ['Té Helado', 'fa-leaf', 'Preparado a diario, servido con hielo.']]
    .map(function(c, i){ return { cat_code:'CAT-0' + (i + 1), name:c[0], icon:c[1], tagline:c[2], sort_order:i + 1, is_active:1, created:at(i) }; }), 'admin');
  // cols: name, price, available
  jdbSeed_(ADDON_SHEET, [['Leche extra', 3000], ['Extra dulce', 1500], ['Perlas de tapioca', 2500], ['Jalea de coco', 2500, 0], ['Espuma de queso', 3500]]
    .map(function(a, i){ return { addon_code:'ADD-0' + (i + 1), name:a[0], price:a[1], is_available:a[2] === 0 ? 0 : 1, sort_order:i + 1, is_active:1, created:at(5) }; }), 'admin');
  var ML = function(l){ return [{ name:'Mediano', price_delta:0 }, { name:'Grande', price_delta:l }]; };
  // cols: name, category, base, sizes, sugar, ice, add-ons, max, featured, available, active, description, demo photo
  var P = [
    ['Té con Leche Clásico', 1, 12000, ML(2500), 1, 1, [1, 2, 3], 3, 1, 1, 1, 'Té negro con leche cremosa.', '1558857563-b371033873b8'],
    ['Té con Leche de Taro', 1, 14000, ML(2500), 1, 1, [1, 2, 3, 5], 3, 1, 1, 1, 'Té con leche y suave taro.', '1525803377221-4f6ccdaa5133'],
    ['Leche con Azúcar Morena', 1, 15000, ML(3000), 0, 1, [3, 5], 2, 0, 1, 1, 'Leche fresca con sirope de azúcar morena.', '1572490122747-3968b75cc699'],
    ['Soda de Limón', 2, 9000, [], 1, 1, [2], 1, 1, 1, 1, 'Soda con limón fresco.', '1621263764928-df1444c5e859'],
    ['Fizz de Mango', 2, 13000, ML(2500), 1, 1, [2, 4], 2, 0, 1, 1, 'Puré de mango con soda.', '1546173159-315724a31696'],
    ['Soda de Fresa', 2, 12000, ML(2500), 1, 1, [2, 4], 2, 0, 1, 1, 'Sirope de fresa, soda y hielo.', '1497534446932-c925b458314e'],
    ['Jugo de Naranja', 3, 9000, [], 0, 1, [], 0, 0, 0, 1, 'Naranjas recién exprimidas.', '1600271886742-f049cd451bba'],
    ['Jugo de Sandía', 3, 9000, [], 0, 1, [], 0, 0, 1, 1, 'Sandía fresca, licuada al momento.', '1595981267035-7b04ca84a82d'],
    ['Té Verde de Maracuyá', 4, 11000, ML(2000), 1, 1, [2, 3, 4], 3, 1, 1, 1, 'Té verde con pulpa de maracuyá.', '1609951651556-5334e2706168'],
    ['Té Helado con Limón', 4, 10000, ML(2000), 1, 1, [2], 1, 0, 1, 1, 'Té negro con hielo y limón.', '1556679343-c7306c1976bc'],
    ['Oolong de Durazno', 4, 12000, ML(2000), 1, 1, [2], 1, 0, 1, 0, 'De temporada — regresa el próximo verano.', '1499638673689-79a0b5115d87']
  ], ord = {};
  jdbSeed_(PROD_SHEET, P.map(function(p, i){ ord[p[1]] = (ord[p[1]] || 0) + 1;
    return { product_code:'DRK-' + ('00' + (i + 1)).slice(-3), name:p[0], category_id:p[1], description:p[11], image_id:demoPhoto_(p[12], 640, 640), base_price:p[2], sizes:p[3],
             has_sugar:p[4], has_ice:p[5], addon_ids:p[6], max_addons:p[7], is_featured:p[8], is_available:p[9], sort_order:ord[p[1]], is_active:p[10], created:at(6 + i) }; }), 'admin');
  bustFront_();
}

// ============== Métodos de pago — imagen QR O nombre del banco + titular + número de cuenta ==============
var PM_ICONS = ['fa-qrcode', 'fa-building-columns', 'fa-wallet', 'fa-mobile-screen', 'fa-money-bill-transfer', 'fa-money-bill-wave', 'fa-credit-card'];
function cleanPayMethod_(d, cur) {
  var name = str_(d, 'name', 40); if (!name) fail_('Ingresa el nombre del método de pago');
  var icon = str_(d, 'icon', 40); if (PM_ICONS.indexOf(icon) === -1) icon = (cur && cur.icon) || 'fa-building-columns';
  var acct = str_(d, 'account_number', 40).replace(/\s+/g, '');
  if (acct && !/^[0-9A-Za-z-]{4,40}$/.test(acct)) fail_('Número de cuenta: solo dígitos, letras y "-"');
  return { name:name, icon:icon, qr_image_id:cur ? cur.qr_image_id || '' : '', bank_name:str_(d, 'bank_name', 80), account_name:str_(d, 'account_name', 80),
           account_number:acct, instructions:str_(d, 'instructions', 200), sort_order:cur ? cur.sort_order : -1, is_active:yes_(d.is_active, 1) };
}
// the client's display rule: a method needs a QR image, or all three bank fields
var payMethodOk_ = function(r){ if (!r.qr_image_id && !(r.bank_name && r.account_name && r.account_number)) fail_('Agrega una imagen QR o completa los tres campos bancarios (nombre del banco, titular de la cuenta, número de cuenta)'); };
MENU.paymethod = { sheet:PM_SHEET, page:'paymethods', label:'Método de pago', clean:cleanPayMethod_, code:['pm_code', 'PM-', 2], nm:function(r){ return r.name; },
  uniq:[['name', 'Ya existe un método de pago llamado']], guard:function(){ return ''; }, log:'METODO_PAGO', img:'qr_image_id', check:payMethodOk_,
  csv:[['Name', 'name'], ['Icon', 'icon'], ['Bank Name', 'bank_name'], ['Account Name', 'account_name'], ['Account Number', 'account_number'], ['Instructions', 'instructions'], ['Active', 'is_active']] };
MENU.product.img = 'image_id';

// admin list: every method + orders / approved amount in the last 30 days
function getPaymentMethods(tok) {
  return api_(tok, 'paymethods', 'v', function(){
    var from = addDays_(todayYmd_(), -29), n = {}, amt = {};
    if (sh_(ORDER_SHEET)) JDB.between(ORDER_SHEET, from, null).forEach(function(o){ n[o.payment_method_id] = (n[o.payment_method_id] || 0) + 1;
      if (PAID_ST[o.status]) amt[o.payment_method_id] = round_((amt[o.payment_method_id] || 0) + o.grand_total); });
    var rows = JDB.all(PM_SHEET).sort(function(a, b){ return a.sort_order - b.sort_order; }).map(function(r){ var o = strip_(r);
      o.qr_url = imgUrl_(r.qr_image_id); o.orders_30 = n[r.id] || 0; o.approved_30 = amt[r.id] || 0; return o; });
    return ok_({ rows:rows, icons:PM_ICONS, csv:MENU.paymethod.csv.map(function(c){ return c[0]; }) });
  });
}
function savePaymentMethod(tok, d)           { return saveMenu_(tok, 'paymethod', d); }
function setPaymentMethodStatus(tok, ids, on) { return setMenuStatus_(tok, 'paymethod', ids, on); }
function importPaymentMethods(tok, rows)      { return importMenu_(tok, 'paymethod', rows); }
// shoppers see active methods only, in the admin's order
function paymentsPublic_() {
  return jdbList_(PM_SHEET).filter(function(m){ return m.is_active; }).sort(function(a, b){ return a.sort_order - b.sort_order; }).map(function(m){
    return { id:m.id, name:m.name, icon:m.icon, qr_url:imgUrl_(m.qr_image_id), bank_name:m.bank_name, account_name:m.account_name, account_number:m.account_number, instructions:m.instructions };
  });
}

// ============== Pedidos — placeOrder (público), registro de pedido, índice de pagos ==============
var ORDER_ST = ['SUBMITTED', 'PAYMENT_REJECTED', 'CONFIRMED', 'PREPARING', 'READY', 'COMPLETED', 'CANCELLED'];
var OPEN_ST = { SUBMITTED:1, PAYMENT_REJECTED:1, CONFIRMED:1, PREPARING:1, READY:1 };
var CUST_LABEL = { SUBMITTED:'Pago en revisión', PAYMENT_REJECTED:'Problema con el pago — vuelve a subir el comprobante', CONFIRMED:'Pedido confirmado', PREPARING:'En preparación',
  READY:'Listo para recoger', COMPLETED:'Completado', CANCELLED:'Cancelado' };
var MAX_LINES = 30, QTY_MAX = 20;
var RCP_TYPES = { 'image/jpeg':'jpg', 'image/png':'png', 'image/webp':'webp', 'image/heic':'heic', 'image/heif':'heic', 'application/pdf':'pdf' };

// the day keys that still hold open orders (the board / review screens read only those rows)
var openDays_ = function(){ return safeParse_(PropertiesService.getScriptProperties().getProperty('OPEN_DAYS'), []) || []; };
var markOpenDay_ = function(day){ var a = openDays_(); if (a.indexOf(day) === -1) { a.push(day); a.sort(); PropertiesService.getScriptProperties().setProperty('OPEN_DAYS', JSON.stringify(a)); } };

// ONE server price rule — rebuilt from the cached public menu, never from the browser's numbers
function buildLines_(items, F, s) {
  if (!Array.isArray(items) || !items.length) fail_('Tu carrito está vacío');
  if (items.length > MAX_LINES) fail_('Hasta ' + MAX_LINES + ' líneas por pedido');
  var P = {}, A = {}, dp = Number(s.currency_decimals) === 0 ? 0 : 2, r = function(v){ var f = Math.pow(10, dp); return Math.round(v * f + 1e-9) / f; };
  F.products.forEach(function(p){ P[p.id] = p; }); F.addons.forEach(function(a){ A[a.id] = a; });
  var lines = items.map(function(it, i){
    it = it || {};
    var p = P[Number(it.product_id)], n = 'Línea ' + (i + 1) + ': ';
    if (!p) fail_(n + 'una bebida ya no está en el menú');
    if (!p.is_available) fail_(n + p.name + ' está agotado');
    var qty = Number(it.qty); if (!(qty >= 1 && qty <= QTY_MAX) || Math.round(qty) !== qty) fail_(n + 'la cantidad debe ser de 1 a ' + QTY_MAX);
    var sz = null;
    if (p.sizes.length) { sz = p.sizes.filter(function(z){ return z.name === it.size; })[0]; if (!sz) fail_(n + 'elige un tamaño para ' + p.name); }
    var sugar = p.has_sugar ? String(it.sugar || '') : '', ice = p.has_ice ? String(it.ice || '') : '';
    if (p.has_sugar && s.sugar_levels.indexOf(sugar) === -1) fail_(n + 'elige un nivel de azúcar para ' + p.name);
    if (p.has_ice && s.ice_levels.indexOf(ice) === -1) fail_(n + 'elige un nivel de hielo para ' + p.name);
    var ids = Array.isArray(it.addon_ids) ? it.addon_ids.map(Number) : [], seen = {};
    if (ids.length > p.max_addons) fail_(n + p.name + ' admite hasta ' + p.max_addons + ' complemento' + (p.max_addons === 1 ? '' : 's'));
    var ads = ids.map(function(id){ var a = A[id];
      if (seen[id]) fail_(n + 'un complemento está listado dos veces'); seen[id] = 1;
      if (!a || p.addon_ids.indexOf(id) === -1) fail_(n + 'un complemento ya no se ofrece en ' + p.name);
      if (!a.is_available) fail_(n + a.name + ' está agotado');
      return { addon_id:a.id, name:a.name, price:a.price }; });
    var addonsTotal = r(ads.reduce(function(t, a){ return t + a.price; }, 0)), unit = r(p.base_price + (sz ? sz.price_delta : 0) + addonsTotal);
    return { line_no:i + 1, product_id:p.id, product_name:p.name, size_name:sz ? sz.name : '', size_delta:sz ? sz.price_delta : 0, sugar:sugar, ice:ice,
             addons:ads, note:String(it.note || '').trim().slice(0, 100), base_price:p.base_price, addons_total:addonsTotal, unit_price:unit, qty:qty, line_total:r(unit * qty) };
  });
  return { lines:lines, subtotal:r(lines.reduce(function(t, l){ return t + l.line_total; }, 0)), count:lines.reduce(function(t, l){ return t + l.qty; }, 0) };
}
// pickup slots for today: every slot_minutes from max(now + prep, opening) to the last-order cut-off
function slots_(s, st) {
  if (!st.open) return [];
  var step = Number(s.slot_minutes) || 15, now = hmMin_(st.now_hm) + (Number(s.prep_minutes) || 15), lo = Math.max(now, hmMin_(st.today.open)), hi = hmMin_(st.today.cutoff), out = [];
  for (var m = Math.ceil(lo / step) * step; m <= hi; m += step) out.push(minHm_(m));
  return out;
}
// receipt: type by the file's own header (magic bytes), size by the Settings limit
function receiptBlob_(f, s) {
  if (!f || !f.b64) fail_('Sube tu comprobante de pago');
  var raw = String(f.b64), b64 = raw.split(',').pop(), max = Number(s.max_receipt_mb) || 5;
  if (b64.length * 0.75 > max * 1048576) fail_('El comprobante es mayor a ' + max + ' MB');
  var bytes = Utilities.base64Decode(b64), h = function(i){ return bytes[i] & 255; }, str = function(a, z){ var o = ''; for (var i = a; i < z && i < bytes.length; i++) o += String.fromCharCode(h(i)); return o; };
  if (!bytes || bytes.length < 12) fail_('Ese archivo está vacío');
  var mime = h(0) === 0xFF && h(1) === 0xD8 && h(2) === 0xFF ? 'image/jpeg' : str(0, 8) === '\x89PNG\r\n\x1a\n' ? 'image/png' : str(0, 4) === 'RIFF' && str(8, 12) === 'WEBP' ? 'image/webp'
    : str(0, 4) === '%PDF' ? 'application/pdf' : str(4, 8) === 'ftyp' && /^(heic|heix|heif|mif1|msf1|hevc)$/.test(str(8, 12)) ? 'image/heic' : '';
  if (!mime) fail_('El comprobante debe ser una foto (JPG, PNG, WEBP, HEIC) o un PDF');
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes).map(function(b){ return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
  return { bytes:bytes, mime:mime, ext:RCP_TYPES[mime], hash:hash, kb:Math.round(bytes.length / 1024) };
}
// private Drive: <Receipts root>/<YYYY>/<MM>/<name> — never link-shared; the folder ids are remembered
function storeReceipt_(rb, name) {
  var day = todayYmd_(), y = subFolder_(receiptsRoot_(), day.slice(0, 4), 'DRV_RCP_' + day.slice(0, 4)), m = subFolder_(y, day.slice(5, 7), 'DRV_RCP_' + day.slice(0, 7));
  return m.createFile(Utilities.newBlob(rb.bytes, rb.mime, name + '.' + rb.ext));
}
// payment index — receipt hash + txn ref -> order no, last 90 days; built once, patched on every write
var PAYIX_KEY = 'payix_v1';
function payIx_() {
  if (_m.payix) return _m.payix;
  var hit = safeParse_(CACHE_.get(PAYIX_KEY), null);
  if (hit) return (_m.payix = hit);
  var ix = { h:{}, t:{} };
  if (sh_(ORDER_SHEET)) JDB.between(ORDER_SHEET, addDays_(todayYmd_(), -90), null).forEach(function(o){ (o.payment_history || []).forEach(function(p){
    if (p.hash) ix.h[p.hash] = o.order_no; if (p.txn_ref) ix.t[p.txn_ref] = o.order_no; }); });
  cachePut_(PAYIX_KEY, ix, CACHE_TTL);
  return (_m.payix = ix);
}
var payIxAdd_ = function(hash, txn, no){ var ix = payIx_(); if (hash) ix.h[hash] = no; if (txn) ix.t[txn] = no; cachePut_(PAYIX_KEY, ix, CACHE_TTL); };
var dupOf_ = function(map, key, no){ var o = key ? map[key] : ''; return o && o !== no ? o : ''; };
var cleanTxn_ = function(v){ var t = String(v == null ? '' : v).trim().toUpperCase().replace(/\s+/g, ''); if (t && !/^[0-9A-Z._\/-]{4,40}$/.test(t)) fail_('ID de transacción: 4-40 letras o dígitos'); return t; };

// what a shopper sees about their own order — no usernames, no Drive ids, no admin note
function orderPublic_(o, withToken) {
  var s = settings_(), tries = Number(s.reupload_max) || 3;
  return { order_no:o.order_no, short_no:o.short_no, placed_at:o.placed_at, status:o.status, status_label:o.status === 'READY' && o.fulfilment === 'DELIVERY' ? 'En camino para entrega' : CUST_LABEL[o.status],
    customer_name:o.customer_name, customer_phone:o.customer_phone, fulfilment:o.fulfilment, delivery_address:o.delivery_address, pickup_mode:o.pickup_mode, pickup_at:o.pickup_at,
    order_note:o.order_note, items:o.items, item_count:o.item_count, subtotal:o.subtotal, delivery_fee:o.delivery_fee, grand_total:o.grand_total,
    payment:{ name:o.payment_snapshot.name, txn_ref:o.txn_ref, attempt:o.payment_attempt, attempts_left:Math.max(0, tries - o.payment_attempt) }, reject_reason:o.reject_reason || '',
    eta_at:o.eta_at || '', paid_at:o.paid_at || '', prep_started_at:o.prep_started_at || '', ready_at:o.ready_at || '', completed_at:o.completed_at || '', cancelled_at:o.cancelled_at || '', updated:o.updated,
    timeline:(o.timeline || []).map(function(t){ return { at:t.at, action:t.action, to:t.to }; }), track_token:withToken ? o.track_token : undefined };
}
// tracking token = the order's day (yyyymmdd) + 24 random hex — a link finds its order by reading ONE day row
var tokenHex_ = function(day){ return String(day).replace(/-/g, '') + (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 24); };
var CREF_RX = /^c-[0-9a-z-]{6,40}$/i;
var phoneKey_ = function(p){ return String(p || '').replace(/\D/g, ''); };

// the ONE public way in: rebuild the order from the menu, check everything, store the receipt privately, write once under the lock
function placeOrder(tok, d) {
  return pub_(function(){
    d = d || {};
    if (d.hp) fail_('No se pudo realizar el pedido');                                   // honeypot: a person never fills it
    var cref = String(d.client_ref || ''); if (!CREF_RX.test(cref)) fail_('Por favor, actualiza la página e intenta de nuevo');
    var done = CACHE_.get('cref_' + cref);                                            // idempotent: a retry gets the order it already made
    if (done) { var o0 = findOrder_(done); if (o0) return ok_({ order:orderPublic_(o0, true), repeat:1 }); }
    var me = whoMaybe_(tok), s = settings_(), st = openState_(s);
    if (!st.open) fail_(st.reason === 'paused' ? 'Los pedidos en línea están pausados en este momento — por favor intenta pronto' : 'Los pedidos en línea están cerrados en este momento');
    var F = front_(), b = buildLines_(d.items, F, s);
    // customer
    var acct = null;
    if (me) { var u = userRows_(), ui = uRow_(u, me.u); acct = ui === -1 ? null : u.data[ui]; }
    var name = str_(d, 'name', 80), phone = str_(d, 'phone', 20), email = acct ? String(acct[U.EMAIL]) : str_(d, 'email', 100).toLowerCase();
    if (name.length < 2) fail_('Ingresa tu nombre');
    if (!PHONE_RX.test(phone)) fail_('Ingresa un número de teléfono válido');
    if (email && !EMAIL_RX.test(email)) fail_('Ingresa un correo electrónico válido o déjalo vacío');
    // fulfilment + time
    var ful = String(d.fulfilment || 'PICKUP').toUpperCase();
    if (ful !== 'PICKUP' && ful !== 'DELIVERY') fail_('Elige retiro o entrega');
    if (ful === 'PICKUP' && !Number(s.pickup_enabled)) fail_('El retiro en tienda no está disponible');
    if (ful === 'DELIVERY' && !Number(s.delivery_enabled)) fail_('La entrega a domicilio no está disponible');
    var addr = str_(d, 'delivery_address', 300); if (ful === 'DELIVERY' && addr.length < 5) fail_('Ingresa la dirección de entrega');
    var mode = d.pickup_mode === 'SCHEDULED' ? 'SCHEDULED' : 'ASAP', at = '';
    if (mode === 'SCHEDULED') { var hm = String(d.pickup_time || ''); if (slots_(s, st).indexOf(hm) === -1) fail_('Ese horario ya no está disponible — elige otro'); at = st.today.date + 'T' + hm; }
    // money: the server's total wins; a changed price stops the order before anything is stored
    var fee = ful === 'DELIVERY' ? round_(s.delivery_fee) : 0, total = round_(b.subtotal + fee);
    if (Number(s.min_order_amount) && b.subtotal < Number(s.min_order_amount)) fail_('El pedido mínimo es ' + s.currency_symbol + round_(s.min_order_amount));
    if (d.expect_total != null && Math.abs(Number(d.expect_total) - total) > 0.001) return { success:false, code:'PRICES', message:'Los precios fueron actualizados — por favor revisa tu pedido', total:total };
    // payment
    var pm = jdbList_(PM_SHEET).filter(function(m){ return m.id === Number(d.payment_method_id) && m.is_active; })[0];
    if (!pm) fail_('Elige un método de pago');
    var txn = cleanTxn_(d.txn_ref), rb = receiptBlob_(d.receipt, s);
    // abuse limits (per phone + the whole shop) — only counted for orders that got this far
    if (throttled_('op_' + phoneKey_(phone), 3, 600)) fail_('Demasiados pedidos desde este número de teléfono — por favor espera unos minutos');
    if (throttled_('op_all', 60, 600)) fail_('Estamos muy ocupados — por favor intenta de nuevo en unos minutos');
    var file = storeReceipt_(rb, 'receipt_' + cref);                                 // Drive write outside the lock
    var order = withLock_(function(){
      var again = CACHE_.get('cref_' + cref); if (again) { var o1 = findOrder_(again); if (o1) return o1; }
      var no = nextNo_('ORD', ORDER_SHEET, 'order_no'), now = nowIso_(), ix = payIx_(), who = me ? me.u : 'guest';
      var rec = { order_no:no, short_no:'#' + no.slice(-3), track_token:tokenHex_(ymd_(now)), client_ref:cref, placed_at:now, checkout_type:me ? 'ACCOUNT' : 'GUEST',
        customer_username:me ? me.u : '', customer_name:name, customer_phone:phone, customer_email:email, fulfilment:ful, delivery_address:ful === 'DELIVERY' ? addr : '',
        pickup_mode:mode, pickup_at:at, order_note:str_(d, 'order_note', 200), items:b.lines, item_count:b.count, subtotal:b.subtotal, delivery_fee:fee, grand_total:total,
        payment_method_id:pm.id, payment_snapshot:{ name:pm.name, bank_name:pm.bank_name, account_name:pm.account_name, account_last4:String(pm.account_number || '').slice(-4), had_qr:pm.qr_image_id ? 1 : 0 },
        receipt_file_id:file.getId(), receipt_hash:rb.hash, receipt_mime:rb.mime, txn_ref:txn, payment_attempt:1,
        payment_history:[{ n:1, file_id:file.getId(), hash:rb.hash, mime:rb.mime, txn_ref:txn, at:now, result:'', reason:'', by:who }],
        dup_receipt_order:dupOf_(ix.h, rb.hash, no), dup_txn_order:dupOf_(ix.t, txn, no), status:'SUBMITTED', verified_by:'', paid_at:'', reject_reason:'', eta_at:'',
        prep_started_at:'', ready_at:'', completed_at:'', cancelled_at:'', cancelled_by:'', cancel_reason:'', refund_status:'NONE', refund_note:'', admin_note:'',
        timeline:[{ at:now, by:who, action:'Pedido realizado', from:'', to:'SUBMITTED', note:pm.name + (txn ? ' · ' + txn : '') }] };
      var saved = JDB.insert(ORDER_SHEET, rec, who);
      payIxAdd_(rb.hash, txn, no); markOpenDay_(ymd_(now));
      CACHE_.put('cref_' + cref, no, 86400);
      return saved;
    });
    try { file.setName(order.order_no + '_1.' + rb.ext); } catch (e) {}
    addLog_(me ? me.u : 'guest:' + phoneKey_(phone).slice(-4), 'PEDIDO_REALIZADO', order.order_no + ' · ' + s.currency_symbol + order.grand_total + ' · ' + pm.name + (order.dup_receipt_order || order.dup_txn_order ? ' · MARCADO' : ''));
    notifyShop_(order, s);
    notifyCustomer_(order, 'received', s);
    return ok_({ order:orderPublic_(order, true) });
  });
}
// one order by its number — the day row comes from the number itself (ORD-YYYYMMDD-NNNN)
function findOrder_(no) {
  var m = /^ORD-(\d{4})(\d{2})(\d{2})-\d{4,}$/.exec(String(no || '')); if (!m || !sh_(ORDER_SHEET)) return null;
  var day = m[1] + '-' + m[2] + '-' + m[3];
  return JDB.between(ORDER_SHEET, day, day).filter(function(o){ return o.order_no === no; })[0] || null;
}

// demo payment methods: a generic sample QR (placeholder image, not a real wallet) + a bank transfer + an archived old account
var DEMO_QR_B64 = 'iVBORw0KGgoAAAANSUhEUgAAAOgAAADoCAAAAADAwvekAAAB4UlEQVR42u3byRGDMAwAQPpvOqkgg67YHOtvAmj9QRLy8XnJOkBBQUFBQUFBQUFBQUEvDT2KK3v9r/9PxwMKCgoKCroDGn4RB6FZ2FQ8oKCgoKCgO6HRwKsB/GsjQUFBQUFB7wztJg6goKCgoKBvhE4lBqCgoKCgoE+Adgvj6ofd23XqQUFBQUFBP3MDTGeBZAvvy02OgYKCgoKCLjjgU4WMxwEKCgoKCroAOlV4n4Gq4OyHY1BQUFBQ0JXQaGDhByQb1dVG9unvoKCgoKCgC6HdgLuFeHXDQUFBQUFBd0L/1UDuNrTbH6BBQUFBQUE3QrMHcLIFfTYRSQ9mgYKCgoKCLoB2X+zZQKoDyunrQUFBQUFBF0KzCUEW0j24kwWDgoKCgoKugHYHhLsbM9XoBgUFBQUF3QktN4qHCvJsPO1OPSgoKCgo6AC0u7qDT9GCPZ1ggIKCgoKCLoB2G8tTGzN1f1BQUFBQ0B3QbKDZwrvauC433EFBQUFBQTdAq4FXG9PV+4CCgoKCgj4JGg04OrjVTkRAQUFBQUFvCM02qM+eH90gUFBQUFDQK0CriUA0McgCoteDgoKCgoLuhFYHqo7mqm5UuYENCgoKCgo6CH3qAgUFBQUFBQUFBQUFBb3U+gKyeQJMD3hPLQAAAABJRU5ErkJggg==';
function seedPayments_() {
  var qr = shopFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(DEMO_QR_B64), 'image/png', 'demo_qr.png')).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW).getId();
  jdbSeed_(PM_SHEET, [
    { pm_code:'PM-01', name:'Billetera QR (Nequi / Daviplata)', icon:'fa-qrcode', qr_image_id:qr, bank_name:'', account_name:'', account_number:'', instructions:'Escanea el código QR desde tu app de Nequi o Daviplata y adjunta el comprobante.', sort_order:1, is_active:1, created:daysAgo_(40, '09:00') },
    { pm_code:'PM-02', name:'Transferencia Bancolombia', icon:'fa-building-columns', qr_image_id:'', bank_name:'Bancolombia', account_name:'Bebidas Demo SAS', account_number:'10234567890', instructions:'Cuenta de Ahorros. Coloca tu número de pedido en la descripción de la transferencia.', sort_order:2, is_active:1, created:daysAgo_(40, '09:05') },
    { pm_code:'PM-03', name:'Cuenta Davivienda', icon:'fa-building-columns', qr_image_id:'', bank_name:'Davivienda', account_name:'Bebidas Demo SAS', account_number:'009988776655', instructions:'', sort_order:3, is_active:0, created:daysAgo_(40, '09:10') }
  ], 'admin');
  bustFront_();
}

// ============== Seguimiento (público) · volver a subir comprobante · cancelar cliente · Mis Pedidos ==============
// the order a caller may see: a live token, or order no + the exact phone, or (signed in) their own order
function ownOrder_(tok, q) {
  q = q || {};
  var me = whoMaybe_(tok), o = null, t = String(q.token || '').trim();
  if (t) {
    if (!/^\d{8}[0-9a-f]{24}$/.test(t)) fail_('Pedido no encontrado');
    if (throttled_('tr_' + t.slice(8, 16), 20, 60)) fail_('Demasiadas consultas — espera un minuto');
    var day = t.slice(0, 4) + '-' + t.slice(4, 6) + '-' + t.slice(6, 8);
    o = sh_(ORDER_SHEET) ? JDB.between(ORDER_SHEET, day, day).filter(function(x){ return x.track_token === t; })[0] : null;
  } else {
    var no = String(q.order_no || '').trim().toUpperCase(), ph = phoneKey_(q.phone);
    if (!no) fail_('Ingresa tu número de pedido');
    if (!me && ph.length < 7) fail_('Ingresa el número de teléfono utilizado en el pedido');
    if (throttled_('tr_' + (ph || me.u), 5, 60)) fail_('Demasiadas consultas — espera un minuto');
    o = findOrder_(no);
    if (o && !(ph && phoneKey_(o.customer_phone) === ph) && !(me && o.customer_username === me.u)) o = null;   // same answer either way
  }
  if (!o) fail_('Pedido no encontrado — verifica el número de pedido y teléfono');
  return o;
}
// order no alone = progress only (numbers are guessable); token, phone or the customer's own account = the full order
function trackOrder(tok, q) {
  return pub_(function(){
    q = q || {};
    if (q.token || String(q.phone || '').trim()) return ok_({ order:orderPublic_(ownOrder_(tok, q), true) });
    var me = whoMaybe_(tok), no = String(q.order_no || '').trim().toUpperCase();
    if (!no) fail_('Ingresa tu número de pedido');
    if (!/^ORD-\d{8}-\d{4,6}$/.test(no)) fail_('Pedido no encontrado — verifica el número de pedido');
    if (throttled_('trn_' + no, 20, 60) || throttled_('trn_all', 120, 60)) fail_('Demasiadas consultas — espera un minuto');
    var o = findOrder_(no);
    if (!o) fail_('Pedido no encontrado — verifica el número de pedido');
    return ok_({ order:me && o.customer_username === me.u ? orderPublic_(o, true) : orderStatus_(o) });
  });
}
// status-only copy: no name, phone, address, notes, txn, reject reason, timeline or token — and nothing to act on
function orderStatus_(o) {
  var x = orderPublic_(o, false);
  ['customer_name', 'customer_phone', 'delivery_address', 'order_note', 'reject_reason', 'timeline', 'track_token'].forEach(function(k){ delete x[k]; });
  x.payment = { name:x.payment.name };
  x.items = (x.items || []).map(function(l){ var c = JSON.parse(JSON.stringify(l)); delete c.note; return c; });
  x.limited = true;
  return x;
}
// a small thumbnail of the owner's own receipt (never the Drive id / url)
function myReceiptThumb(tok, q) {
  return pub_(function(){
    var o = ownOrder_(tok, q), f = DriveApp.getFileById(o.receipt_file_id), b = null;
    try { b = f.getThumbnail(); } catch (e) { b = null; }
    if (!b) return ok_({ url:'', mime:o.receipt_mime });
    return ok_({ url:'data:' + (b.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(b.getBytes()), mime:o.receipt_mime });
  });
}
// rejected payment -> a new receipt sends the order back to review (up to Settings reupload_max attempts)
function reuploadReceipt(tok, q, receipt, txnRef) {
  return pub_(function(){
    var o = ownOrder_(tok, q), s = settings_(), max = Number(s.reupload_max) || 3;
    if (o.status !== 'PAYMENT_REJECTED') fail_('Este pedido no está esperando un nuevo comprobante');
    if (o.payment_attempt >= max) fail_('No quedan intentos de comprobante — por favor contacta a la tienda');
    var txn = cleanTxn_(txnRef), rb = receiptBlob_(receipt, s), n = o.payment_attempt + 1, who = (whoMaybe_(tok) || {}).u || 'guest';
    var file = storeReceipt_(rb, o.order_no + '_' + n);
    var saved = withLock_(function(){
      var ix = payIx_(), now = nowIso_();
      var hit = JDB.mutate(ORDER_SHEET, [o.id], function(x){
        if (x.status !== 'PAYMENT_REJECTED' || x.payment_attempt >= max) return false;
        x.payment_attempt = n; x.receipt_file_id = file.getId(); x.receipt_hash = rb.hash; x.receipt_mime = rb.mime; x.txn_ref = txn;
        x.payment_history = (x.payment_history || []).concat({ n:n, file_id:file.getId(), hash:rb.hash, mime:rb.mime, txn_ref:txn, at:now, result:'', reason:'', by:who });
        x.dup_receipt_order = dupOf_(ix.h, rb.hash, x.order_no); x.dup_txn_order = dupOf_(ix.t, txn, x.order_no);
        x.status = 'SUBMITTED'; x.reject_reason = '';
        x.timeline = (x.timeline || []).concat({ at:now, by:who, action:'Comprobante vuelto a subir', from:'PAYMENT_REJECTED', to:'SUBMITTED', note:'intento ' + n + (txn ? ' · ' + txn : '') });
      }, ymd_(o.placed_at), ymd_(o.placed_at))[0];
      if (!hit) fail_('Este pedido cambió hace un momento — actualiza e intenta de nuevo');
      payIxAdd_(rb.hash, txn, hit.order_no); markOpenDay_(ymd_(hit.placed_at));
      return hit;
    });
    addLog_(who === 'guest' ? 'guest:' + phoneKey_(o.customer_phone).slice(-4) : who, 'COMPROBANTE_REPETIDO', saved.order_no + ' · intento ' + n);
    notifyShop_(saved, s, 'reupload');
    return ok_({ message:'Comprobante enviado — lo revisaremos en breve', order:orderPublic_(saved, true) });
  });
}
// the customer may cancel before the payment is approved
function cancelMyOrder(tok, q, reason) {
  return pub_(function(){
    var o = ownOrder_(tok, q), who = (whoMaybe_(tok) || {}).u || 'guest';
    if (!{ SUBMITTED:1, PAYMENT_REJECTED:1 }[o.status]) fail_('Este pedido ya no puede cancelarse aquí — por favor contacta a la tienda');
    var saved = withLock_(function(){
      var now = nowIso_(), hit = JDB.mutate(ORDER_SHEET, [o.id], function(x){
        if (!{ SUBMITTED:1, PAYMENT_REJECTED:1 }[x.status]) return false;
        x.timeline = (x.timeline || []).concat({ at:now, by:who, action:'Cancelado por el cliente', from:x.status, to:'CANCELLED', note:String(reason || '').trim().slice(0, 200) });
        x.status = 'CANCELLED'; x.cancelled_at = now; x.cancelled_by = 'customer'; x.cancel_reason = String(reason || '').trim().slice(0, 200) || 'Cancelado por el cliente';
      }, ymd_(o.placed_at), ymd_(o.placed_at))[0];
      if (!hit) fail_('Este pedido cambió hace un momento — actualiza e intenta de nuevo');
      return hit;
    });
    addLog_(who === 'guest' ? 'guest:' + phoneKey_(o.customer_phone).slice(-4) : who, 'PEDIDO_CANCELADO', saved.order_no + ' · por el cliente');
    notifyShop_(saved, settings_(), 'cust_cancel');
    return ok_({ message:'Pedido cancelado', order:orderPublic_(saved, true) });
  });
}
// a signed-in customer's own orders, newest first (last 180 days)
function getMyOrders(tok) {
  return api_(tok, null, null, function(me){
    var list = sh_(ORDER_SHEET) ? JDB.between(ORDER_SHEET, addDays_(todayYmd_(), -180), null).filter(function(o){ return o.customer_username === me.u; }) : [];
    return ok_({ orders:list.sort(function(a, b){ return a.placed_at < b.placed_at ? 1 : -1; }).map(function(o){ return orderPublic_(o, true); }) });
  });
}

// ============== Admin orders — open feed · register · 360 · receipts · every status move (G9) ==============
// what the admin screens get: the record minus Drive ids / tracking secrets (receipts come through getReceiptImage)
function orderAdmin_(o) {
  var x = strip_(o);
  delete x.receipt_file_id; delete x.track_token; delete x.client_ref;
  x.payment_history = (o.payment_history || []).map(function(p){ return { n:p.n, mime:p.mime, txn_ref:p.txn_ref, at:p.at, result:p.result, reason:p.reason, by:p.by }; });
  return x;
}
// open work = open statuses + a paid-and-cancelled order still waiting for its refund
var isOpenWork_ = function(o){ return !!OPEN_ST[o.status] || (o.status === 'CANCELLED' && o.refund_status === 'PENDING'); };
// every open order across the open days; a day with nothing open left leaves the set (lazy prune)
function openScan_() {
  var days = openDays_(), rows = [], still = {};
  days.forEach(function(d){ (sh_(ORDER_SHEET) ? JDB.between(ORDER_SHEET, d, d) : []).forEach(function(o){ if (isOpenWork_(o)) { rows.push(o); still[d] = 1; } }); });
  var keep = days.filter(function(d){ return still[d]; });
  if (keep.length !== days.length) PropertiesService.getScriptProperties().setProperty('OPEN_DAYS', JSON.stringify(keep));
  return { rows:rows, dropped:days.length - keep.length };
}
function getOpenOrders(tok) {
  return api_(tok, 'orders', 'v', function(){
    var out = openScan_().rows.map(orderAdmin_), s = settings_();
    return ok_({ rows:out.sort(function(a, b){ return a.placed_at < b.placed_at ? -1 : 1; }), now:nowIso_(), ordering_enabled:Number(s.ordering_enabled) ? 1 : 0, prep_minutes:s.prep_minutes });
  });
}
// register: every order placed in the range + every order still open (so nothing open hides behind a date)
function getOrders(tok, from, to) {
  return api_(tok, 'orders', 'v', function(){
    from = ymd_(from) || todayYmd_(); to = ymd_(to) || from;
    if (to < from) { var t = from; from = to; to = t; }
    var seen = {}, rows = [];
    var add = function(o){ if (seen[o.order_no]) return; seen[o.order_no] = 1; rows.push(orderAdmin_(o)); };
    if (sh_(ORDER_SHEET)) { JDB.between(ORDER_SHEET, from, to).forEach(add); openDays_().forEach(function(d){ JDB.between(ORDER_SHEET, d, d).filter(isOpenWork_).forEach(add); }); }
    return ok_({ rows:rows.sort(function(a, b){ return a.placed_at < b.placed_at ? 1 : -1; }), from:from, to:to, methods:jdbList_(PM_SHEET).map(function(m){ return { id:m.id, name:m.name }; }) });
  });
}
function getOrder360(tok, no) {
  return api_(tok, 'orders', 'v', function(){ var o = findOrder_(no); if (!o) fail_('Pedido no encontrado'); return ok_({ order:orderAdmin_(o) }); });
}
// one receipt (attempt n, default the latest) as a data url — only through the gate, never a Drive link
function getReceiptImage(tok, no, n) {
  return api_(tok, 'orders', 'v', function(){
    var o = findOrder_(no); if (!o) fail_('Pedido no encontrado');
    var h = (o.payment_history || []).filter(function(p){ return p.n === (Number(n) || o.payment_attempt); })[0];
    if (!h) fail_('Comprobante no encontrado');
    var b = DriveApp.getFileById(h.file_id).getBlob();
    return ok_({ url:'data:' + (h.mime || b.getContentType()) + ';base64,' + Utilities.base64Encode(b.getBytes()), mime:h.mime, n:h.n });
  });
}

// ---- transitions: lock -> re-read -> check the status it must be in -> write ONE record -> timeline + log ----
var ST_LABEL = { SUBMITTED:'Revisión de Pago', PAYMENT_REJECTED:'Pago Rechazado', CONFIRMED:'Confirmado', PREPARING:'En Preparación', READY:'Listo', COMPLETED:'Completado', CANCELLED:'Cancelado' };
function moveOrder_(me, no, from, fn) {
  return withLock_(function(){
    var o = findOrder_(no); if (!o) fail_('Pedido no encontrado');
    if (from.indexOf(o.status) === -1) {
      var last = (o.timeline || []).slice(-1)[0] || {};
      fail_('El pedido ' + o.short_no + ' ya está ' + ST_LABEL[o.status] + (last.by ? ' (por ' + last.by + ')' : '') + ' — actualiza');
    }
    var now = nowIso_(), was = o.status;
    var hit = JDB.mutate(ORDER_SHEET, [o.id], function(x){ var r = fn(x, now); x.timeline = (x.timeline || []).concat({ at:now, by:me.u, action:r.action, from:was, to:x.status, note:r.note || '' }); },
      ymd_(o.placed_at), ymd_(o.placed_at))[0];
    if (isOpenWork_(hit)) markOpenDay_(ymd_(hit.placed_at));
    addLog_(me.u, hit.status === 'CANCELLED' && was !== 'CANCELLED' ? 'PEDIDO_CANCELADO' : 'ESTADO_PEDIDO', hit.order_no + ' · ' + ST_LABEL[was] + ' → ' + ST_LABEL[hit.status]);
    return hit;
  });
}
// ETA = approve time + prep minutes, or the scheduled pickup when that is later
function etaOf_(o, now, s) {
  var eta = new Date(Date.parse(now) + (Number(s.prep_minutes) || 15) * 60000).toISOString();
  if (o.pickup_mode === 'SCHEDULED' && o.pickup_at) { var p = localToIso_(o.pickup_at); if (p > eta) eta = p; }
  return eta;
}
// 'yyyy-MM-ddTHH:mm' on the shop's clock -> ISO (the offset is read from the script time zone)
var localToIso_ = function(l){ var guess = new Date(l + ':00Z'), off = Utilities.formatDate(guess, tz_(), 'Z'), m = /^([+-])(\d{2})(\d{2})$/.exec(off || '+0000');
  var mins = m ? (m[1] === '-' ? -1 : 1) * (Number(m[2]) * 60 + Number(m[3])) : 0; return new Date(guess.getTime() - mins * 60000).toISOString(); };
function approvePayment(tok, no, checked) {
  return api_(tok, 'orders', 'e', function(me){
    if (checked !== true) return err_('Marca "He verificado mi aplicación bancaria / billetera" primero');
    var s = settings_(), o = moveOrder_(me, no, ['SUBMITTED'], function(x, now){
      x.status = 'CONFIRMED'; x.verified_by = me.u; x.paid_at = now; x.eta_at = etaOf_(x, now, s);
      var h = x.payment_history[x.payment_history.length - 1]; if (h) { h.result = 'APPROVED'; h.by = me.u; }
      return { action:'Pago aprobado', note:x.payment_snapshot.name + ' · ' + s.currency_symbol + x.grand_total };
    });
    notifyCustomer_(o, 'approved', s);
    return ok_({ message:'Pago aprobado — ' + o.short_no + ' está confirmado', order:orderAdmin_(o) });
  });
}
var REJECT_WHY = ["El monto no coincide", 'Pago no recibido', 'Comprobante ilegible', 'Comprobante ya utilizado', 'Otro'];
function rejectPayment(tok, no, reason) {
  return api_(tok, 'orders', 'e', function(me){
    var why = String(reason || '').trim().slice(0, 200); if (why.length < 3) return err_('Elige o escribe el motivo');
    var s = settings_(), o = moveOrder_(me, no, ['SUBMITTED'], function(x){
      x.status = 'PAYMENT_REJECTED'; x.reject_reason = why;
      var h = x.payment_history[x.payment_history.length - 1]; if (h) { h.result = 'REJECTED'; h.reason = why; h.by = me.u; }
      return { action:'Pago rechazado', note:why };
    });
    notifyCustomer_(o, 'rejected', s, { left:Math.max(0, (Number(s.reupload_max) || 3) - o.payment_attempt) });
    return ok_({ message:'Pago rechazado — el cliente puede enviar un nuevo comprobante', order:orderAdmin_(o) });
  });
}
// the kitchen moves: Confirmed -> Preparing -> Ready -> Completed (Picked up / Delivered)
var NEXT_ST = { CONFIRMED:['PREPARING', 'prep_started_at', 'Preparación iniciada'], PREPARING:['READY', 'ready_at', 'Marcado como listo'], READY:['COMPLETED', 'completed_at', 'Entregado'] };
function stepOrder_(me, no, to, s) {
  var from = Object.keys(NEXT_ST).filter(function(k){ return NEXT_ST[k][0] === to; })[0];
  if (!from) fail_('Paso no válido');
  var o = moveOrder_(me, no, [from], function(x, now){ var n = NEXT_ST[x.status]; x.status = n[0]; x[n[1]] = now;
    return { action:n[0] === 'COMPLETED' ? (x.fulfilment === 'DELIVERY' ? 'Entregado a domicilio' : 'Retirado') : n[2] }; });
  if (to === 'READY') notifyCustomer_(o, 'ready', s);
  if (to === 'COMPLETED') notifyCustomer_(o, 'completed', s);
  return o;
}
function setOrderStatus(tok, no, to) {
  return api_(tok, 'orders', 'e', function(me){ var o = stepOrder_(me, no, to, settings_());
    return ok_({ message:o.short_no + ' → ' + ST_LABEL[o.status], order:orderAdmin_(o) }); });
}
// bulk: Mark ready / Complete only (payments are never approved in bulk) — skip-and-collect
function bulkOrderStatus(tok, nos, to) {
  return api_(tok, 'orders', 'e', function(me){
    if (['READY', 'COMPLETED'].indexOf(to) === -1) return err_('Solo "Marcar como listo" y "Completar" funcionan en lote');
    if (!nos || !nos.length) return err_('Nada seleccionado');
    var s = settings_(), done = [], skipped = [];
    nos.slice(0, 100).forEach(function(no){ try { done.push(orderAdmin_(stepOrder_(me, no, to, s))); } catch (e) { skipped.push(no + ': ' + ((e && e.user) || e)); } });
    return ok_({ count:done.length, orders:done, skipped:skipped.length, skippedNames:skipped });
  });
}
// one step back (mis-tap): Preparing -> Confirmed, Ready -> Preparing, Completed -> Ready
var BACK_ST = { PREPARING:['CONFIRMED', 'prep_started_at'], READY:['PREPARING', 'ready_at'], COMPLETED:['READY', 'completed_at'] };
function undoOrderStep(tok, no, reason) {
  return api_(tok, 'orders', 'e', function(me){
    var why = String(reason || '').trim().slice(0, 200); if (why.length < 3) return err_('Indica el motivo para deshacer');
    var o = moveOrder_(me, no, Object.keys(BACK_ST), function(x){ var b = BACK_ST[x.status]; x.status = b[0]; x[b[1]] = ''; return { action:'Deshecho', note:why }; });
    return ok_({ message:o.short_no + ' regresó a ' + ST_LABEL[o.status], order:orderAdmin_(o) });
  });
}
// cancel (never delete). A paid order (Confirmed / Preparing) is marked "refund pending"
function cancelOrder(tok, no, reason) {
  return api_(tok, 'orders', 'd', function(me){
    var why = String(reason || '').trim().slice(0, 200); if (why.length < 3) return err_('Indica el motivo de cancelación');
    var o = moveOrder_(me, no, ['SUBMITTED', 'PAYMENT_REJECTED', 'CONFIRMED', 'PREPARING'], function(x, now){
      var paid = x.status === 'CONFIRMED' || x.status === 'PREPARING';
      x.status = 'CANCELLED'; x.cancelled_at = now; x.cancelled_by = me.u; x.cancel_reason = why; if (paid) x.refund_status = 'PENDING';
      return { action:'Cancelado', note:why + (paid ? ' · reembolso pendiente' : '') };
    });
    notifyCustomer_(o, 'cancelled', settings_());
    return ok_({ message:o.short_no + ' cancelado' + (o.refund_status === 'PENDING' ? ' — reembolso pendiente' : ''), order:orderAdmin_(o) });
  });
}
function markRefunded(tok, no, note) {
  return api_(tok, 'orders', 'e', function(me){
    var o = withLock_(function(){ var x0 = findOrder_(no); if (!x0) fail_('Pedido no encontrado'); if (x0.refund_status !== 'PENDING') fail_('No hay reembolso pendiente para este pedido');
      var now = nowIso_();
      return JDB.mutate(ORDER_SHEET, [x0.id], function(x){ x.refund_status = 'REFUNDED'; x.refund_note = String(note || '').trim().slice(0, 200);
        x.timeline = (x.timeline || []).concat({ at:now, by:me.u, action:'Reembolso registrado', from:x.status, to:x.status, note:x.refund_note }); }, ymd_(x0.placed_at), ymd_(x0.placed_at))[0]; });
    addLog_(me.u, 'REEMBOLSO_REGISTRADO', o.order_no + (o.refund_note ? ' · ' + o.refund_note : ''));
    notifyCustomer_(o, 'refunded', settings_());
    return ok_({ message:'Reembolso registrado', order:orderAdmin_(o) });
  });
}
function saveAdminNote(tok, no, note) {
  return api_(tok, 'orders', 'e', function(me){
    var o = withLock_(function(){ var x0 = findOrder_(no); if (!x0) fail_('Pedido no encontrado');
      return JDB.update(ORDER_SHEET, x0.id, { admin_note:String(note || '').trim().slice(0, 300) }, ymd_(x0.placed_at)); });
    return ok_({ message:'Nota guardada', order:orderAdmin_(o) });
  });
}
// ============== Email notifications — one branded template, a switch per event (Settings → Notifications) ==============
// shop.* go to the notification emails, customer.* to the order's email; the two master switches gate each side
var MAIL_EVENTS = ['shop.new_order', 'shop.reupload', 'shop.cust_cancel', 'shop.signup', 'shop.daily',
  'customer.received', 'customer.approved', 'customer.rejected', 'customer.ready', 'customer.completed', 'customer.cancelled', 'customer.refunded', 'customer.welcome'];
var MAIL_TONE = { ok:'#34a853', warn:'#fbbc04', bad:'#ea4335', info:'#0074D9' };   // the theme's success / warning / danger / accent
var mailOn_ = function(s, ev){ return !!Number(ev.indexOf('shop.') === 0 ? s.notify_shop : s.notify_customer) && (',' + s.notify_events + ',').indexOf(',' + ev + ',') !== -1; };
// "a@x.com; b@y.com" -> "a@x.com, b@y.com" — up to 5, each checked
var mailList_ = function(v){ var a = String(v || '').split(/[\s,;]+/).filter(String);
  if (a.length > 5) fail_('Hasta 5 correos de notificación'); a.forEach(function(e){ if (!EMAIL_RX.test(e)) fail_('Correo no válido: ' + e); }); return a.join(', '); };
var mailMoney_ = function(s){
  var dp = Number(s.currency_decimals) === 2 ? 2 : 0;
  return function(v){
    var n = Number(v || 0);
    var parts = n.toFixed(dp).split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    var sym = s.currency_symbol || '$';
    return sym + (sym.endsWith(' ') ? '' : ' ') + (dp > 0 ? parts.join(',') : parts[0]);
  };
};
var pickupOf_ = function(o){ return (o.fulfilment === 'DELIVERY' ? 'Entrega' : 'Retiro') + ' · ' + (o.pickup_mode === 'SCHEDULED' ? String(o.pickup_at).slice(11, 16) : 'Inmediato'); };
var appUrl_ = function(){ try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } };

// the one template: header · tone bar · title + intro · order table or key/values · button · shop footer (inline styles — mail clients drop <style>)
function mailHtml_(s, m) {
  var e = escHtml_, M = mailMoney_(s), o = m.order, td = 'padding:8px 0;border-bottom:1px solid #eef1f5;vertical-align:top';
  var lines = o ? (o.items || []).map(function(l){
    var opt = [l.size_name, l.sugar, l.ice].filter(function(x){ return x; }).concat((l.addons || []).map(function(a){ return '+ ' + a.name; })).join(' · ');
    return '<tr><td style="' + td + '"><b>' + l.qty + '× ' + e(l.product_name) + '</b>' + (opt ? '<br><span style="color:#6b7280;font-size:12px">' + e(opt) + '</span>' : '') +
      '</td><td style="' + td + ';text-align:right;white-space:nowrap">' + M(l.line_total) + '</td></tr>'; }).join('') +
    '<tr><td style="padding:6px 0;color:#6b7280">Subtotal</td><td style="padding:6px 0;text-align:right">' + M(o.subtotal) + '</td></tr>' +
    (Number(o.delivery_fee) ? '<tr><td style="padding:6px 0;color:#6b7280">Entrega</td><td style="padding:6px 0;text-align:right">' + M(o.delivery_fee) + '</td></tr>' : '') +
    '<tr><td style="padding:8px 0;font-size:16px;font-weight:700">Total</td><td style="padding:8px 0;font-size:16px;font-weight:700;text-align:right">' + M(o.grand_total) + '</td></tr>' : '';
  var kv = (m.kv || []).map(function(r){ return '<tr><td style="' + td + ';color:#6b7280">' + e(r[0]) + '</td><td style="' + td + ';text-align:right;font-weight:600">' + e(r[1]) + '</td></tr>'; }).join('');
  var foot = [s.shop_address, s.shop_phone, s.shop_email].filter(function(x){ return String(x || '').trim(); }).map(e).join(' · ');
  return '<div style="background:#f5f5f5;padding:24px 12px;font-family:Arial,Helvetica,sans-serif;color:#333333">' +
    '<div style="max-width:560px;margin:0 auto;background:#ffffff;border:1px solid #e6e9ef;border-radius:6px;overflow:hidden">' +
    '<div style="background:#001f3f;color:#ffffff;padding:16px 22px;font-size:18px;font-weight:700">' +
      (s.shop_logo ? '<img src="' + e(imgUrl_(s.shop_logo)) + '" alt="" width="40" height="40" style="border-radius:50%;vertical-align:middle;margin-right:10px;background:#ffffff">' : '') + e(s.shop_name) + '</div>' +
    '<div style="height:4px;background:' + (MAIL_TONE[m.tone] || MAIL_TONE.info) + '"></div><div style="padding:22px">' +
    '<h1 style="margin:0 0 8px;font-size:20px;line-height:1.3;color:#001f3f">' + e(m.title) + '</h1>' +
    '<p style="margin:0;font-size:14px;line-height:1.6">' + m.intro + '</p>' +            // intro arrives escaped
    (o ? '<p style="margin:14px 0 0;color:#6b7280;font-size:13px">Pedido <b style="color:#001f3f">' + e(o.short_no) + '</b> · ' + e(o.order_no) + '<br>' + e(pickupOf_(o)) +
      ((o.payment_snapshot || {}).name ? ' · ' + e(o.payment_snapshot.name) : '') + '</p>' : '') +
    (lines || kv ? '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:12px;font-size:14px;border-collapse:collapse">' + lines + kv + '</table>' : '') +
    (m.button && m.button[1] ? '<p style="margin:22px 0 4px"><a href="' + e(m.button[1]) + '" style="display:inline-block;padding:12px 26px;background:#001f3f;color:#ffffff;text-decoration:none;border-radius:4px;font-weight:700">' + e(m.button[0]) + '</a></p>' : '') +
    '</div><div style="padding:14px 22px;background:#f8fafc;border-top:1px solid #eef1f5;color:#6b7280;font-size:12px;line-height:1.5">' + (foot ? foot + '<br>' : '') + e(m.why) + '</div></div></div>';
}

// event -> { to, subject, html }; send and preview share it. x = extras (welcome / signup: name, email, phone · daily: day · rejected: left)
function mailOf_(ev, o, s, x) {
  x = x || {};
  var e = escHtml_, M = mailMoney_(s), app = appUrl_(), track = o && o.track_token && app ? app + '?track=' + o.track_token : app;
  var hi = function(n){ return e(String(n || '').trim().split(/\s+/)[0] || 'Hola'); }, review = app && app + '?page=review';
  var flag = o && (o.dup_receipt_order || o.dup_txn_order) ? ' <b style="color:#ea4335">Aviso de pago duplicado — revísalo con atención.</b>' : '';
  var cust = function(subj, m){ return { to:EMAIL_RX.test(String(o.customer_email || '')) ? o.customer_email : '', subject:subj,
    m:Object.assign({ order:o, button:['Seguir tu pedido', track], why:'Recibes este correo porque hiciste un pedido en ' + s.shop_name + '.' }, m) }; };
  var shop = function(subj, m){ return { to:String(s.notify_shop_email || ''), subject:subj, m:Object.assign({ why:'Notificación de la tienda — cámbiala en Configuración → Notificaciones.' }, m) }; };
  var r = null;
  switch (ev) {
    case 'shop.new_order': r = shop('nuevo pedido ' + o.short_no + ' (' + M(o.grand_total) + ')', { title:'Nuevo pedido ' + o.short_no, tone:'info', order:o,
      intro:e(o.customer_name) + ' · ' + e(o.customer_phone) + ' — revisa el comprobante en Revisión de Pagos.' + flag, button:['Abrir Revisión de Pagos', review] }); break;
    case 'shop.reupload': r = shop('nuevo comprobante para ' + o.short_no, { title:'Nuevo comprobante para ' + o.short_no, tone:'warn', order:o,
      intro:e(o.customer_name) + ' envió el intento de comprobante ' + o.payment_attempt + (o.txn_ref ? ' (transacción ' + e(o.txn_ref) + ')' : '') + ' — el pedido está nuevamente en Revisión de Pagos.' + flag, button:['Abrir Revisión de Pagos', review] }); break;
    case 'shop.cust_cancel': r = shop(o.short_no + ' cancelado por el cliente', { title:o.short_no + ' fue cancelado por el cliente', tone:'bad', order:o,
      intro:e(o.customer_name) + (o.cancel_reason ? ' dio este motivo: ' + e(o.cancel_reason) : ' canceló antes de que se aprobara el pago') + '.' }); break;
    case 'shop.signup': r = shop('nuevo cliente ' + x.name, { title:'Nueva cuenta de cliente', tone:'ok', intro:e(x.name) + ' acaba de registrarse.', kv:[['Correo electrónico', x.email], ['Teléfono', x.phone || '—']] }); break;
    case 'shop.daily': r = shop('resumen diario ' + x.day, { title:'Resumen diario — ' + x.day, tone:'info', intro:'Ayer en ' + e(s.shop_name) + ', de un vistazo.', kv:dailyOf_(x.day, s), button:['Abrir el panel', app] }); break;
    case 'customer.received': r = cust('recibimos tu pedido ' + o.short_no, { title:'¡Gracias, ' + hi(o.customer_name) + '! — Recibimos tu pedido', tone:'info',
      intro:'Estamos revisando tu pago en este momento. Recibirás otro correo tan pronto sea confirmado.' }); break;
    case 'customer.approved': r = cust('tu pedido ' + o.short_no + ' está confirmado', { title:'Tu pedido ' + o.short_no + ' está confirmado', tone:'ok',
      intro:'Recibimos tu pago — tus bebidas se están preparando' + (o.eta_at ? ' y deberían estar listas alrededor de las <b>' + e(Utilities.formatDate(new Date(o.eta_at), tz_(), 'HH:mm')) + '</b>' : '') + '.' }); break;
    case 'customer.rejected': r = cust('no pudimos confirmar el pago para ' + o.short_no, { title:'No pudimos confirmar tu pago', tone:'bad', button:['Enviar un nuevo comprobante', track],
      intro:'Motivo: <b>' + e(o.reject_reason) + '</b>. Abre tu pedido y envía un nuevo comprobante' + (x.left != null ? ' — te queda' + (x.left === 1 ? ' ' : 'n ') + x.left + ' intento' + (x.left === 1 ? '' : 's') : '') + '.' }); break;
    case 'customer.ready': r = cust('tu pedido ' + o.short_no + ' está listo', { title:'Tu pedido ' + o.short_no + ' está listo', tone:'ok',
      intro:o.fulfilment === 'DELIVERY' ? 'Tu pedido está en camino.' : 'Ven a recogerlo — muestra el número de pedido en el mostrador.' }); break;
    case 'customer.completed': r = cust('gracias — comprobante para ' + o.short_no, { title:'¡Muchas gracias, ' + hi(o.customer_name) + '!', tone:'ok', button:['Pedir de nuevo', app && app + '?page=menu'],
      intro:'¡Disfruta tus bebidas! Aquí tienes tu comprobante — esperamos verte pronto.' }); break;
    case 'customer.cancelled': r = cust('el pedido ' + o.short_no + ' fue cancelado', { title:'Tu pedido ' + o.short_no + ' fue cancelado', tone:'bad',
      intro:(o.cancel_reason ? 'Motivo: <b>' + e(o.cancel_reason) + '</b>. ' : '') + (o.refund_status === 'PENDING' ? 'Tu pago será reembolsado — te enviaremos un correo tan pronto sea enviado.' : 'No se te ha cobrado nada por este pedido.') }); break;
    case 'customer.refunded': r = cust('tu reembolso para ' + o.short_no + ' fue enviado', { title:'Tu reembolso fue enviado', tone:'ok',
      intro:'Te reembolsamos <b>' + M(o.grand_total) + '</b> por el pedido ' + e(o.short_no) + '.' + (o.refund_note ? ' Nota: ' + e(o.refund_note) : '') }); break;
    case 'customer.welcome': r = { to:x.email, subject:'bienvenido a ' + s.shop_name, m:{ title:'¡Bienvenido, ' + hi(x.name) + '!', tone:'ok', button:['Pedir ahora', app && app + '?page=menu'],
      intro:'Tu cuenta está lista. Tus datos están guardados para compras más rápidas, y Mis Pedidos guarda todo tu historial en un solo lugar.', why:'Recibes este correo porque creaste una cuenta en ' + s.shop_name + '.' } }; break;
  }
  return r && { to:r.to, subject:s.shop_name + ' — ' + r.subject, html:mailHtml_(s, r.m) };
}
// one day in numbers for the daily summary
function dailyOf_(day, s) {
  var M = mailMoney_(s), L = sh_(ORDER_SHEET) ? JDB.between(ORDER_SHEET, day, day) : [], paid = L.filter(function(o){ return PAID_ST[o.status]; }), top = {}, drinks = 0;
  var sales = paid.reduce(function(t, o){ return t + Number(o.grand_total); }, 0);
  paid.forEach(function(o){ (o.items || []).forEach(function(l){ drinks += l.qty; top[l.product_name] = (top[l.product_name] || 0) + l.qty; }); });
  var best = Object.keys(top).sort(function(a, b){ return top[b] - top[a]; })[0];
  return [['Ventas', M(sales)], ['Pedidos pagados', String(paid.length)], ['Ticket promedio', M(paid.length ? sales / paid.length : 0)], ['Bebidas vendidas', String(drinks)],
    ['Más vendido', best ? best + ' · ' + top[best] : '—'], ['Cancelados', String(L.filter(function(o){ return o.status === 'CANCELLED'; }).length)]];
}
// send = switch on + a recipient + quota left; never throws, never blocks the action that fired it
function mailSend_(ev, o, s, x) {
  try {
    if (!mailOn_(s, ev)) return false;
    var m = mailOf_(ev, o, s, x); if (!m || !m.to) return false;
    if (MailApp.getRemainingDailyQuota() < 1) { Logger.log('mail quota used up: ' + ev); return false; }
    var opt = { to:m.to, subject:m.subject, htmlBody:m.html, name:String(s.shop_name || 'Tienda').slice(0, 60) };
    if (EMAIL_RX.test(String(s.shop_email || ''))) opt.replyTo = s.shop_email;
    MailApp.sendEmail(opt);
    return true;
  } catch (e) { Logger.log('mailSend_ ' + ev + ': ' + e); return false; }
}
function notifyShop_(o, s, kind) { return mailSend_('shop.' + (kind || 'new_order'), o, s); }
function notifyCustomer_(o, kind, s, x) { return mailSend_('customer.' + kind, o, s, x); }

// ============== Demo orders — ~4 a day over the last 60 days, every status on show today, one flagged, one refund pending ==============
function seedOrders_() {
  var F = menuPublic_(), s = settings_(), P = F.products.filter(function(p){ return p.is_available; }), A = {}, rnd = (function(){ var x = 7; return function(){ x = (x * 16807) % 2147483647; return (x - 1) / 2147483646; }; })();
  F.addons.forEach(function(a){ A[a.id] = a; });
  var rcpt = receiptsRoot_().createFile(Utilities.newBlob(Utilities.base64Decode(DEMO_QR_B64), 'image/png', 'sample_receipt.png')).getId();   // one generic sample image for every seeded receipt
  var methods = jdbList_(PM_SHEET).filter(function(m){ return m.is_active; }), out = [], seq = {}, open = {};
  var custs = [['customer1@demo.com', 'Customer 1', '03001000002'], ['customer2@demo.com', 'Customer 2', '03001000003'], ['', 'Customer 11', '03001000011'], ['', 'Customer 12', '03001000012'], ['', 'Customer 13', '03001000013'], ['', 'Customer 14', '03001000014']];
  var todayPlan = ['SUBMITTED', 'SUBMITTED', 'PAYMENT_REJECTED', 'CONFIRMED', 'PREPARING', 'PREPARING', 'READY', 'COMPLETED', 'COMPLETED'];
  for (var day = 59; day >= 0; day--) {
    var n = day === 0 ? todayPlan.length : 3 + Math.floor(rnd() * 3);
    for (var k = 0; k < n; k++) {
      var at = day === 0 ? new Date(Date.now() - (n - k) * 11 * 60000).toISOString() : daysAgo_(day, p2_(11 + Math.floor(rnd() * 9)) + ':' + p2_(Math.floor(rnd() * 59))),   // today's orders: the last ~1.5 h
          d = ymd_(at), no, c = custs[Math.floor(rnd() * custs.length)];
      seq[d] = (seq[d] || 0) + 1; no = 'ORD-' + d.replace(/-/g, '') + '-' + ('000' + seq[d]).slice(-4);
      var lines = [], nl = 1 + Math.floor(rnd() * 3);
      for (var i = 0; i < nl; i++) { var p = P[Math.floor(rnd() * P.length)], sz = p.sizes.length ? p.sizes[Math.floor(rnd() * p.sizes.length)] : null;
        var ads = p.addon_ids.filter(function(id){ return A[id] && A[id].is_available; }).slice(0, Math.floor(rnd() * (p.max_addons + 1))).map(function(id){ return { addon_id:id, name:A[id].name, price:A[id].price }; });
        var at2 = ads.reduce(function(t, a){ return t + a.price; }, 0), unit = round_(p.base_price + (sz ? sz.price_delta : 0) + at2), q = 1 + Math.floor(rnd() * 2);
        lines.push({ line_no:i + 1, product_id:p.id, product_name:p.name, size_name:sz ? sz.name : '', size_delta:sz ? sz.price_delta : 0, sugar:p.has_sugar ? s.sugar_levels[Math.floor(rnd() * s.sugar_levels.length)] : '',
          ice:p.has_ice ? s.ice_levels[Math.floor(rnd() * s.ice_levels.length)] : '', addons:ads, note:'', base_price:p.base_price, addons_total:round_(at2), unit_price:unit, qty:q, line_total:round_(unit * q) }); }
      var sub = round_(lines.reduce(function(t, l){ return t + l.line_total; }, 0)), st = day === 0 ? todayPlan[k] : rnd() < 0.08 ? 'CANCELLED' : 'COMPLETED', m = methods[Math.floor(rnd() * methods.length)];
      var t = function(min){ return new Date(Date.parse(at) + min * 60000).toISOString(); }, paid = st !== 'SUBMITTED' && st !== 'PAYMENT_REJECTED' && !(st === 'CANCELLED' && k % 2);
      var o = { order_no:no, short_no:'#' + no.slice(-3), track_token:d.replace(/-/g, '') + ('000000000000000000000000' + (seq[d] * 7919 + day * 104729).toString(16)).slice(-24), client_ref:'c-seed-' + no.toLowerCase(),
        placed_at:at, checkout_type:c[0] ? 'ACCOUNT' : 'GUEST', customer_username:c[0], customer_name:c[1], customer_phone:c[2], customer_email:c[0] || '', fulfilment:'PICKUP', delivery_address:'',
        pickup_mode:'ASAP', pickup_at:'', order_note:'', items:lines, item_count:lines.reduce(function(x, l){ return x + l.qty; }, 0), subtotal:sub, delivery_fee:0, grand_total:sub,
        payment_method_id:m.id, payment_snapshot:{ name:m.name, bank_name:m.bank_name, account_name:m.account_name, account_last4:String(m.account_number || '').slice(-4), had_qr:m.qr_image_id ? 1 : 0 },
        receipt_file_id:rcpt, receipt_hash:'seed' + no, receipt_mime:'image/png', txn_ref:'TXN' + (40000 + seq[d] * 13 + day * 101), payment_attempt:1,
        payment_history:[{ n:1, file_id:rcpt, hash:'seed' + no, mime:'image/png', txn_ref:'TXN' + (40000 + seq[d] * 13 + day * 101), at:at, result:paid ? 'APPROVED' : st === 'PAYMENT_REJECTED' ? 'REJECTED' : '', reason:st === 'PAYMENT_REJECTED' ? "El monto no coincide" : '', by:c[0] || 'guest' }],
        dup_receipt_order:'', dup_txn_order:'', status:st, verified_by:paid ? 'admin' : '', paid_at:paid ? t(4) : '', reject_reason:st === 'PAYMENT_REJECTED' ? "El monto no coincide" : '', eta_at:paid ? t(4 + Number(s.prep_minutes)) : '',
        prep_started_at:{ PREPARING:1, READY:1, COMPLETED:1 }[st] ? t(6) : '', ready_at:{ READY:1, COMPLETED:1 }[st] ? t(14) : '', completed_at:st === 'COMPLETED' ? t(20) : '',
        cancelled_at:st === 'CANCELLED' ? t(8) : '', cancelled_by:st === 'CANCELLED' ? 'admin' : '', cancel_reason:st === 'CANCELLED' ? 'El cliente solicitó cancelar' : '', refund_status:st === 'CANCELLED' && paid ? (day > 2 ? 'REFUNDED' : 'PENDING') : 'NONE',
        refund_note:st === 'CANCELLED' && paid && day > 2 ? 'Devuelto a la misma cuenta' : '', admin_note:'', created:at, created_by:c[0] || 'guest',
        timeline:[{ at:at, by:c[0] || 'guest', action:'Pedido realizado', from:'', to:'SUBMITTED', note:m.name }] };
      if (paid) o.timeline.push({ at:t(4), by:'admin', action:'Pago aprobado', from:'SUBMITTED', to:'CONFIRMED', note:'' });
      if (st === 'PAYMENT_REJECTED') o.timeline.push({ at:t(5), by:'admin', action:'Pago rechazado', from:'SUBMITTED', to:'PAYMENT_REJECTED', note:o.reject_reason });
      if (o.prep_started_at) o.timeline.push({ at:t(6), by:'admin', action:'Preparación iniciada', from:'CONFIRMED', to:'PREPARING', note:'' });
      if (o.ready_at) o.timeline.push({ at:t(14), by:'admin', action:'Marcado como listo', from:'PREPARING', to:'READY', note:'' });
      if (o.completed_at) o.timeline.push({ at:t(20), by:'admin', action:'Retirado', from:'READY', to:'COMPLETED', note:'' });
      if (st === 'CANCELLED') o.timeline.push({ at:t(8), by:'admin', action:'Cancelado', from:paid ? 'CONFIRMED' : 'SUBMITTED', to:'CANCELLED', note:o.cancel_reason + (paid ? ' · reembolso pendiente' : '') });
      if (o.refund_status === 'REFUNDED') o.timeline.push({ at:t(90), by:'admin', action:'Reembolso registrado', from:'CANCELLED', to:'CANCELLED', note:o.refund_note });
      out.push(o);
      if (isOpenWork_(o)) open[d] = 1;
    }
  }
  // two old orders from Customer 1's phone placed as a guest -> Customers merges them under the account
  out.filter(function(o){ return o.customer_username === 'customer1@demo.com' && o.status === 'COMPLETED'; }).slice(0, 2).forEach(function(o){
    o.checkout_type = 'GUEST'; o.customer_username = ''; o.customer_email = ''; o.customer_phone = '0300 1000002'; o.created_by = 'guest'; o.timeline[0].by = 'guest'; });
  // today's second payment re-uses the first one's transaction ID -> the review screen shows the duplicate flag
  var subs = out.filter(function(o){ return o.status === 'SUBMITTED'; });
  if (subs.length > 1) { subs[1].txn_ref = subs[0].txn_ref; subs[1].payment_history[0].txn_ref = subs[0].txn_ref; subs[1].dup_txn_order = subs[0].order_no; }
  jdbSeed_(ORDER_SHEET, out, 'system');
  var p = PropertiesService.getScriptProperties(), td = todayYmd_();
  p.setProperty('OPEN_DAYS', JSON.stringify(Object.keys(open).sort()));
  if (seq[td]) p.setProperty('SEQ_ORD_' + td.replace(/-/g, ''), String(seq[td]));
  CACHE_.remove(PAYIX_KEY); _m.payix = null;
}
var p2_ = function(n){ return ('0' + n).slice(-2); };

// ============== Numbers — dashboard · reports · customers (one rule: sales = paid orders placed in the period) ==============
var isPaid_ = function(o){ return !!PAID_ST[o.status]; };
var ordersIn_ = function(from, to){ return sh_(ORDER_SHEET) ? JDB.between(ORDER_SHEET, from, to) : []; };
var hourOf_ = function(iso){ return Utilities.formatDate(new Date(iso), tz_(), 'HH:mm').slice(0, 2); };
var daysIn_ = function(from, to){ var out = []; for (var d = from; d <= to && out.length < 400; d = addDays_(d, 1)) out.push(d); return out; };
var digits_ = function(v){ return String(v || '').replace(/\D/g, ''); };
var inc_ = function(m, k, v){ m[k] = (m[k] || 0) + (v === undefined ? 1 : v); };
// sales · paid · average · drinks · add-on revenue · cancelled, over one list
function kpis_(list) {
  var k = { orders:list.length, paid:0, sales:0, drinks:0, addon_rev:0, cancelled:0, refund_pending:0 };
  list.forEach(function(o){
    if (o.status === 'CANCELLED') { k.cancelled++; if (o.refund_status === 'PENDING') k.refund_pending++; }
    if (!isPaid_(o)) return;
    k.paid++; k.sales += Number(o.grand_total) || 0;
    (o.items || []).forEach(function(l){ k.drinks += l.qty; (l.addons || []).forEach(function(a){ k.addon_rev += a.price * l.qty; }); });
  });
  k.sales = round_(k.sales); k.addon_rev = round_(k.addon_rev); k.avg = k.paid ? round_(k.sales / k.paid) : 0;
  return k;
}
// range key -> this period + the one it is compared with (Today = same weekday last week)
function dashRange_(key, from, to) {
  var t = todayYmd_(), n;
  if (key === 'today') return { from:t, to:t, pf:addDays_(t, -7), pt:addDays_(t, -7) };
  if (key === '7d' || key === '30d') { n = key === '7d' ? 7 : 30; return { from:addDays_(t, 1 - n), to:t, pf:addDays_(t, 1 - 2 * n), pt:addDays_(t, -n) }; }
  if (key === 'mtd') { var f = t.slice(0, 8) + '01', pm = addDays_(f, -1).slice(0, 8) + '01', last = addDays_(f, -1), same = pm.slice(0, 8) + t.slice(8);
    return { from:f, to:t, pf:pm, pt:same > last ? last : same }; }
  from = ymd_(from) || t; to = ymd_(to) || from; if (to < from) { var x = from; from = to; to = x; }
  if (daysIn_(from, to).length > 366) fail_('Elige un rango de un año o menos');
  n = daysIn_(from, to).length; return { from:from, to:to, pf:addDays_(from, -n), pt:addDays_(from, -1) };
}
function getDashboard(tok, key, from, to) {
  return api_(tok, 'dashboard', 'v', function(){
    var r = dashRange_(String(key || 'today'), from, to), cur = ordersIn_(r.from, r.to), paid = cur.filter(isPaid_), hourly = r.from === r.to;
    // trend: per hour on one day, else per day (every day listed, zeros included)
    var sales = {}, cnt = {};
    paid.forEach(function(o){ var k = hourly ? hourOf_(o.placed_at) : ymd_(o.placed_at); inc_(sales, k, Number(o.grand_total) || 0); inc_(cnt, k); });
    var keys = hourly ? Object.keys(cnt).concat(openHours_(r.from)).filter(function(h, i, a){ return a.indexOf(h) === i; }).sort() : daysIn_(r.from, r.to);
    var trend = keys.map(function(k){ return { k:k, sales:round_(sales[k] || 0), paid:cnt[k] || 0 }; });
    // top drinks by qty
    var q = {}, amt = {}, nm = {};
    paid.forEach(function(o){ (o.items || []).forEach(function(l){ inc_(q, l.product_id, l.qty); inc_(amt, l.product_id, l.line_total); nm[l.product_id] = l.product_name; }); });
    var top = Object.keys(q).map(function(id){ return { name:nm[id], qty:q[id], sales:round_(amt[id]) }; }).sort(function(a, b){ return b.qty - a.qty || b.sales - a.sales; }).slice(0, 10);
    // info boxes
    var lines = 0, withAdd = 0, appr = [], subm = 0, rej = 0;
    paid.forEach(function(o){ (o.items || []).forEach(function(l){ lines++; if ((l.addons || []).length) withAdd++; });
      if (o.paid_at) appr.push((Date.parse(o.paid_at) - Date.parse(o.placed_at)) / 60000); });
    cur.forEach(function(o){ (o.payment_history || []).forEach(function(p){ subm++; if (p.result === 'REJECTED') rej++; }); });
    var k = kpis_(cur), s = settings_(), prods = jdbList_(PROD_SHEET).filter(function(p){ return p.is_active; }), adds = jdbList_(ADDON_SHEET).filter(function(a){ return a.is_active; });
    return ok_({ range:{ key:key, from:r.from, to:r.to, prev_from:r.pf, prev_to:r.pt, hourly:hourly }, kpi:k, prev:kpis_(ordersIn_(r.pf, r.pt)), trend:trend, top:top,
      info:{ attach:lines ? Math.round(withAdd / lines * 100) : 0, guest:cur.length ? Math.round(cur.filter(function(o){ return o.checkout_type === 'GUEST'; }).length / cur.length * 100) : 0,
        approval_min:appr.length ? Math.round(appr.reduce(function(a, b){ return a + b; }, 0) / appr.length) : null, rejected:rej, rejected_pct:subm ? Math.round(rej / subm * 100) : 0,
        cancelled:k.cancelled, refund_pending:k.refund_pending },
      shop:{ sold_out_drinks:prods.filter(function(p){ return !p.is_available; }).length, sold_out_addons:adds.filter(function(a){ return !a.is_available; }).length,
        active_methods:jdbList_(PM_SHEET).filter(function(m){ return m.is_active; }).length, ordering_enabled:Number(s.ordering_enabled) ? 1 : 0, open:openState_(s, new Date()) } });
  });
}
// the hours of that day the shop is open (so a quiet hour still shows as a zero column)
function openHours_(day) {
  if (Number(settings_().always_open)) return Array.apply(null, Array(24)).map(function(x, i){ return p2_(i); });
  var h = (settings_().hours || DEFAULT_HOURS).filter(function(x){ return x.day === dow_(day); })[0] || {};
  if (!h.open || Number(h.closed)) return [];
  var out = []; for (var x = Number(h.open.slice(0, 2)); x <= Number(String(h.close).slice(0, 2)) && x < 24; x++) out.push(p2_(x));
  return out;
}

// ============== Reports — one period, optional filters, every breakdown tab in ONE payload ==============
function getReport(tok, from, to, f) {
  return api_(tok, 'reports', 'v', function(){
    from = ymd_(from) || todayYmd_(); to = ymd_(to) || from; if (to < from) { var x = from; from = to; to = x; }
    if (daysIn_(from, to).length > 366) fail_('Elige un período de un año o menos');
    f = f || {};
    var list = ordersIn_(from, to).filter(function(o){ return (!f.method || String(o.payment_method_id) === String(f.method)) && (!f.ful || o.fulfilment === f.ful) && (!f.type || o.checkout_type === f.type); });
    var paid = list.filter(isPaid_), k = kpis_(list), pct = function(v){ return k.sales ? Math.round(v / k.sales * 1000) / 10 : 0; };
    var prods = {}; jdbList_(PROD_SHEET).forEach(function(p){ prods[p.id] = p; });
    var cats = {}; jdbList_(CAT_SHEET).forEach(function(c){ cats[c.id] = c.name; });
    // by day
    var D = {}; list.forEach(function(o){ var d = ymd_(o.placed_at), r = D[d] || (D[d] = { orders:0, paid:0, sales:0, cancelled:0 }); r.orders++;
      if (o.status === 'CANCELLED') r.cancelled++; if (isPaid_(o)) { r.paid++; r.sales += Number(o.grand_total) || 0; } });
    var byDay = daysIn_(from, to).map(function(d){ var r = D[d] || { orders:0, paid:0, sales:0, cancelled:0 }; return { day:d, orders:r.orders, paid:r.paid, sales:round_(r.sales), avg:r.paid ? round_(r.sales / r.paid) : 0, cancelled:r.cancelled }; });
    // drinks · sizes · add-ons · hours · methods — one pass over the paid lines
    var P = {}, S = {}, A = {}, H = {}, M = {}, offer = {};
    paid.forEach(function(o){
      var h = hourOf_(o.placed_at), hr = H[h] || (H[h] = { orders:0, sales:0 }); hr.orders++; hr.sales += Number(o.grand_total) || 0;
      var mk = o.payment_method_id + '|' + ymd_(o.placed_at), m = M[mk] || (M[mk] = { method_id:o.payment_method_id, method:(o.payment_snapshot || {}).name || '', day:ymd_(o.placed_at), orders:0, amount:0 });
      m.orders++; m.amount += Number(o.grand_total) || 0;
      (o.items || []).forEach(function(l){
        var p = P[l.product_id] || (P[l.product_id] = { product_id:l.product_id, drink:l.product_name, category:cats[(prods[l.product_id] || {}).category_id] || '', qty:0, sales:0 }); p.qty += l.qty; p.sales += l.line_total;
        if (l.size_name) { var sk = l.product_id + '|' + l.size_name, z = S[sk] || (S[sk] = { drink:l.product_name, size:l.size_name, qty:0, sales:0 }); z.qty += l.qty; z.sales += l.line_total; }
        ((prods[l.product_id] || {}).addon_ids || []).forEach(function(a){ inc_(offer, a); });            // lines of drinks that offer each add-on
        (l.addons || []).forEach(function(a){ var r = A[a.addon_id] || (A[a.addon_id] = { addon:a.name, times:0, lines:0, revenue:0 }); r.times += l.qty; r.lines++; r.revenue += a.price * l.qty; });
      });
    });
    var sortBy = function(k2){ return function(a, b){ return b[k2] - a[k2]; }; };
    var money = function(r, keys){ keys.forEach(function(k2){ r[k2] = round_(r[k2]); }); return r; };
    return ok_({ from:from, to:to, summary:k, byDay:byDay,
      byDrink:Object.keys(P).map(function(i){ var r = money(P[i], ['sales']); r.share = pct(r.sales); return r; }).sort(sortBy('sales')),
      bySize:Object.keys(S).map(function(i){ return money(S[i], ['sales']); }).sort(function(a, b){ return a.drink < b.drink ? -1 : a.drink > b.drink ? 1 : b.qty - a.qty; }),
      byAddon:Object.keys(A).map(function(i){ var r = money(A[i], ['revenue']); r.attach = offer[i] ? Math.round(r.lines / offer[i] * 1000) / 10 : 0; delete r.lines; return r; }).sort(sortBy('revenue')),
      byMethod:Object.keys(M).map(function(i){ return money(M[i], ['amount']); }).sort(function(a, b){ return a.method < b.method ? -1 : a.method > b.method ? 1 : a.day < b.day ? -1 : 1; }),
      byHour:Object.keys(H).sort().map(function(h){ return { hour:h, orders:H[h].orders, sales:round_(H[h].sales) }; }),
      cancelled:list.filter(function(o){ return o.status === 'CANCELLED'; }).map(function(o){ return { order_no:o.order_no, short_no:o.short_no, day:ymd_(o.placed_at), customer:o.customer_name, phone:o.customer_phone,
        amount:o.grand_total, reason:o.cancel_reason, by:o.cancelled_by, paid:!!o.paid_at, refund_status:o.refund_status, refund_note:o.refund_note }; }),
      methods:jdbList_(PM_SHEET).map(function(m){ return { id:m.id, name:m.name }; }) });
  });
}

// ============== Customers — every account + every guest phone that ordered; a guest phone that matches an account counts under it ==============
function getCustomers(tok) {
  return api_(tok, 'customers', 'v', function(){
    var C = {}, byPhone = {};
    userRows_().data.slice(1).forEach(function(r){
      if (String(r[U.ROLE]) !== 'Customer') return;
      var u = String(r[U.NAME]), ph = digits_(r[U.PHONE]);
      C[u] = { key:u, type:'Account', username:u, name:String(r[U.FULL] || u), phone:String(r[U.PHONE] || ''), email:String(r[U.EMAIL] || ''), status:String(r[U.STATUS] || ''), created:iso_(r[U.CREATED]) || '' };
      if (ph && !byPhone[ph]) byPhone[ph] = u;
    });
    (sh_(ORDER_SHEET) ? JDB.all(ORDER_SHEET) : []).forEach(function(o){
      var ph = digits_(o.customer_phone), key = o.customer_username && C[o.customer_username] ? o.customer_username : byPhone[ph] || ('g:' + ph), c = C[key];
      if (!ph && !o.customer_username) return;
      if (!c) c = C[key] = { key:key, type:'Guest', username:'', name:o.customer_name, phone:o.customer_phone, email:o.customer_email || '', status:'', created:'' };
      if (c.type === 'Account' && !o.customer_username) inc_(c, 'guest_orders');
      inc_(c, 'orders'); if (o.status === 'CANCELLED') inc_(c, 'cancelled');
      if (isPaid_(o)) { inc_(c, 'paid'); inc_(c, 'total', Number(o.grand_total) || 0);
        (o.items || []).forEach(function(l){ c.dq = c.dq || {}; inc_(c.dq, l.product_name, l.qty); (l.addons || []).forEach(function(a){ c.aq = c.aq || {}; inc_(c.aq, a.name, l.qty); }); }); }
      if (!c.last_at || o.placed_at > c.last_at) { c.last_at = o.placed_at; c.last_status = o.status; c.last_no = o.order_no; if (c.type === 'Guest') c.name = o.customer_name; }
    });
    var fav = function(m){ return m ? Object.keys(m).sort(function(a, b){ return m[b] - m[a] || (a < b ? -1 : 1); })[0] : ''; };
    var rows = Object.keys(C).map(function(k){ var c = C[k];
      return { key:c.key, type:c.type, username:c.username, name:c.name, phone:c.phone, email:c.email, status:c.status, created:c.created, orders:c.orders || 0, paid:c.paid || 0, cancelled:c.cancelled || 0,
        total:round_(c.total || 0), avg:c.paid ? round_(c.total / c.paid) : 0, last_at:c.last_at || '', last_status:c.last_status || '', last_no:c.last_no || '', guest_orders:c.guest_orders || 0,
        fav_drink:fav(c.dq), fav_addon:fav(c.aq) }; });
    return ok_({ rows:rows.sort(function(a, b){ return (b.last_at || '') < (a.last_at || '') ? -1 : 1; }) });
  });
}

// ============== About — the live RBAC matrix for the About page ==============
function getAboutInfo(tok) {
  return api_(tok, 'about', 'v', function(){
    var perms = {};
    readRoles_().forEach(function(r){ perms[r.key] = r.perms; });
    return ok_({ pages:RBAC_PAGES, roles:rolesOut_(), perms:perms });
  });
}

// ============== Activity Logs ==============
var LOGS_VIEW_MAX = 1000;   // newest N shipped to the UI — full history stays in the sheet
function getLogs(tok) {
  return api_(tok, 'logs', 'v', function(){
    var s = sh_(LOGS_SHEET), last = s ? s.getLastRow() : 0;
    if (last < 2) return ok_({ data:[] });
    var n = Math.min(last - 1, LOGS_VIEW_MAX);                    // tail only — bounded however big the log grows
    return ok_({ data: s.getRange(last - n + 1, 1, n, 4).getValues().map(function(r){
      return { Timestamp:iso_(r[0]) || '', User:r[1] || '', Action:r[2] || '', Details:r[3] || '' };
    }).reverse() });                                              // newest first
  });
}

// ============== Nightly job — backup copy + retention, counter pruning, old reset links, open-days index ==============
var BACKUP_PREFIX = 'Respaldo Tienda de Bebidas ';
// editor / trigger only: a web visitor runs as the owner with a different (or blank) active user, so it is refused
var ownerRun_ = function(){ var a = Session.getActiveUser().getEmail(), e = Session.getEffectiveUser().getEmail(); return !!a && a === e; };
function backup_() {
  var s = settings_(), folder = subFolder_(receiptsRoot_(), 'Backups', 'DRV_BACKUP'), stamp = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HHmm');
  var copy = DriveApp.getFileById(ss_().getId()).makeCopy(BACKUP_PREFIX + stamp, folder);           // private like the source — never shared
  if (s.backup_folder_id !== folder.getId()) settingSet_('backup_folder_id', folder.getId());
  var cut = Date.now() - (Number(s.backup_keep_days) || 30) * 864e5, gone = 0, it = folder.getFiles();
  while (it.hasNext()) { var f = it.next(); if (f.getId() !== copy.getId() && String(f.getName()).indexOf(BACKUP_PREFIX) === 0 && f.getDateCreated().getTime() < cut) { f.setTrashed(true); gone++; } }
  PropertiesService.getScriptProperties().setProperty('BACKUP_LAST', todayYmd_());
  return { name:copy.getName(), removed:gone };
}
// number counters are per prefix per day — keep the last 3 days (properties store limit)
function pruneCounters_() {
  var p = PropertiesService.getScriptProperties(), cut = addDays_(todayYmd_(), -3).replace(/-/g, ''), n = 0;
  p.getKeys().forEach(function(k){ var m = /^SEQ_[A-Z]+_(\d{8})$/.exec(k); if (m && m[1] < cut) { p.deleteProperty(k); n++; } });
  return n;
}
// reset links older than 7 days are dead weight — one rewrite of the survivors
function purgeResets_() {
  var sh = sh_(RESETS_SHEET); if (!sh || sh.getLastRow() < 2) return 0;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, RESET_HEAD.length).getValues(), cut = Date.now() - 7 * 864e5,
      keep = rows.filter(function(r){ return new Date(r[4] || r[2]).getTime() >= cut; });
  if (keep.length === rows.length) return 0;
  sh.getRange(2, 1, rows.length, RESET_HEAD.length).clearContent();
  putText_(sh, 2, keep);
  return rows.length - keep.length;
}
function nightlyJob() {
  if (!ownerRun_()) return err_('No permitido — esto se ejecuta desde su activador nocturno');
  var out = { pruned:pruneCounters_(), resets:withLock_(purgeResets_), open_days:openScan_().dropped };
  if (PropertiesService.getScriptProperties().getProperty('BACKUP_LAST') !== todayYmd_()) { try { out.backup = backup_(); } catch (e) { out.backup = { error:String((e && e.message) || e) }; } }
  out.daily_mail = mailSend_('shop.daily', null, settings_(), { day:addDays_(todayYmd_(), -1) });
  addLog_('system', 'Tarea Nocturna', 'contadores depurados ' + out.pruned + ' · enlaces de restablecimiento depurados ' + out.resets + ' · días abiertos eliminados ' + out.open_days + (out.backup ? ' · respaldo ' + (out.backup.name || out.backup.error) : ''));
  return ok_(out);
}
// ONE nightly trigger (02:00 script time); re-installing replaces it
function installNightly_() {
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction() === 'nightlyJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('nightlyJob').timeBased().atHour(2).everyDays(1).create();
}
function installNightlyTrigger(tok) {
  return api_(tok, 'settings', 'e', function(me){ installNightly_(); addLog_(me.u, 'Activador Nocturno Instalado', '02:00 diario');
    return ok_({ message:'Tarea nocturna instalada — el respaldo y la limpieza se ejecutan todas las noches a las 02:00' }); });
}
function backupNow(tok) {
  return api_(tok, 'settings', 'e', function(me){ var b = backup_(); addLog_(me.u, 'Respaldo Realizado', b.name + (b.removed ? ' · ' + b.removed + ' antiguos eliminados' : ''));
    return ok_({ message:'Respaldo guardado como "' + b.name + '"' + (b.removed ? ' · ' + b.removed + ' respaldo(s) antiguo(s) eliminado(s)' : '') }); });
}
function getMaintenance(tok) {
  return api_(tok, 'settings', 'e', function(){ var p = PropertiesService.getScriptProperties();
    var q = -1; try { q = MailApp.getRemainingDailyQuota(); } catch (e) {}
    return ok_({ nightly:ScriptApp.getProjectTriggers().some(function(t){ return t.getHandlerFunction() === 'nightlyJob'; }), last_backup:p.getProperty('BACKUP_LAST') || '', keep:settings_().backup_keep_days, mail_quota:q }); });
}
// test email for the Notifications card
function sendTestEmail(tok) {
  return api_(tok, 'settings', 'e', function(me){
    var s = settings_(), to = s.notify_shop_email || s.shop_email;
    if (!to || !String(to).split(', ').every(function(x){ return EMAIL_RX.test(x); })) return err_('Guarda primero un correo de la tienda o de notificaciones');
    if (MailApp.getRemainingDailyQuota() < 1) return err_("El límite de correos de hoy se ha agotado — se restablecerá mañana");
    MailApp.sendEmail({ to:to, subject:s.shop_name + ' — correo de prueba (ejemplo de nuevo pedido)', htmlBody:mailOf_('shop.new_order', sampleOrder_(), s).html, name:String(s.shop_name || 'Tienda').slice(0, 60) });
    addLog_(me.u, 'Correo de Prueba', 'Enviado a ' + to);
    return ok_({ message:'Correo de prueba enviado a ' + to });
  });
}
// preview any event with the newest real order (or a sample) — nothing is sent
function previewEmail(tok, ev) {
  return api_(tok, 'settings', 'v', function(){
    if (MAIL_EVENTS.indexOf(ev) === -1) return err_('Correo desconocido');
    var s = settings_(), o = sampleOrder_(), x = { name:'Customer 1', email:'customer1@demo.com', phone:'03001000002', day:addDays_(todayYmd_(), -1), left:2 };
    if (ev === 'customer.rejected' && !o.reject_reason) o.reject_reason = "El monto no coincide";
    if (ev === 'customer.cancelled' && !o.cancel_reason) o.cancel_reason = 'Falta de un ingrediente';
    var m = mailOf_(ev, o, s, x);
    return ok_({ subject:m.subject, html:m.html, to:m.to || (ev.indexOf('shop.') === 0 ? '(sin correo de notificación aún)' : "(correo del cliente)") });
  });
}
// the newest order of the last 30 days, else a made-up one — for the preview and the test email
function sampleOrder_() {
  var L = sh_(ORDER_SHEET) ? JDB.between(ORDER_SHEET, addDays_(todayYmd_(), -30), null) : [];
  var o = L.sort(function(a, b){ return a.placed_at < b.placed_at ? 1 : -1; })[0];
  return o ? JSON.parse(JSON.stringify(o)) : { short_no:'#001', order_no:'ORD-' + todayYmd_().replace(/-/g, '') + '-0001', customer_name:'Cliente 1', customer_phone:'3001234567', customer_email:'cliente1@demo.com',
    items:[{ qty:2, product_name:'Té con Leche Clásico', size_name:'Grande', sugar:'Normal', ice:'Menos hielo', addons:[{ name:'Perlas de tapioca' }], line_total:29000 }], subtotal:29000, delivery_fee:0, grand_total:29000,
    payment_snapshot:{ name:'Billetera QR (Nequi)' }, fulfilment:'PICKUP', pickup_mode:'ASAP', track_token:'', payment_attempt:1 };
}

// ============== Setup Demo Data ==============
// generic placeholders only — YouTube-safe, no real PII
var DEMO_SHOP = { shop_name:'Bebidas Demo', shop_address:'Carrera 7 # 71-21, Bogotá, Colombia', shop_phone:'+57 300 123 4567', shop_email:'tienda@demo.com',
  map_url:'https://maps.google.com/?q=Bogota+Colombia', about_text:'Bebidas frescas preparadas a tu gusto — elige tamaño, nivel de azúcar, hielo y complementos.',
  hero_image:demoPhoto_('1551024709-8f23befc6f87', 1800, 700), hero_title:'Tu bebida, a tu gusto', hero_subtitle:'Elige tu bebida, agrega tus extras, paga con Nequi o transferencia — lista cuando llegues.', footer_note:'¡Muchas gracias por tu compra — esperamos verte pronto!' };
var daysAgo_ = function(n, hhmm){ var d = new Date(Date.now() - n * 864e5), p = (hhmm || '09:00').split(':');
  d.setHours(+p[0], +p[1], 0, 0); return d.toISOString(); };

// wipes ALL sheets and rebuilds fresh with demo data
function setupDemoData() {
  try {
    // editor-only: from the deployed web app getActiveUser() is blank, so a visitor calling it is refused
    if (Session.getActiveUser().getEmail() !== Session.getEffectiveUser().getEmail())
      return err_('No permitido — ejecuta setupDemoData desde el editor de Apps Script');

    _m = {};                                                                     // sheets are about to be dropped — drop the handles too
    var ss = ss_(), temp = ss.insertSheet('__temp_' + Date.now());               // a spreadsheet must always keep >=1 sheet
    ss.getSheets().forEach(function(sh){ if (sh.getSheetId() !== temp.getSheetId()) ss.deleteSheet(sh); });
    _m.sh = {};
    var props = PropertiesService.getScriptProperties();                         // record + number counters restart with the data
    props.getKeys().forEach(function(k){ if (/^(SEQ_|JID_|OPEN_DAYS)/.test(k)) props.deleteProperty(k); });
    CACHE_.removeAll([SET_KEY, UIX_KEY, ROLES_KEY, FRONT_KEY, PAYIX_KEY]);

    var usersSheet = head_(ss.insertSheet(USERS_SHEET), USER_HEAD), logsSheet = head_(ss.insertSheet(LOGS_SHEET), LOG_HEAD);
    bustRoles_(); ensureRbac_(); resetsSheet_(); settingsSheet_();

    // cols: username, pwd, role, status, full name, phone, address
    var demoUsers = [
      ['admin', 'admin123', 'Admin', 'Active', 'Admin 1', '03001000001', ''],
      ['customer1@demo.com', 'customer123', 'Customer', 'Active', 'Customer 1', '03001000002', 'House 1, Street 1, Demo City'],
      ['customer2@demo.com', 'customer123', 'Customer', 'Active', 'Customer 2', '03001000003', 'House 2, Street 2, Demo City'],
      ['customer3@demo.com', 'customer123', 'Customer', 'Inactive', 'Customer 3', '03001000004', 'House 3, Street 3, Demo City']
    ], n = demoUsers.length;
    putText_(usersSheet, 2, demoUsers.map(function(u, i){
      return userRow_({ name:u[0], email:u[0] === 'admin' ? 'admin@demo.com' : u[0], pwd:u[1], role:u[2], status:u[3], by:'System',
                        ts:new Date(Date.now() - (n - i + 20) * 864e5).toISOString(), full:u[4], phone:u[5], addr:u[6] });
    }));

    var demoLogs = [
      ['admin', 'Configuración del Sistema', 'Datos de prueba inicializados'], ['admin', 'Inicio Exitoso', 'Usuario inició sesión exitosamente'],
      ['admin', 'AJUSTES_GUARDADOS', 'Detalles de la tienda actualizados'], ['customer1@demo.com', 'REGISTRO', 'Cuenta de cliente creada'],
      ['customer2@demo.com', 'REGISTRO', 'Cuenta de cliente creada'], ['customer1@demo.com', 'Inicio Exitoso', 'Usuario inició sesión exitosamente'],
      ['admin', 'Permisos Actualizados', 'Customer · about · v=0']
    ], m = demoLogs.length;
    putText_(logsSheet, 2, demoLogs.map(function(l, i){ return [new Date(Date.now() - (m - i) * 7200000).toISOString(), l[0], l[1], l[2]]; }));   // 2h apart

    var s = {}; Object.keys(SET_DEFAULTS).forEach(function(k){ s[k] = DEMO_SHOP.hasOwnProperty(k) ? DEMO_SHOP[k] : SET_DEFAULTS[k]; });
    writeSettings_(s);
    seedMenu_();
    seedPayments_();
    seedOrders_();

    ss.deleteSheet(temp);
    _m = {};
    try { installNightly_(); } catch (e) {}                                       // the nightly job arms itself on install
    addLog_('admin', 'Configuración del Sistema', 'Datos de prueba inicializados');
    return ok_({ message:'Datos de prueba creados. Accesos: admin / admin123 · customer1@demo.com / customer123' });
  } catch (e) { return err_('Falló la configuración: ' + ((e && e.user) || (e && e.message) || e)); }
}
