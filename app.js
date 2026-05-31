// ============================================================
//  E-자산통합관리 - 메인 서버 (app.js)
//  실행: node app.js  →  http://localhost:3000
// ============================================================
const express = require('express');
const Database = require('better-sqlite3');
const path = require('path');
const crypto = require('crypto');
const nodemailer = require('nodemailer');

const PORT = process.env.PORT || 3000;
// 클라우드(Fly.io)에서는 DB_PATH 환경변수로 볼륨 경로 사용
// 로컬에서는 기존대로 it_manager.db 사용
const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'it_manager.db');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// 직원용 장애 접수 페이지
app.get('/report', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'report.html'));
});

// ──────────────────────────────────────────────
//  DB 초기화
// ──────────────────────────────────────────────
const db = new Database(DB_PATH);
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT NOT NULL,
    department  TEXT,
    position    TEXT,
    email       TEXT,
    phone       TEXT,
    created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  CREATE TABLE IF NOT EXISTS assets (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    name             TEXT NOT NULL,
    category         TEXT,
    serial_number    TEXT,
    manufacturer     TEXT,
    model            TEXT,
    purchase_date    TEXT,
    purchase_price   REAL,
    status           TEXT DEFAULT '사용중',
    location         TEXT,
    assigned_user_id INTEGER,
    notes            TEXT,
    created_at       DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (assigned_user_id) REFERENCES users(id) ON DELETE SET NULL
  );

  CREATE TABLE IF NOT EXISTS tickets (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    title        TEXT NOT NULL,
    description  TEXT,
    priority     TEXT DEFAULT '보통',
    status       TEXT DEFAULT '접수',
    requester    TEXT,
    assignee     TEXT,
    asset_id     INTEGER,
    created_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
    updated_at   DATETIME DEFAULT CURRENT_TIMESTAMP,
    resolved_at  DATETIME
  );

  CREATE TABLE IF NOT EXISTS software (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    name           TEXT NOT NULL,
    version        TEXT,
    license_key    TEXT,
    license_count  INTEGER DEFAULT 1,
    used_count     INTEGER DEFAULT 0,
    vendor         TEXT,
    purchase_date  TEXT,
    expiry_date    TEXT,
    notes          TEXT,
    created_at     DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  CREATE TABLE IF NOT EXISTS maintenance (
    id                    INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_id              INTEGER,
    type                  TEXT,
    description           TEXT,
    performed_by          TEXT,
    cost                  REAL,
    maintenance_date      TEXT,
    next_maintenance_date TEXT,
    created_at            DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (asset_id) REFERENCES assets(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS admin_accounts (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT NOT NULL UNIQUE,
    password    TEXT NOT NULL,
    role        TEXT DEFAULT 'sub',
    name        TEXT,
    team        TEXT,
    employee_id TEXT,
    granted_at  TEXT,
    created_at  DATETIME DEFAULT CURRENT_TIMESTAMP
  );

  CREATE TABLE IF NOT EXISTS support_messages (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT,
    category   TEXT,
    subject    TEXT,
    content    TEXT,
    status     TEXT DEFAULT '접수',
    email_sent INTEGER DEFAULT 0,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
`);

// 마스터 관리자 시드 (없을 때만)
const masterExists = db.prepare("SELECT id FROM admin_accounts WHERE username='admin'").get();
if (!masterExists) {
  db.prepare("INSERT INTO admin_accounts (username, password, role) VALUES ('admin', '1234', 'master')").run();
}

// 기존 DB에 컬럼 추가 (이미 있으면 무시)
try { db.exec('ALTER TABLE users ADD COLUMN employee_id TEXT'); } catch(e) {}
try { db.exec('ALTER TABLE tickets ADD COLUMN process_notes TEXT'); } catch(e) {}
try { db.exec('ALTER TABLE tickets ADD COLUMN reported_asset_number TEXT'); } catch(e) {}
try { db.exec('ALTER TABLE tickets ADD COLUMN manufacturer_model TEXT'); } catch(e) {}
try { db.exec('ALTER TABLE software ADD COLUMN assigned_user_id INTEGER REFERENCES users(id)'); } catch(e) {}

// ──────────────────────────────────────────────
//  헬퍼
// ──────────────────────────────────────────────
const ok   = (res, data) => res.json({ success: true,  data });
const fail = (res, msg, code = 400) => res.status(code).json({ success: false, message: msg });
const escapeHtml = (value) => String(value ?? '')
  .replace(/&/g, '&amp;')
  .replace(/</g, '&lt;')
  .replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;')
  .replace(/'/g, '&#39;');

async function sendSupportMail({ name, category, subject, content, heading = 'E-자산통합관리 — 고객센터 문의' }) {
  const smtpUser = process.env.SMTP_USER;
  const smtpPass = process.env.SMTP_PASS;
  if (!smtpUser || !smtpPass) return false;

  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST || 'smtp.office365.com',
    port: parseInt(process.env.SMTP_PORT || '587'),
    secure: false,
    auth: { user: smtpUser, pass: smtpPass },
    tls: { ciphers: 'SSLv3', rejectUnauthorized: false }
  });
  const now = new Date().toLocaleString('ko-KR', { timeZone: 'Asia/Seoul' });
  const to = process.env.SUPPORT_EMAIL || process.env.SMTP_TO || 'changgyo@cgessence.co.kr';

  await transporter.sendMail({
    from: `"E-자산통합관리 고객센터" <${smtpUser}>`,
    to,
    subject: `[E-자산통합관리] [${category||'문의'}] ${subject}`,
    html: `
      <div style="font-family:Segoe UI,sans-serif;max-width:600px;margin:0 auto;border:1px solid #dee2e6;border-radius:10px;overflow:hidden">
        <div style="background:#1e3a5f;padding:20px 28px">
          <h2 style="color:#fff;margin:0;font-size:1.1rem">${escapeHtml(heading)}</h2>
        </div>
        <div style="padding:24px 28px;background:#fff">
          <table style="width:100%;border-collapse:collapse;font-size:.9rem">
            <tr><td style="padding:8px 0;color:#6c757d;width:90px">문 의 자</td><td style="padding:8px 0;font-weight:600">${escapeHtml(name||'-')}</td></tr>
            <tr><td style="padding:8px 0;color:#6c757d">구 &nbsp;&nbsp;&nbsp;&nbsp; 분</td><td style="padding:8px 0">${escapeHtml(category||'-')}</td></tr>
            <tr><td style="padding:8px 0;color:#6c757d">제 &nbsp;&nbsp;&nbsp;&nbsp; 목</td><td style="padding:8px 0;font-weight:600">${escapeHtml(subject)}</td></tr>
            <tr><td style="padding:8px 0;color:#6c757d;vertical-align:top">내 &nbsp;&nbsp;&nbsp;&nbsp; 용</td><td style="padding:8px 0;white-space:pre-line;line-height:1.7">${escapeHtml(content)}</td></tr>
            <tr><td style="padding:8px 0;color:#6c757d">접수 시간</td><td style="padding:8px 0;font-size:.82rem;color:#6c757d">${escapeHtml(now)}</td></tr>
          </table>
        </div>
        <div style="padding:12px 28px;background:#f8f9fa;font-size:.8rem;color:#adb5bd">
          본 메일은 E-자산통합관리(it-manager.fly.dev)에서 자동 발송되었습니다.
        </div>
      </div>`
  });

  return true;
}

// ──────────────────────────────────────────────
//  세션 / 인증
// ──────────────────────────────────────────────
const sessions = new Map(); // token -> { username, role, name }
const passwordResetBuckets = new Map(); // ip -> timestamps

function makeToken() {
  return crypto.randomBytes(32).toString('hex');
}

function requireAuth(req, res, next) {
  const token = req.headers['x-auth-token'];
  const sess = sessions.get(token);
  if (!sess) return fail(res, '인증이 필요합니다.', 401);
  req.user = sess;
  next();
}

function requireMaster(req, res, next) {
  if (req.user.role !== 'master') return fail(res, '권한이 없습니다.', 403);
  next();
}

// ──────────────────────────────────────────────
//  인증 API
// ──────────────────────────────────────────────
app.post('/api/auth/login', (req, res) => {
  try {
    const { username, password } = req.body;
    if (!username || !password) return fail(res, '아이디와 비밀번호를 입력하세요.');
    const account = db.prepare('SELECT * FROM admin_accounts WHERE username=?').get(username);
    if (!account || account.password !== password) return fail(res, '아이디 또는 비밀번호가 올바르지 않습니다.');
    const token = makeToken();
    sessions.set(token, { username: account.username, role: account.role, name: account.name || account.username });
    ok(res, { token, username: account.username, role: account.role, name: account.name || account.username });
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/auth/logout', requireAuth, (req, res) => {
  const token = req.headers['x-auth-token'];
  sessions.delete(token);
  ok(res, {});
});

app.get('/api/auth/me', requireAuth, (req, res) => {
  ok(res, req.user);
});

app.post('/api/password-reset-request', async (req, res) => {
  try {
    const ip = req.headers['fly-client-ip'] || req.headers['x-forwarded-for'] || req.ip || 'unknown';
    const nowMs = Date.now();
    const recent = (passwordResetBuckets.get(ip) || []).filter(t => nowMs - t < 10 * 60 * 1000);
    if (recent.length >= 5) return fail(res, '요청이 많습니다. 잠시 후 다시 시도해주세요.', 429);
    recent.push(nowMs);
    passwordResetBuckets.set(ip, recent);

    const username = String(req.body.username || '').trim().slice(0, 80);
    const name = String(req.body.name || '').trim().slice(0, 80);
    const contact = String(req.body.contact || '').trim().slice(0, 160);
    const message = String(req.body.message || '').trim().slice(0, 1000);
    if (!username || !contact) return fail(res, '아이디와 연락처를 입력해주세요.');

    const account = db.prepare(
      'SELECT username, role, name, team, employee_id FROM admin_accounts WHERE username=?'
    ).get(username);

    const category = '비밀번호 찾기';
    const subject = `비밀번호 재설정 요청 - ${username}`;
    const requestName = name || contact || username;
    const content = [
      '관리자 로그인 비밀번호 재설정 요청이 접수되었습니다.',
      '',
      `아이디: ${username}`,
      `이름: ${name || '-'}`,
      `연락처/이메일: ${contact}`,
      `요청 내용: ${message || '-'}`,
      '',
      account
        ? `계정 확인: 일치 (${account.role || '-'} / ${account.name || account.username}${account.team ? ' / ' + account.team : ''})`
        : '계정 확인: 일치하는 관리자 계정 없음',
      `요청 IP: ${ip}`,
      `User-Agent: ${req.headers['user-agent'] || '-'}`
    ].join('\n');

    const row = db.prepare(
      'INSERT INTO support_messages (name, category, subject, content) VALUES (?, ?, ?, ?)'
    ).run(requestName, category, subject, content);

    let emailSent = false;
    try {
      emailSent = await sendSupportMail({
        name: requestName,
        category,
        subject,
        content,
        heading: 'E-자산통합관리 — 비밀번호 재설정 요청'
      });
      if (emailSent) db.prepare('UPDATE support_messages SET email_sent=1 WHERE id=?').run(row.lastInsertRowid);
    } catch (mailErr) {
      console.error('비밀번호 찾기 메일 발송 실패:', mailErr.message);
    }

    ok(res, { emailSent });
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  부관리자 관리 (master 전용)
// ──────────────────────────────────────────────
app.get('/api/admins', requireAuth, requireMaster, (req, res) => {
  try {
    const rows = db.prepare("SELECT id, username, name, team, employee_id, granted_at, created_at FROM admin_accounts WHERE role='sub' ORDER BY id DESC").all();
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/admins', requireAuth, requireMaster, (req, res) => {
  try {
    const { username, password, name, team, employee_id, granted_at } = req.body;
    if (!username || !password) return fail(res, '아이디와 비밀번호는 필수입니다.');
    const dup = db.prepare('SELECT id FROM admin_accounts WHERE username=?').get(username);
    if (dup) return fail(res, '이미 존재하는 아이디입니다.');
    const r = db.prepare(`
      INSERT INTO admin_accounts (username, password, role, name, team, employee_id, granted_at)
      VALUES (?, ?, 'sub', ?, ?, ?, ?)
    `).run(username, password, name, team, employee_id, granted_at);
    ok(res, { id: r.lastInsertRowid });
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/admins/:id', requireAuth, requireMaster, (req, res) => {
  try {
    db.prepare('DELETE FROM admin_accounts WHERE id=? AND role=\'sub\'').run(req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  대시보드
// ──────────────────────────────────────────────
app.get('/api/dashboard', requireAuth, (req, res) => {
  try {
    const assetCount    = db.prepare('SELECT COUNT(*) as cnt FROM assets').get().cnt;
    const userCount     = db.prepare('SELECT COUNT(*) as cnt FROM users').get().cnt;
    const ticketOpen    = db.prepare("SELECT COUNT(*) as cnt FROM tickets WHERE status != '완료'").get().cnt;
    const softwareCount = db.prepare('SELECT COUNT(*) as cnt FROM software').get().cnt;

    const assetByStatus = db.prepare(`
      SELECT status, COUNT(*) as cnt FROM assets GROUP BY status
    `).all();

    const ticketByPriority = db.prepare(`
      SELECT priority, COUNT(*) as cnt FROM tickets WHERE status != '완료' GROUP BY priority
    `).all();

    const recentTickets = db.prepare(`
      SELECT * FROM tickets ORDER BY created_at DESC LIMIT 5
    `).all();

    const expiringSoon = db.prepare(`
      SELECT * FROM software
      WHERE expiry_date IS NOT NULL AND expiry_date != ''
        AND DATE(expiry_date) <= DATE('now', '+30 days')
        AND DATE(expiry_date) >= DATE('now')
      ORDER BY expiry_date ASC
    `).all();

    ok(res, { assetCount, userCount, ticketOpen, softwareCount,
              assetByStatus, ticketByPriority, recentTickets, expiringSoon });
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  자산 관리
// ──────────────────────────────────────────────
app.get('/api/assets', requireAuth, (req, res) => {
  try {
    const rows = db.prepare(`
      SELECT a.*, u.name as user_name, u.department as user_dept
      FROM assets a
      LEFT JOIN users u ON a.assigned_user_id = u.id
      ORDER BY a.id DESC
    `).all();
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/assets', requireAuth, (req, res) => {
  try {
    const { name, category, serial_number, manufacturer, model,
            purchase_date, purchase_price, status, location,
            assigned_user_id, notes } = req.body;
    if (!name) return fail(res, '자산명은 필수입니다.');
    const r = db.prepare(`
      INSERT INTO assets (name, category, serial_number, manufacturer, model,
        purchase_date, purchase_price, status, location, assigned_user_id, notes)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(name, category, serial_number, manufacturer, model,
           purchase_date, purchase_price || null, status || '사용중', location,
           assigned_user_id || null, notes);
    ok(res, { id: r.lastInsertRowid });
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/assets/bulk', requireAuth, (req, res) => {
  try {
    const { assets } = req.body;
    if (!Array.isArray(assets) || assets.length === 0)
      return fail(res, '등록할 자산 데이터가 없습니다.');
    const insert = db.prepare(`
      INSERT INTO assets (name, category, serial_number, manufacturer, model,
        purchase_date, purchase_price, status, location, notes)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const insertMany = db.transaction((list) => {
      let count = 0;
      for (const a of list) {
        if (!a.name) continue;
        insert.run(
          a.name||'', a.category||'', a.serial_number||'', a.manufacturer||'', a.model||'',
          a.purchase_date||'', a.purchase_price||null, a.status||'사용중', a.location||'', a.notes||''
        );
        count++;
      }
      return count;
    });
    const count = insertMany(assets);
    ok(res, { count });
  } catch (e) { fail(res, e.message, 500); }
});

app.put('/api/assets/bulk-assign', requireAuth, (req, res) => {
  try {
    const { ids, assigned_user_id } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) return fail(res, '자산을 선택하세요.');
    const uid = assigned_user_id ? parseInt(assigned_user_id) : null;
    const update = db.transaction(() =>
      ids.forEach(id => db.prepare('UPDATE assets SET assigned_user_id=? WHERE id=?').run(uid, id))
    );
    update();
    ok(res, { count: ids.length });
  } catch (e) { fail(res, e.message, 500); }
});

app.put('/api/assets/:id', requireAuth, (req, res) => {
  try {
    const { name, category, serial_number, manufacturer, model,
            purchase_date, purchase_price, status, location,
            assigned_user_id, notes } = req.body;
    db.prepare(`
      UPDATE assets SET name=?, category=?, serial_number=?, manufacturer=?, model=?,
        purchase_date=?, purchase_price=?, status=?, location=?,
        assigned_user_id=?, notes=?
      WHERE id=?
    `).run(name, category, serial_number, manufacturer, model,
           purchase_date, purchase_price || null, status, location,
           assigned_user_id || null, notes, req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/assets/bulk', requireAuth, (req, res) => {
  try {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) return fail(res, '삭제할 항목이 없습니다.');
    const del = db.transaction(() => ids.forEach(id => db.prepare('DELETE FROM assets WHERE id=?').run(id)));
    del();
    ok(res, { count: ids.length });
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/assets/:id', requireAuth, (req, res) => {
  try {
    db.prepare('DELETE FROM assets WHERE id=?').run(req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  사용자 관리
// ──────────────────────────────────────────────
app.get('/api/users', requireAuth, (req, res) => {
  try {
    const rows = db.prepare(`
      SELECT u.*, COUNT(a.id) as asset_count
      FROM users u
      LEFT JOIN assets a ON a.assigned_user_id = u.id
      GROUP BY u.id
      ORDER BY u.id DESC
    `).all();
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/users', requireAuth, (req, res) => {
  try {
    const { name, department, position, employee_id, email, phone } = req.body;
    if (!name) return fail(res, '이름은 필수입니다.');
    const r = db.prepare(`
      INSERT INTO users (name, department, position, employee_id, email, phone)
      VALUES (?, ?, ?, ?, ?, ?)
    `).run(name, department, position, employee_id, email, phone);
    ok(res, { id: r.lastInsertRowid });
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/users/bulk', requireAuth, (req, res) => {
  try {
    const { users } = req.body;
    if (!Array.isArray(users) || users.length === 0)
      return fail(res, '등록할 사용자 데이터가 없습니다.');
    const insert = db.prepare(
      'INSERT INTO users (name, department, position, employee_id, email, phone) VALUES (?, ?, ?, ?, ?, ?)'
    );
    const insertMany = db.transaction((list) => {
      let count = 0;
      for (const u of list) {
        if (!u.name) continue;
        insert.run(u.name||'', u.department||'', u.position||'', u.employee_id||'', u.email||'', u.phone||'');
        count++;
      }
      return count;
    });
    const count = insertMany(users);
    ok(res, { count });
  } catch (e) { fail(res, e.message, 500); }
});

app.put('/api/users/:id', requireAuth, (req, res) => {
  try {
    const { name, department, position, employee_id, email, phone } = req.body;
    db.prepare(`
      UPDATE users SET name=?, department=?, position=?, employee_id=?, email=?, phone=?
      WHERE id=?
    `).run(name, department, position, employee_id, email, phone, req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/users/bulk', requireAuth, (req, res) => {
  try {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) return fail(res, '삭제할 항목이 없습니다.');
    const del = db.transaction(() => ids.forEach(id => db.prepare('DELETE FROM users WHERE id=?').run(id)));
    del();
    ok(res, { count: ids.length });
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/users/:id', requireAuth, (req, res) => {
  try {
    db.prepare('DELETE FROM users WHERE id=?').run(req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  장애/작업 티켓
// ──────────────────────────────────────────────
app.get('/api/tickets', requireAuth, (req, res) => {
  try {
    const rows = db.prepare(`
      SELECT t.*
      FROM tickets t
      ORDER BY
        CASE t.priority WHEN '긴급' THEN 1 WHEN '높음' THEN 2 WHEN '보통' THEN 3 ELSE 4 END,
        t.created_at DESC
    `).all();
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/tickets', requireAuth, (req, res) => {
  try {
    const { title, description, priority, status, requester, assignee, manufacturer_model, reported_asset_number, process_notes } = req.body;
    if (!title) return fail(res, '제목은 필수입니다.');
    const r = db.prepare(`
      INSERT INTO tickets (title, description, priority, status, requester, assignee, manufacturer_model, reported_asset_number, process_notes)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(title, description, priority || '보통', status || '접수',
           requester, assignee, manufacturer_model || null, reported_asset_number || null, process_notes || null);
    ok(res, { id: r.lastInsertRowid });
  } catch (e) { fail(res, e.message, 500); }
});

// 공개 장애 접수 (인증 불필요 - 직원용 report.html에서 사용)
app.post('/api/report', (req, res) => {
  try {
    const { title, description, priority, requester, asset_number } = req.body;
    if (!title) return fail(res, '제목은 필수입니다.');
    const r = db.prepare(`
      INSERT INTO tickets (title, description, priority, status, requester, assignee, asset_id, reported_asset_number)
      VALUES (?, ?, ?, '접수', ?, '', null, ?)
    `).run(title, description || '', priority || '보통', requester, asset_number || null);
    ok(res, { id: r.lastInsertRowid });
  } catch (e) { fail(res, e.message, 500); }
});

app.put('/api/tickets/:id', requireAuth, (req, res) => {
  try {
    const { title, description, priority, status, requester, assignee, manufacturer_model, reported_asset_number, process_notes } = req.body;
    const resolved_at = status === '완료'
      ? db.prepare("SELECT datetime('now','localtime') as t").get().t
      : null;
    db.prepare(`
      UPDATE tickets SET title=?, description=?, priority=?, status=?,
        requester=?, assignee=?, manufacturer_model=?, reported_asset_number=?, process_notes=?,
        updated_at=datetime('now','localtime'),
        resolved_at=COALESCE(?, resolved_at)
      WHERE id=?
    `).run(title, description, priority, status,
           requester, assignee, manufacturer_model || null, reported_asset_number || null, process_notes || null,
           resolved_at, req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/tickets/:id', requireAuth, (req, res) => {
  try {
    db.prepare('DELETE FROM tickets WHERE id=?').run(req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  소프트웨어 라이선스
// ──────────────────────────────────────────────
app.get('/api/software', requireAuth, (req, res) => {
  try {
    const rows = db.prepare(`
      SELECT s.*, u.name as user_name, u.department as user_dept
      FROM software s
      LEFT JOIN users u ON s.assigned_user_id = u.id
      ORDER BY s.id DESC
    `).all();
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/software', requireAuth, (req, res) => {
  try {
    const { name, version, license_key, license_count, used_count,
            vendor, purchase_date, expiry_date, notes, assigned_user_id } = req.body;
    if (!name) return fail(res, '소프트웨어명은 필수입니다.');
    const r = db.prepare(`
      INSERT INTO software (name, version, license_key, license_count, used_count,
        vendor, purchase_date, expiry_date, notes, assigned_user_id)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `).run(name, version, license_key, license_count || 1, used_count || 0,
           vendor, purchase_date, expiry_date, notes, assigned_user_id || null);
    ok(res, { id: r.lastInsertRowid });
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/software/bulk', requireAuth, (req, res) => {
  try {
    const { software } = req.body;
    if (!Array.isArray(software) || software.length === 0)
      return fail(res, '등록할 소프트웨어 데이터가 없습니다.');
    const insert = db.prepare(`
      INSERT INTO software (name, version, license_key, license_count, used_count,
        vendor, purchase_date, expiry_date, notes)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const insertMany = db.transaction((list) => {
      let count = 0;
      for (const s of list) {
        if (!s.name) continue;
        insert.run(
          s.name||'', s.version||'', s.license_key||'',
          parseInt(s.license_count)||1, parseInt(s.used_count)||0,
          s.vendor||'', s.purchase_date||'', s.expiry_date||'', s.notes||''
        );
        count++;
      }
      return count;
    });
    const count = insertMany(software);
    ok(res, { count });
  } catch (e) { fail(res, e.message, 500); }
});

app.put('/api/software/:id', requireAuth, (req, res) => {
  try {
    const { name, version, license_key, license_count, used_count,
            vendor, purchase_date, expiry_date, notes, assigned_user_id } = req.body;
    db.prepare(`
      UPDATE software SET name=?, version=?, license_key=?, license_count=?,
        used_count=?, vendor=?, purchase_date=?, expiry_date=?, notes=?, assigned_user_id=?
      WHERE id=?
    `).run(name, version, license_key, license_count, used_count,
           vendor, purchase_date, expiry_date, notes, assigned_user_id || null, req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/software/:id', requireAuth, (req, res) => {
  try {
    db.prepare('DELETE FROM software WHERE id=?').run(req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  유지보수 이력
// ──────────────────────────────────────────────
app.get('/api/maintenance', requireAuth, (req, res) => {
  try {
    const rows = db.prepare(`
      SELECT m.*, a.name as asset_name, a.category as asset_category
      FROM maintenance m
      LEFT JOIN assets a ON m.asset_id = a.id
      ORDER BY m.maintenance_date DESC, m.id DESC
    `).all();
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/maintenance', requireAuth, (req, res) => {
  try {
    const { asset_id, type, description, performed_by, cost,
            maintenance_date, next_maintenance_date } = req.body;
    const r = db.prepare(`
      INSERT INTO maintenance (asset_id, type, description, performed_by, cost,
        maintenance_date, next_maintenance_date)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(asset_id || null, type, description, performed_by,
           cost || null, maintenance_date, next_maintenance_date);
    ok(res, { id: r.lastInsertRowid });
  } catch (e) { fail(res, e.message, 500); }
});

app.delete('/api/maintenance/:id', requireAuth, (req, res) => {
  try {
    db.prepare('DELETE FROM maintenance WHERE id=?').run(req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  고객센터 (CGEssence 문의 메일 발송)
// ──────────────────────────────────────────────
app.get('/api/support', requireAuth, (req, res) => {
  try {
    const rows = db.prepare('SELECT * FROM support_messages ORDER BY created_at DESC').all();
    ok(res, rows);
  } catch (e) { fail(res, e.message, 500); }
});

app.put('/api/support/:id/status', requireAuth, (req, res) => {
  try {
    db.prepare("UPDATE support_messages SET status=? WHERE id=?").run(req.body.status, req.params.id);
    ok(res, {});
  } catch (e) { fail(res, e.message, 500); }
});

app.post('/api/support', requireAuth, async (req, res) => {
  try {
    const { name, category, subject, content } = req.body;
    if (!subject || !content) return fail(res, '제목과 내용을 입력해주세요.');

    // 1. DB에 항상 저장
    const row = db.prepare(
      'INSERT INTO support_messages (name, category, subject, content) VALUES (?, ?, ?, ?)'
    ).run(name||'', category||'', subject, content);

    // 2. SMTP 설정 있으면 이메일도 발송 (선택적)
    let emailSent = false;
    try {
      emailSent = await sendSupportMail({ name, category, subject, content });
      if (emailSent) db.prepare('UPDATE support_messages SET email_sent=1 WHERE id=?').run(row.lastInsertRowid);
    } catch (mailErr) {
      console.error('메일 발송 실패:', mailErr.message);
    }

    ok(res, { emailSent });
  } catch (e) { fail(res, e.message, 500); }
});

// ──────────────────────────────────────────────
//  서버 시작
// ──────────────────────────────────────────────
app.listen(PORT, () => {
  console.log('');
  console.log('====================================');
  console.log(' E-자산통합관리 구동 완료!');
  console.log('====================================');
  console.log(` 브라우저에서 접속하세요:`);
  console.log(` http://localhost:${PORT}`);
  console.log('====================================');
  console.log('');
});
