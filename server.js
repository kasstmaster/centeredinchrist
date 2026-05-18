const express = require('express');
const session = require('express-session');
const bcrypt = require('bcrypt');
const fs = require('fs');
const path = require('path');
const cron = require('node-cron');
const nodemailer = require('nodemailer');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const csrf = require('csurf');

const app = express();
const PORT = process.env.PORT || 3000;
const OWNER_EMAIL = 'daveraudman@gmail.com';
const OWNER_KEY = process.env.OWNER_KEY;
const SESSION_SECRET = process.env.SESSION_SECRET;
const STAFF_DATA_PATH = path.join(__dirname, 'data', 'staff-auth.json');


if (process.env.NODE_ENV === 'production' && !SESSION_SECRET) {
  throw new Error('SESSION_SECRET must be set in production.');
}

fs.mkdirSync(path.dirname(STAFF_DATA_PATH), { recursive: true });

function loadStaffData() {
  if (!fs.existsSync(STAFF_DATA_PATH)) {
    return { passwordHash: null, sessionVersion: 1 };
  }
  return JSON.parse(fs.readFileSync(STAFF_DATA_PATH, 'utf8'));
}

function saveStaffData(data) {
  fs.writeFileSync(STAFF_DATA_PATH, JSON.stringify(data, null, 2));
}

function getTransporter() {
  const { SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_SECURE, SMTP_FROM } = process.env;
  if (!SMTP_HOST || !SMTP_PORT || !SMTP_USER || !SMTP_PASS || !SMTP_FROM) return null;
  return nodemailer.createTransport({
    host: SMTP_HOST,
    port: Number(SMTP_PORT),
    secure: SMTP_SECURE === 'true',
    auth: { user: SMTP_USER, pass: SMTP_PASS },
  });
}

app.use(express.json());
app.use(helmet());
app.use(session({
  name: 'cic_staff_session',
  secret: SESSION_SECRET || 'dev-only-session-secret',
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    sameSite: 'lax',
    secure: process.env.NODE_ENV === 'production',
    maxAge: 1000 * 60 * 60 * 24 * 30,
  },
}));

app.use(express.static(__dirname));


const csrfProtection = csrf();
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: 'Too many login attempts, please try again later.' }
});


function requireAuth(req, res, next) {
  const data = loadStaffData();
  if (!req.session?.authenticated) return res.status(401).json({ error: 'Unauthorized' });
  if (req.session.sessionVersion !== data.sessionVersion) {
    req.session.destroy(() => {});
    return res.status(401).json({ error: 'Session expired' });
  }
  next();
}

app.get('/api/csrf-token', csrfProtection, (req, res) => {
  res.json({ csrfToken: req.csrfToken() });
});

app.post('/api/staff/login', loginLimiter, csrfProtection, async (req, res) => {
  const { password } = req.body || {};
  const data = loadStaffData();
  if (!data.passwordHash) return res.status(503).json({ error: 'Password not configured yet' });
  const ok = typeof password === 'string' && await bcrypt.compare(password, data.passwordHash);
  if (!ok) return res.status(401).json({ error: 'Invalid password' });

  req.session.authenticated = true;
  req.session.sessionVersion = data.sessionVersion;
  res.json({ ok: true });
});

app.post('/api/staff/logout', csrfProtection, (req, res) => {
  req.session.destroy(() => {
    res.clearCookie('cic_staff_session');
    res.json({ ok: true });
  });
});

app.get('/api/staff/session', (req, res) => {
  const data = loadStaffData();
  const authenticated = !!req.session?.authenticated && req.session.sessionVersion === data.sessionVersion;
  res.json({ authenticated });
});

app.get('/api/prayer-requests', requireAuth, (_req, res) => {
  res.json({
    title: 'Prayer Requests',
    intro: 'Please keep the following in prayer this week:',
    items: [
      'Spiritual growth and unity in our church family.',
      'Comfort and healing for anyone facing illness or grief.',
      'Open hearts in our local community.'
    ]
  });
});

app.post('/api/admin/password', csrfProtection, async (req, res) => {
  if (!OWNER_KEY || req.headers['x-owner-key'] !== OWNER_KEY) {
    return res.status(403).json({ error: 'Forbidden' });
  }

  const { newPassword } = req.body || {};
  if (typeof newPassword !== 'string' || newPassword.length < 12) {
    return res.status(400).json({ error: 'New password must be at least 12 characters' });
  }

  const data = loadStaffData();
  data.passwordHash = await bcrypt.hash(newPassword, 12);
  data.sessionVersion += 1;
  saveStaffData(data);

  res.json({ ok: true, message: 'Password updated and all sessions invalidated' });
});

async function sendMonthlyReminderEmail() {
  const transporter = getTransporter();
  if (!transporter) return;

  await transporter.sendMail({
    from: process.env.SMTP_FROM,
    to: OWNER_EMAIL,
    subject: 'Monthly reminder: change staff password',
    text: 'Please update the website staff password today.',
  });
}

cron.schedule('0 9 1 * *', async () => {
  try {
    await sendMonthlyReminderEmail();
    console.log('Monthly password reminder email sent.');
  } catch (error) {
    console.error('Failed to send monthly reminder email:', error.message);
  }
}, { timezone: 'America/Los_Angeles' });

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
