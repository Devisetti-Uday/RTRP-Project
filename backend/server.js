require('dotenv').config();
const express = require('express');
const session = require('express-session');
const MongoStore = require('connect-mongo');
const cors = require('cors');
const path = require('path');
const connectDB = require('./config/db');

// Connect to MongoDB
connectDB();

const app = express();

// ─── Middleware ───────────────────────────────────────────────────────────────
app.use(express.json({ limit: '10kb' }));
app.use(express.urlencoded({ extended: true, limit: '10kb' }));

// CORS
app.use(cors({
  origin: process.env.FRONTEND_URL || 'http://localhost:5000',
  credentials: true
}));

// Session
app.use(session({
  secret: process.env.SESSION_SECRET || 'fallback_secret_key',
  resave: false,
  saveUninitialized: false,
  store: MongoStore.create({
    mongoUrl: process.env.MONGO_URI,
    collectionName: 'sessions'
  }),
  cookie: {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    maxAge: 24 * 60 * 60 * 1000 // 24 hours
  }
}));

// ─── API Routes ───────────────────────────────────────────────────────────────
app.use('/api/auth', require('./routes/authRoutes'));
app.use('/api/enquiry', require('./routes/enquiryRoutes'));
app.use('/api/admin', require('./routes/adminRoutes'));

// ─── Serve Static Frontend ────────────────────────────────────────────────────
const frontendPath = path.join(__dirname, '../frontend');
app.use(express.static(frontendPath));

// Route: serve specific HTML pages
app.get('/login', (req, res) => res.sendFile(path.join(frontendPath, 'views', 'login.html')));
app.get('/signup', (req, res) => res.sendFile(path.join(frontendPath, 'views', 'signup.html')));
app.get('/loading', (req, res) => res.sendFile(path.join(frontendPath, 'views', 'loading.html')));
app.get('/dashboard', (req, res) => res.sendFile(path.join(frontendPath, 'views', 'dashboard.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(frontendPath, 'views', 'admin.html')));

// Default route
app.get('/', (req, res) => res.redirect('/login'));

// ─── 404 Handler ─────────────────────────────────────────────────────────────
app.use((req, res) => {
  res.status(404).json({ success: false, message: 'Route not found.' });
});

// ─── Global Error Handler ─────────────────────────────────────────────────────
app.use((err, req, res, next) => {
  console.error('Server Error:', err.stack);
  res.status(500).json({ success: false, message: 'Internal server error.' });
});

// ─── Start Server ─────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`\n🚀 Server running on http://localhost:${PORT}`);
  console.log(`📊 College Result Analytics Dashboard`);
  console.log(`🔐 Environment: ${process.env.NODE_ENV || 'development'}\n`);
});
