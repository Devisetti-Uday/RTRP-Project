// Middleware to check if user is authenticated
const isAuthenticated = (req, res, next) => {
  if (req.session && req.session.userId) {
    return next();
  }
  return res.status(401).json({ success: false, message: 'Access denied. Please log in.' });
};

// Middleware to check if user is Admin
const isAdmin = (req, res, next) => {
  if (req.session && req.session.userId && req.session.role === 'Admin') {
    return next();
  }
  return res.status(403).json({ success: false, message: 'Access denied. Admin privileges required.' });
};

// Middleware to check session for page routes (redirects instead of JSON)
const requireLogin = (req, res, next) => {
  if (req.session && req.session.userId) {
    return next();
  }
  return res.redirect('/login.html');
};

module.exports = { isAuthenticated, isAdmin, requireLogin };
