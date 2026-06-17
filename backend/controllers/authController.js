const User = require('../models/User');
const { validationResult } = require('express-validator');

// @desc    Register new user
// @route   POST /api/auth/signup
// @access  Public
const signup = async (req, res) => {
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ success: false, errors: errors.array() });
    }

    const { name, email, password, college, role } = req.body;

    // Check if user already exists
    const existingUser = await User.findOne({ email });
    if (existingUser) {
      return res.status(409).json({ success: false, message: 'An account with this email already exists.' });
    }

    // Create new user
    const user = await User.create({ name, email, password, college, role: role || 'Student' });

    // Start session
    req.session.userId = user._id;
    req.session.name = user.name;
    req.session.role = user.role;
    req.session.college = user.college;

    return res.status(201).json({
      success: true,
      message: 'Account created successfully!',
      user: { name: user.name, email: user.email, role: user.role, college: user.college }
    });
  } catch (error) {
    console.error('Signup error:', error);
    return res.status(500).json({ success: false, message: 'Server error. Please try again.' });
  }
};

// @desc    Login user
// @route   POST /api/auth/login
// @access  Public
const login = async (req, res) => {
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ success: false, errors: errors.array() });
    }

    const { email, password } = req.body;

    // Find user with password field
    const user = await User.findOne({ email }).select('+password');
    if (!user) {
      return res.status(401).json({ success: false, message: 'Invalid email or password.' });
    }

    // Verify password
    const isMatch = await user.comparePassword(password);
    if (!isMatch) {
      return res.status(401).json({ success: false, message: 'Invalid email or password.' });
    }

    // Start session
    req.session.userId = user._id;
    req.session.name = user.name;
    req.session.role = user.role;
    req.session.college = user.college;

    return res.status(200).json({
      success: true,
      message: 'Login successful!',
      user: { name: user.name, email: user.email, role: user.role, college: user.college }
    });
  } catch (error) {
    console.error('Login error:', error);
    return res.status(500).json({ success: false, message: 'Server error. Please try again.' });
  }
};

// @desc    Logout user
// @route   POST /api/auth/logout
// @access  Private
const logout = (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      return res.status(500).json({ success: false, message: 'Could not log out. Try again.' });
    }
    res.clearCookie('connect.sid');
    return res.status(200).json({ success: true, message: 'Logged out successfully.' });
  });
};

// @desc    Get current session user info
// @route   GET /api/auth/me
// @access  Private
const getMe = (req, res) => {
  if (req.session && req.session.userId) {
    return res.status(200).json({
      success: true,
      user: {
        id: req.session.userId,
        name: req.session.name,
        role: req.session.role,
        college: req.session.college
      }
    });
  }
  return res.status(401).json({ success: false, message: 'Not authenticated.' });
};

module.exports = { signup, login, logout, getMe };
