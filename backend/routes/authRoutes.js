const express = require('express');
const router = express.Router();
const { body } = require('express-validator');
const { signup, login, logout, getMe } = require('../controllers/authController');
const { isAuthenticated } = require('../middleware/auth');

// Signup validation
const signupValidation = [
  body('name').trim().notEmpty().withMessage('Name is required').isLength({ max: 100 }),
  body('email').isEmail().withMessage('Valid email is required').normalizeEmail(),
  body('password').isLength({ min: 6 }).withMessage('Password must be at least 6 characters'),
  body('confirmPassword').custom((value, { req }) => {
    if (value !== req.body.password) throw new Error('Passwords do not match');
    return true;
  }),
  body('college').trim().notEmpty().withMessage('College name is required'),
  body('role').optional().isIn(['Student', 'Faculty', 'Researcher']).withMessage('Invalid role')
];

// Login validation
const loginValidation = [
  body('email').isEmail().withMessage('Valid email is required').normalizeEmail(),
  body('password').notEmpty().withMessage('Password is required')
];

router.post('/signup', signupValidation, signup);
router.post('/login', loginValidation, login);
router.post('/logout', isAuthenticated, logout);
router.get('/me', isAuthenticated, getMe);

module.exports = router;
