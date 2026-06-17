const express = require('express');
const router = express.Router();
const { body } = require('express-validator');
const { submitEnquiry } = require('../controllers/enquiryController');
const { isAuthenticated } = require('../middleware/auth');

const enquiryValidation = [
  body('name').trim().notEmpty().withMessage('Name is required'),
  body('email').isEmail().withMessage('Valid email is required').normalizeEmail(),
  body('phone').trim().matches(/^[0-9+\-\s()]{7,15}$/).withMessage('Valid phone number required'),
  body('college').trim().notEmpty().withMessage('College name is required'),
  body('role').isIn(['Student', 'Faculty', 'Researcher', 'Administrator', 'Other']).withMessage('Invalid role'),
  body('message').trim().notEmpty().withMessage('Message is required').isLength({ max: 2000 })
];

router.post('/', isAuthenticated, enquiryValidation, submitEnquiry);

module.exports = router;
