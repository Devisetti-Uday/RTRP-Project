const express = require('express');
const router = express.Router();
const { getUsers, getEnquiries, updateEnquiryStatus, getAnalytics } = require('../controllers/adminController');
const { isAdmin } = require('../middleware/auth');

// All routes protected by isAdmin middleware
router.get('/users', isAdmin, getUsers);
router.get('/enquiries', isAdmin, getEnquiries);
router.patch('/enquiries/:id', isAdmin, updateEnquiryStatus);
router.get('/analytics', isAdmin, getAnalytics);

module.exports = router;
