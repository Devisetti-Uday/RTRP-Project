const User = require('../models/User');
const Enquiry = require('../models/Enquiry');

// @desc    Get all users (excluding password, id)
// @route   GET /api/admin/users
// @access  Admin
const getUsers = async (req, res) => {
  try {
    const users = await User.find({})
      .select('name email college role created_at')
      .sort({ created_at: -1 });

    return res.status(200).json({ success: true, count: users.length, users });
  } catch (error) {
    console.error('Get users error:', error);
    return res.status(500).json({ success: false, message: 'Server error.' });
  }
};

// @desc    Get all enquiries
// @route   GET /api/admin/enquiries
// @access  Admin
const getEnquiries = async (req, res) => {
  try {
    const enquiries = await Enquiry.find({}).sort({ created_at: -1 });
    return res.status(200).json({ success: true, count: enquiries.length, enquiries });
  } catch (error) {
    console.error('Get enquiries error:', error);
    return res.status(500).json({ success: false, message: 'Server error.' });
  }
};

// @desc    Update enquiry status
// @route   PATCH /api/admin/enquiries/:id
// @access  Admin
const updateEnquiryStatus = async (req, res) => {
  try {
    const { status } = req.body;

    if (!['Pending', 'Completed'].includes(status)) {
      return res.status(400).json({ success: false, message: 'Invalid status value.' });
    }

    const enquiry = await Enquiry.findByIdAndUpdate(
      req.params.id,
      { status },
      { new: true, runValidators: true }
    );

    if (!enquiry) {
      return res.status(404).json({ success: false, message: 'Enquiry not found.' });
    }

    return res.status(200).json({ success: true, message: 'Status updated.', enquiry });
  } catch (error) {
    console.error('Update enquiry error:', error);
    return res.status(500).json({ success: false, message: 'Server error.' });
  }
};

// @desc    Get analytics summary
// @route   GET /api/admin/analytics
// @access  Admin
const getAnalytics = async (req, res) => {
  try {
    const [totalUsers, totalEnquiries, pendingEnquiries, completedEnquiries] = await Promise.all([
      User.countDocuments({}),
      Enquiry.countDocuments({}),
      Enquiry.countDocuments({ status: 'Pending' }),
      Enquiry.countDocuments({ status: 'Completed' })
    ]);

    // Role distribution
    const roleDistribution = await User.aggregate([
      { $group: { _id: '$role', count: { $sum: 1 } } }
    ]);

    return res.status(200).json({
      success: true,
      analytics: {
        totalUsers,
        totalEnquiries,
        pendingEnquiries,
        completedEnquiries,
        roleDistribution
      }
    });
  } catch (error) {
    console.error('Analytics error:', error);
    return res.status(500).json({ success: false, message: 'Server error.' });
  }
};

module.exports = { getUsers, getEnquiries, updateEnquiryStatus, getAnalytics };
