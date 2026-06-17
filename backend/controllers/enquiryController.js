const Enquiry = require('../models/Enquiry');
const nodemailer = require('nodemailer');
const { validationResult } = require('express-validator');

// Create transporter
const createTransporter = () => {
  return nodemailer.createTransport({
    host: process.env.EMAIL_HOST,
    port: parseInt(process.env.EMAIL_PORT),
    secure: process.env.EMAIL_PORT == 465,
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    }
  });
};

// Send auto-response email
const sendAutoResponse = async (enquiry) => {
  try {
    const transporter = createTransporter();

    const mailOptions = {
      from: process.env.EMAIL_FROM || 'College Analytics <noreply@collegeanalytics.com>',
      to: enquiry.email,
      subject: '✅ Enquiry Received — College Result Analytics',
      html: `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <style>
            body { font-family: 'Segoe UI', Arial, sans-serif; background: #f4f6f9; margin: 0; padding: 0; }
            .container { max-width: 600px; margin: 30px auto; background: #fff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 20px rgba(0,0,0,0.1); }
            .header { background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%); padding: 40px 30px; text-align: center; }
            .header h1 { color: #00d4ff; margin: 0; font-size: 22px; letter-spacing: 2px; }
            .header p { color: #a0c4ff; margin: 8px 0 0; font-size: 13px; }
            .body { padding: 40px 30px; }
            .greeting { font-size: 18px; color: #1a1a2e; font-weight: 600; margin-bottom: 16px; }
            .message { color: #555; line-height: 1.7; font-size: 15px; }
            .detail-box { background: #f0f7ff; border-left: 4px solid #0f3460; border-radius: 0 8px 8px 0; padding: 20px; margin: 24px 0; }
            .detail-box p { margin: 6px 0; color: #333; font-size: 14px; }
            .detail-box strong { color: #0f3460; }
            .status-badge { display: inline-block; background: #fff3cd; color: #856404; border: 1px solid #ffc107; padding: 4px 12px; border-radius: 20px; font-size: 13px; font-weight: 600; }
            .footer { background: #f8f9fa; padding: 20px 30px; text-align: center; border-top: 1px solid #eee; }
            .footer p { color: #888; font-size: 12px; margin: 0; }
          </style>
        </head>
        <body>
          <div class="container">
            <div class="header">
              <h1>🎓 COLLEGE RESULT ANALYTICS</h1>
              <p>Academic Intelligence Platform</p>
            </div>
            <div class="body">
              <p class="greeting">Dear ${enquiry.name},</p>
              <p class="message">
                Thank you for contacting the <strong>College Result Analytics Team</strong>.
                We have successfully received your enquiry and our team will review it shortly.
              </p>
              <div class="detail-box">
                <p><strong>Enquiry Details:</strong></p>
                <p>📧 Email: ${enquiry.email}</p>
                <p>📞 Phone: ${enquiry.phone}</p>
                <p>🏫 College: ${enquiry.college}</p>
                <p>👤 Role: ${enquiry.role}</p>
                <p>📋 Status: <span class="status-badge">⏳ Pending</span></p>
              </div>
              <p class="message">
                Our team typically responds within <strong>24–48 business hours</strong>.
                We look forward to assisting you with your academic analytics needs.
              </p>
              <p class="message">Warm regards,<br><strong>College Result Analytics Team</strong></p>
            </div>
            <div class="footer">
              <p>This is an automated response. Please do not reply to this email.</p>
              <p>© ${new Date().getFullYear()} College Result Analytics. All rights reserved.</p>
            </div>
          </div>
        </body>
        </html>
      `
    };

    await transporter.sendMail(mailOptions);
    console.log(`✅ Auto-response email sent to ${enquiry.email}`);
  } catch (error) {
    console.error('❌ Email sending failed:', error.message);
    // Don't throw — email failure shouldn't break the enquiry submission
  }
};

// @desc    Submit enquiry
// @route   POST /api/enquiry
// @access  Private
const submitEnquiry = async (req, res) => {
  try {
    const errors = validationResult(req);
    if (!errors.isEmpty()) {
      return res.status(400).json({ success: false, errors: errors.array() });
    }

    const { name, email, phone, college, role, message } = req.body;

    const enquiry = await Enquiry.create({ name, email, phone, college, role, message });

    // Send auto-response email (non-blocking)
    sendAutoResponse(enquiry);

    return res.status(201).json({
      success: true,
      message: 'Enquiry submitted successfully! You will receive a confirmation email shortly.',
      enquiry: { id: enquiry._id, status: enquiry.status }
    });
  } catch (error) {
    console.error('Enquiry submission error:', error);
    return res.status(500).json({ success: false, message: 'Server error. Please try again.' });
  }
};

module.exports = { submitEnquiry };
