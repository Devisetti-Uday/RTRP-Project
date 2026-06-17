/**
 * Admin Seeder Script
 * Run: node seedAdmin.js
 * Creates a default Admin account in MongoDB
 */

require('dotenv').config();
const mongoose = require('mongoose');
const bcrypt = require('bcryptjs');

const MONGO_URI = process.env.MONGO_URI || 'mongodb+srv://devisettiuday_db_user:fN8VSxnOKdgjhnmj@cluster0.ybjz9hb.mongodb.net/?appName=Cluster0';

// Inline User schema for this script
const userSchema = new mongoose.Schema({
  name: String,
  email: { type: String, unique: true, lowercase: true },
  password: { type: String, select: false },
  college: String,
  role: String,
  created_at: { type: Date, default: Date.now }
});

const User = mongoose.model('User', userSchema);

async function seedAdmin() {
  try {
    await mongoose.connect(MONGO_URI);
    console.log('✅ Connected to MongoDB');

    const adminEmail = 'admin@college.edu';
    const adminPass  = 'admin123';

    // Check if admin already exists
    const existing = await User.findOne({ email: adminEmail });
    if (existing) {
      console.log('ℹ️  Admin already exists:', adminEmail);
      await mongoose.disconnect();
      return;
    }

    // Hash password
    const salt = await bcrypt.genSalt(12);
    const hashedPassword = await bcrypt.hash(adminPass, salt);

    await User.create({
      name: 'System Administrator',
      email: adminEmail,
      password: hashedPassword,
      college: 'College Result Analytics HQ',
      role: 'Admin'
    });

    console.log('\n🎉 Admin account created successfully!');
    console.log('─────────────────────────────────────');
    console.log('📧 Email:    admin@college.edu');
    console.log('🔑 Password: admin123');
    console.log('⚠️  Change this password in production!');
    console.log('─────────────────────────────────────\n');

    await mongoose.disconnect();
  } catch (err) {
    console.error('❌ Seeder error:', err.message);
    process.exit(1);
  }
}

seedAdmin();
