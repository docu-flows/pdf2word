require('dotenv').config();
const express = require('express');
const cors = require('cors');
const mongoose = require('mongoose');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { execFile } = require('child_process');
const jwt = require('jsonwebtoken');
const PDFDocument = require('pdfkit');
const Conversion = require('./models/Conversion');
const Admin = require('./models/Admin');

const JWT_SECRET = process.env.JWT_SECRET || 'pdf2word-secret-key-2024';

// Python script for pdf2docx conversion
const CONVERT_SCRIPT = path.join(__dirname, 'convert.py');

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json());

// Ensure upload/output directories exist
const uploadDir = path.join(__dirname, 'uploads');
const outputDir = path.join(__dirname, 'outputs');
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir, { recursive: true });
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

// Multer config for PDF upload
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const uniqueName = Date.now() + '-' + file.originalname;
    cb(null, uniqueName);
  }
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new Error('Only PDF files are allowed'), false);
    }
  },
  limits: { fileSize: 50 * 1024 * 1024 } // 50MB limit
});

// Connect to MongoDB
mongoose.connect(process.env.MONGODB_URI)
  .then(() => { console.log('Connected to MongoDB'); seedAdmin(); })
  .catch(err => console.error('MongoDB connection error:', err));

// Upload PDF
app.post('/api/upload', upload.single('pdf'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No PDF file uploaded' });
    }

    const clientId = req.headers['x-client-id'];
    if (!clientId) {
      return res.status(400).json({ error: 'Client ID required' });
    }

    const conversion = new Conversion({
      clientId,
      originalName: req.file.originalname,
      pdfPath: req.file.path,
      fileSize: req.file.size
    });
    await conversion.save();

    res.json({ id: conversion._id, originalName: conversion.originalName });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Start conversion
app.post('/api/convert/:id', async (req, res) => {
  try {
    const conversion = await Conversion.findById(req.params.id);
    if (!conversion) {
      return res.status(404).json({ error: 'Conversion not found' });
    }

    conversion.status = 'processing';
    conversion.progress = 10;
    await conversion.save();

    // Run conversion in background
    convertPdfToWord(conversion).catch(err => {
      console.error('Conversion error:', err);
    });

    res.json({ message: 'Conversion started', id: conversion._id });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Check conversion status
app.get('/api/status/:id', async (req, res) => {
  try {
    const conversion = await Conversion.findById(req.params.id);
    if (!conversion) {
      return res.status(404).json({ error: 'Conversion not found' });
    }
    res.json({
      status: conversion.status,
      progress: conversion.progress,
      originalName: conversion.originalName,
      error: conversion.error
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Download converted Word file
app.get('/api/download/:id', async (req, res) => {
  try {
    const conversion = await Conversion.findById(req.params.id);
    if (!conversion) {
      return res.status(404).json({ error: 'Conversion not found' });
    }
    if (conversion.status !== 'completed') {
      return res.status(400).json({ error: 'Conversion not completed yet' });
    }

    const wordFileName = conversion.originalName.replace(/\.pdf$/i, '.docx');
    res.download(conversion.wordPath, wordFileName);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Conversion history
app.get('/api/history', async (req, res) => {
  try {
    const clientId = req.headers['x-client-id'];
    if (!clientId) {
      return res.status(400).json({ error: 'Client ID required' });
    }

    const conversions = await Conversion.find({ clientId })
      .sort({ createdAt: -1 })
      .limit(20)
      .select('originalName status progress createdAt fileSize');
    res.json(conversions);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Auth middleware
function authMiddleware(req, res, next) {
  const token = req.headers.authorization?.replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token provided' });
  try {
    req.admin = jwt.verify(token, JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: 'Invalid token' });
  }
}

// Seed default admin on startup
async function seedAdmin() {
  const count = await Admin.countDocuments();
  if (count === 0) {
    await Admin.create({ username: 'adminpdf2w', password: 'w!n#23^74', name: 'Administrator' });
    //console.log('Default admin created (admin / admin123)');
  }
}

// Admin login
app.post('/api/admin/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    const admin = await Admin.findOne({ username });
    if (!admin || !(await admin.comparePassword(password))) {
      return res.status(401).json({ error: 'Invalid username or password' });
    }
    const token = jwt.sign({ id: admin._id, username: admin.username }, JWT_SECRET, { expiresIn: '8h' });
    res.json({ token, name: admin.name });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Verify token
app.get('/api/admin/me', authMiddleware, (req, res) => {
  res.json({ username: req.admin.username });
});

// Admin: Usage report
app.get('/api/admin/report', authMiddleware, async (req, res) => {
  try {
    const { from, to } = req.query;
    const match = {};
    if (from || to) {
      match.createdAt = {};
      if (from) match.createdAt.$gte = new Date(from);
      if (to) match.createdAt.$lte = new Date(to + 'T23:59:59.999Z');
    }

    // Summary stats
    const [stats] = await Conversion.aggregate([
      { $match: match },
      { $group: {
        _id: null,
        total: { $sum: 1 },
        completed: { $sum: { $cond: [{ $eq: ['$status', 'completed'] }, 1, 0] } },
        failed: { $sum: { $cond: [{ $eq: ['$status', 'failed'] }, 1, 0] } },
        totalSize: { $sum: '$fileSize' }
      }}
    ]);

    // Daily breakdown
    const daily = await Conversion.aggregate([
      { $match: match },
      { $group: {
        _id: { $dateToString: { format: '%Y-%m-%d', date: '$createdAt' } },
        count: { $sum: 1 },
        completed: { $sum: { $cond: [{ $eq: ['$status', 'completed'] }, 1, 0] } },
        failed: { $sum: { $cond: [{ $eq: ['$status', 'failed'] }, 1, 0] } },
        totalSize: { $sum: '$fileSize' }
      }},
      { $sort: { _id: -1 } },
      { $limit: 30 }
    ]);

    // Unique clients
    const clients = await Conversion.aggregate([
      { $match: match },
      { $group: { _id: '$clientId' } },
      { $count: 'total' }
    ]);

    // Recent conversions (all clients)
    const recent = await Conversion.find(match)
      .sort({ createdAt: -1 })
      .limit(50)
      .select('clientId originalName status fileSize createdAt');

    res.json({
      summary: stats || { total: 0, completed: 0, failed: 0, totalSize: 0 },
      uniqueClients: clients[0]?.total || 0,
      daily,
      recent
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Admin: Change password
app.post('/api/admin/change-password', authMiddleware, async (req, res) => {
  try {
    const { currentPassword, newPassword } = req.body;
    const admin = await Admin.findById(req.admin.id);
    if (!(await admin.comparePassword(currentPassword))) {
      return res.status(400).json({ error: 'Current password is incorrect' });
    }
    admin.password = newPassword;
    await admin.save();
    res.json({ message: 'Password changed successfully' });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Admin: Download report as PDF
app.get('/api/admin/report/pdf', authMiddleware, async (req, res) => {
  try {
    const { from, to } = req.query;
    const match = {};
    if (from || to) {
      match.createdAt = {};
      if (from) match.createdAt.$gte = new Date(from);
      if (to) match.createdAt.$lte = new Date(to + 'T23:59:59.999Z');
    }

    const [stats] = await Conversion.aggregate([
      { $match: match },
      { $group: {
        _id: null,
        total: { $sum: 1 },
        completed: { $sum: { $cond: [{ $eq: ['$status', 'completed'] }, 1, 0] } },
        failed: { $sum: { $cond: [{ $eq: ['$status', 'failed'] }, 1, 0] } },
        totalSize: { $sum: '$fileSize' }
      }}
    ]);

    const daily = await Conversion.aggregate([
      { $match: match },
      { $group: {
        _id: { $dateToString: { format: '%Y-%m-%d', date: '$createdAt' } },
        count: { $sum: 1 },
        completed: { $sum: { $cond: [{ $eq: ['$status', 'completed'] }, 1, 0] } },
        failed: { $sum: { $cond: [{ $eq: ['$status', 'failed'] }, 1, 0] } },
        totalSize: { $sum: '$fileSize' }
      }},
      { $sort: { _id: -1 } },
      { $limit: 30 }
    ]);

    const clients = await Conversion.aggregate([
      { $match: match },
      { $group: { _id: '$clientId' } },
      { $count: 'total' }
    ]);

    const recent = await Conversion.find(match)
      .sort({ createdAt: -1 })
      .limit(50)
      .select('clientId originalName status fileSize createdAt');

    const summary = stats || { total: 0, completed: 0, failed: 0, totalSize: 0 };
    const uniqueClients = clients[0]?.total || 0;

    const fmtSize = (bytes) => {
      if (!bytes) return '0 B';
      if (bytes < 1024) return bytes + ' B';
      if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
      if (bytes < 1073741824) return (bytes / 1048576).toFixed(1) + ' MB';
      return (bytes / 1073741824).toFixed(2) + ' GB';
    };

    // Build PDF
    const doc = new PDFDocument({ size: 'A4', margin: 40 });
    const dateLabel = from || to ? `${from || '...'} to ${to || '...'}` : 'All time';

    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="usage-report-${new Date().toISOString().slice(0,10)}.pdf"`);
    doc.pipe(res);

    // Title
    doc.fontSize(20).font('Helvetica-Bold').text('PDF to Word - Usage Report', { align: 'center' });
    doc.moveDown(0.3);
    doc.fontSize(11).font('Helvetica').fillColor('#666').text(`Period: ${dateLabel}`, { align: 'center' });
    doc.text(`Generated: ${new Date().toLocaleString()}`, { align: 'center' });
    doc.moveDown(1);

    // Summary
    doc.fillColor('#000').fontSize(14).font('Helvetica-Bold').text('Summary');
    doc.moveDown(0.3);
    doc.fontSize(10).font('Helvetica');

    const summaryY = doc.y;
    const colW = 125;
    const summaryData = [
      { label: 'Total Conversions', value: String(summary.total) },
      { label: 'Completed', value: String(summary.completed) },
      { label: 'Failed', value: String(summary.failed) },
      { label: 'Unique Clients', value: String(uniqueClients) }
    ];
    summaryData.forEach((item, i) => {
      const x = 40 + i * colW;
      doc.save();
      doc.roundedRect(x, summaryY, colW - 8, 50, 4).fill('#f8f9fb').stroke('#e0e3e8');
      doc.fillColor('#333').fontSize(18).font('Helvetica-Bold').text(item.value, x, summaryY + 8, { width: colW - 8, align: 'center' });
      doc.fillColor('#888').fontSize(8).font('Helvetica').text(item.label.toUpperCase(), x, summaryY + 32, { width: colW - 8, align: 'center' });
      doc.restore();
    });
    doc.y = summaryY + 60;
    doc.fillColor('#666').fontSize(10).font('Helvetica').text(`Total file size processed: ${fmtSize(summary.totalSize)}`);
    doc.moveDown(1);

    // Daily Breakdown Table
    if (daily.length > 0) {
      doc.fillColor('#000').fontSize(14).font('Helvetica-Bold').text('Daily Breakdown');
      doc.moveDown(0.5);

      const tableX = 40;
      const cols = [100, 70, 80, 70, 100];
      const headers = ['Date', 'Total', 'Completed', 'Failed', 'Size'];
      let ty = doc.y;

      // Header row
      doc.rect(tableX, ty, cols.reduce((a, b) => a + b, 0), 20).fill('#f0f2f5');
      doc.fillColor('#555').fontSize(9).font('Helvetica-Bold');
      let cx = tableX;
      headers.forEach((h, i) => {
        doc.text(h, cx + 6, ty + 5, { width: cols[i] - 12 });
        cx += cols[i];
      });
      ty += 20;

      // Data rows
      doc.font('Helvetica').fillColor('#333').fontSize(9);
      daily.forEach((d, idx) => {
        if (ty > 750) { doc.addPage(); ty = 40; }
        if (idx % 2 === 0) doc.rect(tableX, ty, cols.reduce((a, b) => a + b, 0), 18).fill('#fafafa');
        doc.fillColor('#333');
        cx = tableX;
        const row = [d._id, String(d.count), String(d.completed), String(d.failed), fmtSize(d.totalSize)];
        row.forEach((val, i) => {
          doc.text(val, cx + 6, ty + 4, { width: cols[i] - 12 });
          cx += cols[i];
        });
        ty += 18;
      });
      doc.y = ty;
      doc.moveDown(1);
    }

    // Recent Conversions Table
    if (recent.length > 0) {
      if (doc.y > 600) doc.addPage();
      doc.fillColor('#000').fontSize(14).font('Helvetica-Bold').text('Recent Conversions');
      doc.moveDown(0.5);

      const tableX = 40;
      const cols = [160, 70, 70, 80, 120];
      const headers = ['File', 'Status', 'Size', 'Client', 'Date'];
      let ty = doc.y;

      doc.rect(tableX, ty, cols.reduce((a, b) => a + b, 0), 20).fill('#f0f2f5');
      doc.fillColor('#555').fontSize(9).font('Helvetica-Bold');
      let cx = tableX;
      headers.forEach((h, i) => {
        doc.text(h, cx + 6, ty + 5, { width: cols[i] - 12 });
        cx += cols[i];
      });
      ty += 20;

      doc.font('Helvetica').fontSize(8);
      recent.forEach((item, idx) => {
        if (ty > 750) { doc.addPage(); ty = 40; }
        if (idx % 2 === 0) doc.rect(tableX, ty, cols.reduce((a, b) => a + b, 0), 18).fill('#fafafa');
        doc.fillColor('#333');
        cx = tableX;
        const fname = item.originalName.length > 28 ? item.originalName.slice(0, 28) + '...' : item.originalName;
        const row = [fname, item.status, fmtSize(item.fileSize), (item.clientId || '').slice(0, 8) + '...', new Date(item.createdAt).toLocaleString()];
        row.forEach((val, i) => {
          doc.text(val, cx + 6, ty + 4, { width: cols[i] - 12 });
          cx += cols[i];
        });
        ty += 18;
      });
    }

    doc.end();
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Serve React build in production
const publicDir = path.join(__dirname, 'public');
if (fs.existsSync(publicDir)) {
  app.use(express.static(publicDir));
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api')) return next();
    res.sendFile(path.join(publicDir, 'index.html'));
  });
}

// PDF to Word conversion using pdf2docx (preserves formatting)
// Keeps: fonts, font sizes, bold/italic/underline, colors, tables, images, layout
async function convertPdfToWord(conversion) {
  try {
    conversion.progress = 20;
    await conversion.save();

    const outputPath = path.join(outputDir, `${conversion._id}.docx`);

    // Use Python pdf2docx for high-fidelity conversion
    await new Promise((resolve, reject) => {
      conversion.progress = 40;
      conversion.save();

      execFile('python3', [CONVERT_SCRIPT, conversion.pdfPath, outputPath],
        { timeout: 180000 },
        (error, stdout, stderr) => {
          if (error) {
            reject(new Error(`Conversion failed: ${stderr || error.message}`));
            return;
          }
          if (stdout.trim() === 'OK') {
            resolve();
          } else {
            reject(new Error(`Unexpected output: ${stdout} ${stderr}`));
          }
        }
      );
    });

    conversion.progress = 90;
    await conversion.save();

    if (!fs.existsSync(outputPath)) {
      throw new Error('Converted file not found');
    }

    conversion.wordPath = outputPath;
    conversion.status = 'completed';
    conversion.progress = 100;
    await conversion.save();

  } catch (error) {
    conversion.status = 'failed';
    conversion.error = error.message;
    await conversion.save();
  }
}

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
