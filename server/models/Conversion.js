const mongoose = require('mongoose');

const conversionSchema = new mongoose.Schema({
  clientId: { type: String, required: true, index: true },
  originalName: { type: String, required: true },
  pdfPath: { type: String, required: true },
  wordPath: { type: String, default: null },
  status: {
    type: String,
    enum: ['pending', 'processing', 'completed', 'failed'],
    default: 'pending'
  },
  progress: { type: Number, default: 0 },
  fileSize: { type: Number },
  error: { type: String, default: null }
}, { timestamps: true });

module.exports = mongoose.model('Conversion', conversionSchema);
