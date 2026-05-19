import React, { useState, useEffect, useRef, useCallback } from 'react';
import { BrowserRouter, Routes, Route, Navigate, Link } from 'react-router-dom';
import axios from 'axios';
import Login from './pages/Login';
import Admin from './pages/Admin';
import './App.css';

const API_URL = process.env.REACT_APP_API_URL || '/api';

// Generate a unique client ID per machine/browser (persists in localStorage)
function getClientId() {
  let id = localStorage.getItem('pdf2word_client_id');
  if (!id) {
    id = crypto.randomUUID();
    localStorage.setItem('pdf2word_client_id', id);
  }
  return id;
}
const CLIENT_ID = getClientId();

function Converter() {
  const [file, setFile] = useState(null);
  const [fileName, setFileName] = useState('');
  const [conversionId, setConversionId] = useState(null);
  const [status, setStatus] = useState('idle');
  const [progress, setProgress] = useState(0);
  const [showDownload, setShowDownload] = useState(false);
  const [error, setError] = useState('');
  const [history, setHistory] = useState([]);
  const fileInputRef = useRef(null);
  const pollRef = useRef(null);

  const fetchHistory = useCallback(async () => {
    try {
      const res = await axios.get(`${API_URL}/history`, {
        headers: { 'X-Client-Id': CLIENT_ID }
      });
      setHistory(res.data);
    } catch (err) {
      console.error('Failed to load history');
    }
  }, []);

  useEffect(() => {
    fetchHistory();
  }, [fetchHistory]);

  useEffect(() => {
    if (status === 'processing' && conversionId) {
      pollRef.current = setInterval(async () => {
        try {
          const res = await axios.get(`${API_URL}/status/${conversionId}`);
          setProgress(res.data.progress);

          if (res.data.status === 'completed') {
            setStatus('completed');
            setProgress(100);
            setShowDownload(true);
            clearInterval(pollRef.current);
            fetchHistory();
          } else if (res.data.status === 'failed') {
            setStatus('failed');
            setError(res.data.error || 'Conversion failed');
            clearInterval(pollRef.current);
          }
        } catch (err) {
          console.error('Polling error:', err);
        }
      }, 500);
    }

    return () => {
      if (pollRef.current) clearInterval(pollRef.current);
    };
  }, [status, conversionId, fetchHistory]);

  const handleFileChange = (e) => {
    const selected = e.target.files[0];
    if (selected && selected.type === 'application/pdf') {
      setFile(selected);
      setFileName(selected.name);
      setError('');
      setStatus('idle');
      setProgress(0);
      setShowDownload(false);
      setConversionId(null);
    } else {
      setError('Please select a valid PDF file');
    }
  };

  const handleBrowse = () => {
    fileInputRef.current.click();
  };

  const handleConvert = async () => {
    if (!file) {
      setError('Please select a PDF file first');
      return;
    }

    try {
      setStatus('uploading');
      setProgress(5);
      setError('');

      const formData = new FormData();
      formData.append('pdf', file);

      const uploadRes = await axios.post(`${API_URL}/upload`, formData, {
        headers: { 'Content-Type': 'multipart/form-data', 'X-Client-Id': CLIENT_ID }
      });

      const id = uploadRes.data.id;
      setConversionId(id);

      setStatus('processing');
      await axios.post(`${API_URL}/convert/${id}`);

    } catch (err) {
      setStatus('failed');
      setError(err.response?.data?.error || 'Upload failed');
    }
  };

  const handleDownload = async () => {
    try {
      const res = await axios.get(`${API_URL}/download/${conversionId}`, {
        responseType: 'blob'
      });

      const url = window.URL.createObjectURL(new Blob([res.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', fileName.replace(/\.pdf$/i, '.docx'));
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      setError('Download failed');
    }
  };

  const handleReset = () => {
    setFile(null);
    setFileName('');
    setConversionId(null);
    setStatus('idle');
    setProgress(0);
    setShowDownload(false);
    setError('');
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const formatFileSize = (bytes) => {
    if (!bytes) return '';
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / 1048576).toFixed(1) + ' MB';
  };

  const formatDate = (dateStr) => {
    return new Date(dateStr).toLocaleString();
  };

  return (
    <div className="app">
      <div className="container">
        <div className="page-header">
          <h1 className="title">PDF to Word Converter</h1>
           
        </div>
        <p className="subtitle">Convert your PDF files to editable Word documents</p>

        {/* File Browser */}
        <div className="upload-area">
          <input
            type="file"
            accept=".pdf"
            ref={fileInputRef}
            onChange={handleFileChange}
            hidden
          />

          <div className="file-icon">📄</div>

          {fileName ? (
            <div className="file-info">
              <p className="file-name">{fileName}</p>
              <p className="file-size">{file && formatFileSize(file.size)}</p>
            </div>
          ) : (
            <p className="upload-text">Select a PDF file to convert</p>
          )}

          <button
            className="btn btn-browse"
            onClick={handleBrowse}
            disabled={status === 'uploading' || status === 'processing'}
          >
            Browse PDF
          </button>
        </div>

        {/* Error */}
        {error && <div className="error-msg">{error}</div>}

        {/* Progress Bar */}
        {(status === 'uploading' || status === 'processing') && (
          <div className="progress-section">
            <div className="progress-bar">
              <div
                className="progress-fill"
                style={{ width: `${progress}%` }}
              />
            </div>
            <p className="progress-text">
              {status === 'uploading' ? 'Uploading...' : `Converting... ${progress}%`}
            </p>
          </div>
        )}

        {/* Convert Button */}
        <button
          className="btn btn-convert"
          onClick={handleConvert}
          disabled={!file || status === 'uploading' || status === 'processing'}
        >
          {status === 'processing' ? 'Converting...' : 'Start Convert'}
        </button>

        {/* Download Popup */}
        {showDownload && (
          <div className="popup-overlay" onClick={() => setShowDownload(false)}>
            <div className="popup" onClick={(e) => e.stopPropagation()}>
              <div className="popup-icon">✅</div>
              <h2>Conversion Complete!</h2>
              <p>Your Word document is ready to download.</p>
              <p className="popup-filename">{fileName.replace(/\.pdf$/i, '.docx')}</p>
              <div className="popup-buttons">
                <button className="btn btn-download" onClick={handleDownload}>
                  Download Word File
                </button>
                <button className="btn btn-new" onClick={() => { setShowDownload(false); handleReset(); }}>
                  Convert Another
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Conversion History */}
        {history.length > 0 && (
          <div className="history-section">
            <h3>Recent Conversions</h3>
            <div className="history-list">
              {history.map((item) => (
                <div key={item._id} className="history-item">
                  <span className="history-name">{item.originalName}</span>
                  <span className={`history-status status-${item.status}`}>
                    {item.status}
                  </span>
                  <span className="history-date">{formatDate(item.createdAt)}</span>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function ProtectedRoute({ children }) {
  const token = localStorage.getItem('admin_token');
  if (!token) return <Navigate to="/login" replace />;
  return children;
}

function App() {
  const [, setAuth] = useState(!!localStorage.getItem('admin_token'));

  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<Converter />} />
        <Route path="/login" element={
          <Login onLogin={() => { setAuth(true); window.location.href = '/admin'; }} />
        } />
        <Route path="/admin" element={
          <ProtectedRoute>
            <Admin onLogout={() => { setAuth(false); window.location.href = '/login'; }} />
          </ProtectedRoute>
        } />
      </Routes>
    </BrowserRouter>
  );
}

export default App;
