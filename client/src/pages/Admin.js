import React, { useState, useEffect, useCallback } from 'react';
import axios from 'axios';

const API_URL = process.env.REACT_APP_API_URL || '/api';

function Admin({ onLogout }) {
  const [report, setReport] = useState(null);
  const [loading, setLoading] = useState(true);
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const adminName = localStorage.getItem('admin_name') || 'Admin';

  const getAuthHeader = () => ({
    headers: { Authorization: `Bearer ${localStorage.getItem('admin_token')}` }
  });

  const fetchReport = useCallback(async () => {
    setLoading(true);
    try {
      const params = new URLSearchParams();
      if (dateFrom) params.append('from', dateFrom);
      if (dateTo) params.append('to', dateTo);
      const res = await axios.get(`${API_URL}/admin/report?${params}`, getAuthHeader());
      setReport(res.data);
    } catch (err) {
      if (err.response?.status === 401) onLogout();
    } finally {
      setLoading(false);
    }
  }, [dateFrom, dateTo, onLogout]);

  useEffect(() => {
    fetchReport();
  }, [fetchReport]);

  const handleLogout = () => {
    localStorage.removeItem('admin_token');
    localStorage.removeItem('admin_name');
    onLogout();
  };

  const handleDownloadPDF = async () => {
    try {
      const params = new URLSearchParams();
      if (dateFrom) params.append('from', dateFrom);
      if (dateTo) params.append('to', dateTo);
      const res = await axios.get(`${API_URL}/admin/report/pdf?${params}`, {
        ...getAuthHeader(),
        responseType: 'blob'
      });
      const url = window.URL.createObjectURL(new Blob([res.data], { type: 'application/pdf' }));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `usage-report-${new Date().toISOString().slice(0, 10)}.pdf`);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      if (err.response?.status === 401) onLogout();
    }
  };

  const handleDownloadFile = async (id, originalName) => {
    try {
      const res = await axios.get(`${API_URL}/download/${id}`, { responseType: 'blob' });
      const url = window.URL.createObjectURL(new Blob([res.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', originalName.replace(/\.pdf$/i, '.docx'));
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Download failed');
    }
  };

  const formatSize = (bytes) => {
    if (!bytes) return '0 B';
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
    if (bytes < 1073741824) return (bytes / 1048576).toFixed(1) + ' MB';
    return (bytes / 1073741824).toFixed(2) + ' GB';
  };

  const formatDate = (d) => new Date(d).toLocaleString();

  return (
    <div className="app">
      <div className="container admin-container">
        {/* Header */}
        <div className="admin-header">
          <div>
            <h1 className="title" style={{ textAlign: 'left', fontSize: 22 }}>Admin Dashboard</h1>
            <p className="subtitle" style={{ textAlign: 'left', marginBottom: 0 }}>Welcome, {adminName}</p>
          </div>
          <div className="admin-header-actions">
            <a href="/" className="btn btn-browse">Home</a>
            <button className="btn btn-logout" onClick={handleLogout}>Logout</button>
          </div>
        </div>

        {/* Date Filter */}
        <div className="filter-section">
          <h3>Report Usage</h3>
          <div className="filter-row">
            <div className="form-group">
              <label>From</label>
              <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} />
            </div>
            <div className="form-group">
              <label>To</label>
              <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} />
            </div>
            <button className="btn btn-browse" onClick={fetchReport}>Filter</button>
            <button className="btn btn-pdf-download" onClick={handleDownloadPDF}>Download PDF</button>
          </div>
        </div>

        {loading ? (
          <div className="loading-text">Loading report...</div>
        ) : report && (
          <>
            {/* Summary Cards */}
            <div className="stat-cards">
              <div className="stat-card">
                <div className="stat-number">{report.summary.total}</div>
                <div className="stat-label">Total Conversions</div>
              </div>
              <div className="stat-card stat-success">
                <div className="stat-number">{report.summary.completed}</div>
                <div className="stat-label">Completed</div>
              </div>
              <div className="stat-card stat-danger">
                <div className="stat-number">{report.summary.failed}</div>
                <div className="stat-label">Failed</div>
              </div>
              <div className="stat-card stat-info">
                <div className="stat-number">{report.uniqueClients}</div>
                <div className="stat-label">Unique Clients</div>
              </div>
            </div>

            <div className="stat-total-size">
              Total file size processed: <strong>{formatSize(report.summary.totalSize)}</strong>
            </div>

            {/* Daily Breakdown */}
            {report.daily.length > 0 && (
              <div className="report-section">
                <h3>Daily Breakdown</h3>
                <div className="report-table-wrap">
                  <table className="report-table">
                    <thead>
                      <tr>
                        <th>Date</th>
                        <th>Total</th>
                        <th>Completed</th>
                        <th>Failed</th>
                        <th>Size</th>
                      </tr>
                    </thead>
                    <tbody>
                      {report.daily.map((d) => (
                        <tr key={d._id}>
                          <td>{d._id}</td>
                          <td>{d.count}</td>
                          <td className="text-success">{d.completed}</td>
                          <td className="text-danger">{d.failed}</td>
                          <td>{formatSize(d.totalSize)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Recent Conversions */}
            {report.recent.length > 0 && (
              <div className="report-section">
                <h3>Recent Conversions (All Clients)</h3>
                <div className="report-table-wrap">
                  <table className="report-table">
                    <thead>
                      <tr>
                        <th>File</th>
                        <th>Status</th>
                        <th>Size</th>
                        <th>Client</th>
                        <th>Date</th>
                        <th>Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      {report.recent.map((item) => (
                        <tr key={item._id}>
                          <td className="cell-filename">{item.originalName}</td>
                          <td>
                            <span className={`history-status status-${item.status}`}>
                              {item.status}
                            </span>
                          </td>
                          <td>{formatSize(item.fileSize)}</td>
                          <td className="cell-client">{item.clientId?.slice(0, 8)}...</td>
                          <td className="cell-date">{formatDate(item.createdAt)}</td>
                          <td>
                            {item.status === 'completed' && (
                              <button
                                className="btn btn-table-download"
                                onClick={() => handleDownloadFile(item._id, item.originalName)}
                              >
                                Download
                              </button>
                            )}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

export default Admin;
