import { useState, useEffect } from 'react';
import ApproveRequestModal from './ApproveRequestModal';
import { RefreshCw, CheckCircle } from 'lucide-react';

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:4001/api';

export default function IndexingRequestsTable({ showToast }) {
  const [requests, setRequests] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedRequest, setSelectedRequest] = useState(null);

  const fetchRequests = async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/admin/indexing-requests`);
      if (resp.ok) {
        const data = await resp.json();
        setRequests(data.requests || []);
      }
    } catch (err) {
      console.error(err);
      showToast('Failed to fetch requests', 'error');
    } finally {
      setLoading(false);
    }
  };

  const handleRefresh = () => {
    setLoading(true);
    fetchRequests();
  };

  useEffect(() => {
    let isMounted = true;
    
    const loadData = async () => {
      try {
        const resp = await fetch(`${API_BASE_URL}/admin/indexing-requests`);
        if (resp.ok && isMounted) {
          const data = await resp.json();
          setRequests(data.requests || []);
        }
      } catch (err) {
        console.error(err);
        if (isMounted) showToast('Failed to fetch requests', 'error');
      } finally {
        if (isMounted) setLoading(false);
      }
    };
    
    loadData();
    
    return () => {
      isMounted = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const handleApprove = async (requestId, networkPath) => {
    try {
      const resp = await fetch(`${API_BASE_URL}/admin/indexing-requests/${requestId}/approve`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ networkPath })
      });

      if (resp.ok) {
        showToast('Request approved and indexing started!', 'success');
        setSelectedRequest(null);
        fetchRequests(); // Refresh list
      } else {
        const errData = await resp.json();
        showToast(`Failed to approve: ${errData.error}`, 'error');
      }
    } catch (err) {
      console.error(err);
      showToast('Failed to approve request', 'error');
    }
  };

  return (
    <div className="folders-card">
      <div className="card-header">
        <h2 className="card-title">Pending Indexing Requests</h2>
        <div className="header-actions">
          <button className="icon-btn" onClick={handleRefresh} title="Refresh requests" disabled={loading}>
            <RefreshCw size={16} className={loading ? 'spin' : ''} />
          </button>
        </div>
      </div>
      
      <div className="table-container">
        {loading ? (
          <div style={{ padding: 20, textAlign: 'center', color: '#605e5c' }}>Loading requests...</div>
        ) : requests.length === 0 ? (
          <div style={{ padding: 40, textAlign: 'center', color: '#a19f9d' }}>
            <CheckCircle size={48} style={{ marginBottom: 16, color: '#107c10' }} />
            <p style={{ margin: 0, fontWeight: 600, color: '#323130' }}>All caught up!</p>
            <p style={{ margin: '4px 0 0 0', fontSize: 13 }}>No pending requests at this time.</p>
          </div>
        ) : (
          <table className="folders-table">
            <thead>
              <tr>
                <th>Project / Job Number</th>
                <th>Requested By</th>
                <th>Date Requested</th>
                <th style={{ width: 120 }}>Action</th>
              </tr>
            </thead>
            <tbody>
              {requests.map(req => (
                <tr key={req._id}>
                  <td style={{ fontWeight: 500 }}>{req.projectName}</td>
                  <td>{req.userEmail || 'Unknown'}</td>
                  <td>{new Date(req.createdAt).toLocaleString()}</td>
                  <td>
                    <button 
                      className="btn btn-primary"
                      style={{ padding: '4px 12px', fontSize: 13 }}
                      onClick={() => setSelectedRequest(req)}
                    >
                      Approve
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {selectedRequest && (
        <ApproveRequestModal 
          request={selectedRequest}
          onClose={() => setSelectedRequest(null)}
          onSubmit={(path) => handleApprove(selectedRequest._id, path)}
        />
      )}
    </div>
  );
}
