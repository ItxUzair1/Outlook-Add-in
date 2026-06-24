
import { 
  Folder, 
  FileText, 
  AlertTriangle, 
  CheckCircle2, 
  Activity 
} from 'lucide-react';

export default function MetricsRow({ foldersCount, stats, indexingStatus, onErrorClick }) {
  return (
    <div className="metrics-row">
      <div className="metric-card">
        <div className="metric-icon blue">
          <Folder size={24} />
        </div>
        <div className="metric-details">
          <span className="metric-value">{foldersCount}</span>
          <span className="metric-label">Target Folders</span>
        </div>
      </div>

      <div className="metric-card">
        <div className="metric-icon blue">
          <FileText size={24} />
        </div>
        <div className="metric-details">
          <span className="metric-value">{stats.totalFilesFound || 0}</span>
          <span className="metric-label">Total Emails Found</span>
        </div>
      </div>

      <div className="metric-card">
        <div className="metric-icon green">
          <CheckCircle2 size={24} />
        </div>
        <div className="metric-details">
          <span className="metric-value">{stats.filesIndexed || 0}</span>
          <span className="metric-label">Emails Indexed</span>
        </div>
      </div>

      <div 
        className="metric-card" 
        style={stats.filesFailed > 0 ? { cursor: 'pointer', border: '1px solid #fecdd3' } : {}}
        onClick={() => { if (stats.filesFailed > 0 && onErrorClick) onErrorClick(); }}
      >
        <div className="metric-icon red">
          <AlertTriangle size={24} />
        </div>
        <div className="metric-details">
          <span className="metric-value">{stats.filesFailed || 0}</span>
          <span className="metric-label">Errors {stats.filesFailed > 0 && '(Click to View)'}</span>
        </div>
      </div>

      <div className="metric-card">
        <div className="metric-icon orange">
          <Activity size={24} className={indexingStatus === 'uploading' ? 'pulsing' : ''} />
        </div>
        <div className="metric-details">
          <span className="metric-value">{stats.speed || 0} /s</span>
          <span className="metric-label">Indexing Speed</span>
        </div>
      </div>
    </div>
  );
}
