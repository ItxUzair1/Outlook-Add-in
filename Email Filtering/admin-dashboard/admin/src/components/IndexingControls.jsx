
import { Activity, Play, Pause, RotateCcw } from 'lucide-react';

export default function IndexingControls({
  indexingStatus,
  schedulerStatus,
  stats,
  foldersCount,
  onStart,
  onPause,
  onReset,
  onFastSync,
  onStartScheduler,
  onStopScheduler
}) {
  const calculateProgress = () => {
    if (!stats.totalFilesFound) return 0;
    const progress = (stats.filesIndexed / stats.totalFilesFound) * 100;
    return Math.min(Math.round(progress), 100);
  };

  return (
    <div className="status-panel">
      <h3 className="panel-title">
        <Activity size={18} style={{ color: 'var(--primary-color)' }} /> Indexing Controls
      </h3>
      
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
        <span style={{ fontSize: '14px', fontWeight: '500' }}>Indexer Status:</span>
        <span className={`status-badge ${indexingStatus}`}>
          {indexingStatus === 'uploading' ? 'Uploading' : indexingStatus === 'scanning' ? 'Scanning Disk' : indexingStatus}
        </span>
      </div>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', paddingBottom: '16px', borderBottom: '1px solid var(--border-color)' }}>
        <span style={{ fontSize: '14px', fontWeight: '500' }}>Live Scheduler (15m):</span>
        <span className={`status-badge ${schedulerStatus === 'active' ? 'completed' : 'idle'}`}>
          {schedulerStatus === 'active' ? 'Active' : 'Inactive'}
        </span>
      </div>

      {/* Progress Tracker */}
      <div className="progress-container">
        <div className="progress-header">
          <span>Indexing Progress</span>
          <span style={{ fontWeight: '600' }}>{calculateProgress()}%</span>
        </div>
        <div className="progress-bar-bg">
          <div className="progress-bar-fill" style={{ width: `${calculateProgress()}%` }}></div>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '6px', fontSize: '11px', color: 'var(--text-light)' }}>
          <span>Processed: {stats.filesIndexed || 0}</span>
          <span>Total: {stats.totalFilesFound || 0}</span>
        </div>
      </div>

      {/* Status Indicator text */}
      {indexingStatus === 'uploading' && stats.currentFilePath && (
        <div style={{ fontSize: '11px', color: 'var(--text-light)', wordBreak: 'break-all', marginBottom: '16px', backgroundColor: '#f3f2f1', padding: '8px', borderRadius: '4px' }}>
          <strong>Indexing:</strong> {stats.currentFilePath}
        </div>
      )}

      {/* Control Buttons */}
      <div className="controls-row">
        {indexingStatus === 'idle' || indexingStatus === 'paused' || indexingStatus === 'completed' ? (
          <button 
            className="control-btn start" 
            onClick={onStart}
            disabled={foldersCount === 0}
          >
            <Play size={16} /> {indexingStatus === 'paused' ? 'Resume' : 'Start Indexing'}
          </button>
        ) : (
          <button className="control-btn pause" onClick={onPause}>
            <Pause size={16} /> Pause Job
          </button>
        )}

        <button 
          className="control-btn reset" 
          onClick={onReset}
          title="Reset statistics & logs"
        >
          <RotateCcw size={16} /> Reset Progress
        </button>

        <button 
          className="control-btn" 
          style={{ backgroundColor: '#0078d4', color: '#fff' }}
          onClick={onFastSync}
          title="Instantly sync folder permissions to Meilisearch without re-parsing files"
        >
          <Activity size={16} /> Fast Sync
        </button>
      </div>

      <div className="controls-row" style={{ marginTop: '12px' }}>
        {schedulerStatus === 'inactive' ? (
          <button 
            className="control-btn start" 
            style={{ backgroundColor: '#107c41' }}
            onClick={onStartScheduler}
            disabled={foldersCount === 0}
          >
            <Play size={16} /> Start Live Scheduler
          </button>
        ) : (
          <button 
            className="control-btn pause" 
            onClick={onStopScheduler}
          >
            <Pause size={16} /> Stop Scheduler
          </button>
        )}
      </div>
    </div>
  );
}

