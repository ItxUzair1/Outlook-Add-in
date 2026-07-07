import { Activity, Play, Pause, RotateCcw, Wrench } from 'lucide-react';

export default function IndexingControls({
  indexingStatus,
  schedulerStatus,
  stats,
  foldersCount,
  onStart,
  onPause,
  onReset,
  onFastSync,
  onRepairMetadata,
  onRetryErrors,
  onStartScheduler,
  onStopScheduler
}) {
  const calculateProgress = () => {
    if ((indexingStatus === 'repairing' || indexingStatus === 'retrying') && stats.totalFilesFound) {
      const progress = ((stats.filesIndexedThisSession || 0) / stats.totalFilesFound) * 100;
      return Math.min(Math.round(progress), 100);
    }
    if (!stats.totalFilesFound) return 0;
    const processed = (stats.filesSkipped || 0) + (stats.filesIndexedThisSession || 0);
    const progress = (processed / stats.totalFilesFound) * 100;
    return Math.min(Math.round(progress), 100);
  };

  const isBusy = indexingStatus === 'uploading' || indexingStatus === 'scanning' || indexingStatus === 'repairing' || indexingStatus === 'retrying';
  const statusLabel = indexingStatus === 'uploading'
    ? 'Uploading'
    : indexingStatus === 'scanning'
      ? 'Scanning Disk'
      : indexingStatus === 'repairing'
        ? 'Repairing Metadata'
        : indexingStatus === 'retrying'
          ? 'Retrying Errors'
          : indexingStatus;

  return (
    <div className="status-panel">
      <h3 className="panel-title">
        <Activity size={18} style={{ color: 'var(--primary-color)' }} /> Indexing Controls
      </h3>
      
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px' }}>
        <span style={{ fontSize: '14px', fontWeight: '500' }}>Indexer Status:</span>
        <span className={`status-badge ${indexingStatus}`}>
          {statusLabel}
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
          <span>{indexingStatus === 'repairing' ? 'Repair Progress' : indexingStatus === 'retrying' ? 'Recovery Progress' : 'Indexing Progress'}</span>
          <span style={{ fontWeight: '600' }}>{calculateProgress()}%</span>
        </div>
        <div className="progress-bar-bg">
          <div className="progress-bar-fill" style={{ width: `${calculateProgress()}%` }}></div>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: '6px', fontSize: '11px', color: 'var(--text-light)' }}>
          {indexingStatus === 'repairing' || indexingStatus === 'retrying' ? (
            <>
              <span>Processed: {stats.filesIndexedThisSession || 0}</span>
              <span>Total: {stats.totalFilesFound || 0}</span>
            </>
          ) : (
            <>
              <span>Indexed: {stats.filesIndexed || 0}</span>
              <span>Scanned: {stats.totalFilesFound || 0}</span>
            </>
          )}
        </div>
      </div>

      {/* Status Indicator text */}
      {isBusy && stats.currentFilePath && (
        <div style={{ fontSize: '11px', color: 'var(--text-light)', wordBreak: 'break-all', marginBottom: '16px', backgroundColor: '#f3f2f1', padding: '8px', borderRadius: '4px' }}>
          <strong>
            {indexingStatus === 'scanning' ? 'Scanning:' : indexingStatus === 'repairing' ? 'Repairing:' : indexingStatus === 'retrying' ? 'Recovering:' : 'Indexing:'}
          </strong>{' '}
          {stats.currentFilePath}
        </div>
      )}

      {indexingStatus === 'repairing' && (
        <div style={{ fontSize: '11px', color: '#605e5c', marginBottom: '16px', padding: '8px 10px', borderRadius: '4px', backgroundColor: '#f3f2f1' }}>
          Repair runs in the background — the app should stay responsive. You can watch progress in the log below.
        </div>
      )}

      {indexingStatus === 'retrying' && (
        <div style={{ fontSize: '11px', color: '#605e5c', marginBottom: '16px', padding: '8px 10px', borderRadius: '4px', backgroundColor: '#f3f2f1' }}>
          Recovery runs in the background — processing error emails with safe fallback parser. Watch logs below.
        </div>
      )}

      {/* Control Buttons */}
      <div className="controls-row">
        {!isBusy || indexingStatus === 'paused' ? (
          <button 
            className="control-btn start" 
            onClick={onStart}
            disabled={foldersCount === 0 || indexingStatus === 'repairing' || indexingStatus === 'retrying'}
          >
            <Play size={16} /> {indexingStatus === 'paused' ? 'Resume' : 'Start Indexing'}
          </button>
        ) : (
          <button className="control-btn pause" onClick={onPause}>
            <Pause size={16} /> {indexingStatus === 'repairing' ? 'Stop Repair' : indexingStatus === 'retrying' ? 'Stop Recovery' : 'Pause Job'}
          </button>
        )}

        <button 
          className="control-btn reset" 
          onClick={onReset}
          title="Reset statistics & logs"
          disabled={isBusy}
        >
          <RotateCcw size={16} /> Reset Progress
        </button>

        <button 
          className="control-btn" 
          style={{ backgroundColor: '#0078d4', color: '#fff' }}
          onClick={onFastSync}
          title="Instantly sync folder permissions to Meilisearch without re-parsing files"
          disabled={isBusy}
        >
          <Activity size={16} /> Fast Sync
        </button>

        <button 
          className="control-btn" 
          style={{ backgroundColor: '#8a2be2', color: '#fff' }}
          onClick={onRepairMetadata}
          disabled={isBusy}
          title="Fix missing To, Cc, and Date on already-indexed emails (no full re-index)"
        >
          <Wrench size={16} /> Repair Metadata
        </button>

        <button 
          className="control-btn" 
          style={{ backgroundColor: '#a4262c', color: '#fff' }}
          onClick={onRetryErrors}
          disabled={isBusy || !stats.filesFailed}
          title={`Retry parsing the ${stats.filesFailed || 0} previously failed (error) emails using the robust fallback parser`}
        >
          <RotateCcw size={16} /> Retry Error Emails
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

