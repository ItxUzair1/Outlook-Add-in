
import { LogOut, Server, BarChart2, ClipboardList } from 'lucide-react';
import logoImg from '../assets/logo.png';

export default function Header({ onLogout, version, activeTab, onTabChange }) {
  return (
    <header className="dashboard-header">
      <div className="header-brand">
        <img src={logoImg} alt="Koyomail logo" className="brand-logo" />
        <h1 className="brand-name">KOYOMAIL</h1>
        {version && <span className="brand-badge version-badge">v{version}</span>}
      </div>

      {/* Nav tabs */}
      <nav className="header-nav">
        <button
          className={`header-nav-tab${activeTab === 'indexer' ? ' active' : ''}`}
          onClick={() => onTabChange?.('indexer')}
        >
          <Server size={15} />
          Indexer
        </button>
        <button
          className={`header-nav-tab${activeTab === 'requests' ? ' active' : ''}`}
          onClick={() => onTabChange?.('requests')}
        >
          <ClipboardList size={15} />
          Requests
        </button>
        <button
          className={`header-nav-tab${activeTab === 'analytics' ? ' active' : ''}`}
          onClick={() => onTabChange?.('analytics')}
        >
          <BarChart2 size={15} />
          Analytics
        </button>
      </nav>

      <div className="header-actions">
        <span className="user-email">Administrator</span>
        <button className="logout-btn" onClick={onLogout}>
          <LogOut size={14} /> Log Out
        </button>
      </div>
    </header>
  );
}
