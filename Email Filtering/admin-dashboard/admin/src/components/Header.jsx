
import { LogOut } from 'lucide-react';
import logoImg from '../assets/logo.png';

export default function Header({ onLogout, version }) {
  return (
    <header className="dashboard-header">
      <div className="header-brand">
        <img src={logoImg} alt="Koyomail logo" className="brand-logo" />
        <h1 className="brand-name">KOYOMAIL</h1>
        <span className="brand-badge">Indexer Dashboard</span>
        {version && <span className="brand-badge version-badge">v{version}</span>}
      </div>
      <div className="header-actions">
        <span className="user-email">Administrator</span>
        <button className="logout-btn" onClick={onLogout}>
          <LogOut size={14} /> Log Out
        </button>
      </div>
    </header>
  );
}
