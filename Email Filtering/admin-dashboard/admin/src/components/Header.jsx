
import { LogOut } from 'lucide-react';
import logoImg from '../assets/logo.png';

export default function Header({ onLogout }) {
  return (
    <header className="dashboard-header">
      <div className="header-brand">
        <img src={logoImg} alt="Koyomail logo" className="brand-logo" />
        <h1 className="brand-name">KOYOMAIL</h1>
        <span className="brand-badge">Indexer Dashboard</span>
      </div>
      <div className="header-actions">
        <span className="user-email">Administrator (admin@koyomail.com)</span>
        <button className="logout-btn" onClick={onLogout}>
          <LogOut size={14} /> Log Out
        </button>
      </div>
    </header>
  );
}
