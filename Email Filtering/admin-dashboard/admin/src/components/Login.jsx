import { useState } from 'react';
import { Mail, Lock } from 'lucide-react';
import logoImg from '../assets/logo.png';

export default function Login({ onLoginSuccess }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loginError, setLoginError] = useState('');

  const handleLogin = (e) => {
    e.preventDefault();
    if (email === 'admin@koyomail.com' && password === 'admin123') {
      onLoginSuccess();
      setLoginError('');
    } else {
      setLoginError('Invalid email or password. Please use admin@koyomail.com / admin123.');
    }
  };

  return (
    <div className="login-wrapper">
      <div className="login-card glass animated-fade">
        <img src={logoImg} alt="Koyomail Logo" className="login-logo" />
        <h2 className="login-title">Koyomail Universal Search</h2>
        <p className="login-subtitle">Admin Indexer Dashboard</p>
        
        {loginError && <div className="login-error">{loginError}</div>}
        
        <form onSubmit={handleLogin}>
          <div className="form-group">
            <label htmlFor="email">Email Address</label>
            <div style={{ position: 'relative' }}>
              <Mail size={16} style={{ position: 'absolute', left: 12, top: 12, color: 'var(--text-light)' }} />
              <input 
                type="email" 
                id="email" 
                className="input-control" 
                style={{ paddingLeft: 38 }}
                value={email}
                onChange={e => setEmail(e.target.value)}
                placeholder="admin@koyomail.com"
                required 
              />
            </div>
          </div>
          
          <div className="form-group">
            <label htmlFor="password">Password</label>
            <div style={{ position: 'relative' }}>
              <Lock size={16} style={{ position: 'absolute', left: 12, top: 12, color: 'var(--text-light)' }} />
              <input 
                type="password" 
                id="password" 
                className="input-control" 
                style={{ paddingLeft: 38 }}
                value={password}
                onChange={e => setPassword(e.target.value)}
                placeholder="••••••••"
                required 
              />
            </div>
          </div>
          
          <button type="submit" className="login-btn">Sign In to Indexer</button>
        </form>
      </div>
    </div>
  );
}
