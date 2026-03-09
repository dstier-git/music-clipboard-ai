import React, { useEffect, useState } from 'react';
import { checkHealth } from '../api';

interface Props {
  backendOnline: boolean;
  backendPort: number;
}

export default function Settings({ backendOnline, backendPort }: Props) {
  const [openaiKeyStatus, setOpenaiKeyStatus] = useState<'present' | 'missing' | 'checking'>('checking');
  const [electronVer, setElectronVer] = useState('');
  const [platform, setPlatform] = useState('');

  useEffect(() => {
    setElectronVer(window.electronAPI.app.getVersion());
    setPlatform(window.electronAPI.app.getPlatform());

    // Check for OPENAI_API_KEY presence via health endpoint extension
    checkHealth()
      .then(ok => {
        if (!ok) {
          setOpenaiKeyStatus('missing');
          return;
        }
        // If backend is up, try a lightweight check (the health endpoint
        // returns openai_key_set in its response body)
        return fetch(`http://127.0.0.1:${backendPort}/api/v1/health`)
          .then(r => r.json())
          .then((body: { openai_key_set?: boolean }) => {
            setOpenaiKeyStatus(body.openai_key_set ? 'present' : 'missing');
          });
      })
      .catch(() => setOpenaiKeyStatus('missing'));
  }, [backendPort]);

  return (
    <div className="settings-panel">
      <div className="settings-section">
        <h2>Backend</h2>
        <div className="settings-row">
          <span className="settings-label">Status</span>
          <span className={`badge ${backendOnline ? 'ok' : 'err'}`}>
            {backendOnline ? 'Online' : 'Offline'}
          </span>
        </div>
        <div className="settings-row">
          <span className="settings-label">Port</span>
          <span className="settings-value">{backendPort || '—'}</span>
        </div>
      </div>

      <div className="settings-section">
        <h2>OpenAI</h2>
        <div className="settings-row">
          <span className="settings-label">API key</span>
          {openaiKeyStatus === 'checking' ? (
            <span className="spinner" />
          ) : (
            <span className={`badge ${openaiKeyStatus === 'present' ? 'ok' : 'warn'}`}>
              {openaiKeyStatus === 'present' ? 'Detected via env' : 'Not set (OPENAI_API_KEY)'}
            </span>
          )}
        </div>
        {openaiKeyStatus === 'missing' && (
          <div className="settings-row">
            <span className="settings-label" style={{ color: 'var(--text-muted)', fontSize: 12 }}>
              Set <code>OPENAI_API_KEY</code> in your shell environment and restart the app.
            </span>
          </div>
        )}
      </div>

      <div className="settings-section">
        <h2>Application</h2>
        <div className="settings-row">
          <span className="settings-label">Electron</span>
          <span className="settings-value">{electronVer || '—'}</span>
        </div>
        <div className="settings-row">
          <span className="settings-label">Platform</span>
          <span className="settings-value">{platform || '—'}</span>
        </div>
      </div>
    </div>
  );
}
