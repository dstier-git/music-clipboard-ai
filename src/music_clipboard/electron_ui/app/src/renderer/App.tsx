import React, { useEffect, useState } from 'react';
import { initAPI, checkHealth } from './api';
import ChatView from './components/ChatView';
import Settings from './components/Settings';

type Tab = 'chat' | 'settings';

export default function App() {
  const [tab, setTab] = useState<Tab>('chat');
  const [backendOnline, setBackendOnline] = useState(false);
  const [backendPort, setBackendPort] = useState(0);
  const [ready, setReady] = useState(false);

  useEffect(() => {
    let cancelled = false;

    async function boot() {
      try {
        await initAPI();
        const cfg = await window.electronAPI.backend.getConfig();
        if (!cancelled) setBackendPort(cfg.port);
      } catch {
        // If config fetch fails, keep defaults
      }
      setReady(true);
    }

    boot();
    return () => { cancelled = true; };
  }, []);

  // Poll backend health every 5 s
  useEffect(() => {
    if (!ready) return;

    let timer: ReturnType<typeof setInterval>;

    async function poll() {
      const ok = await checkHealth().catch(() => false);
      setBackendOnline(ok);
    }

    poll();
    timer = setInterval(poll, 5_000);
    return () => clearInterval(timer);
  }, [ready]);

  if (!ready) {
    return (
      <div className="app-shell" style={{ alignItems: 'center', justifyContent: 'center' }}>
        <span className="spinner" />
      </div>
    );
  }

  return (
    <div className="app-shell">
      <div className="titlebar">
        <span className="title">Music Clipboard AI</span>
      </div>

      <div className="main-area">
        {/* Sidebar */}
        <nav className="sidebar">
          <button
            className={`sidebar-btn ${tab === 'chat' ? 'active' : ''}`}
            onClick={() => setTab('chat')}
          >
            <span className="icon">💬</span>
            Chat
          </button>
          <button
            className={`sidebar-btn ${tab === 'settings' ? 'active' : ''}`}
            onClick={() => setTab('settings')}
          >
            <span className="icon">⚙️</span>
            Settings
            <span className={`status-dot ${backendOnline ? 'online' : ''}`} />
          </button>

          <div className="sidebar-spacer" />
        </nav>

        {/* Content */}
        <div className="content">
          {tab === 'chat' && <ChatView backendOnline={backendOnline} />}
          {tab === 'settings' && <Settings backendOnline={backendOnline} backendPort={backendPort} />}
        </div>
      </div>
    </div>
  );
}
