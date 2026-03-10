import { useEffect, useMemo, useRef, useState } from 'react';

type ActionType = 'extract' | 'midi-export' | 'openai-edit';
type TabType = 'chat' | 'settings';

type ChatMessage = {
  id: string;
  role: 'system' | 'user' | 'assistant';
  text: string;
};

type BackendStatus = {
  running: boolean;
  port: number | null;
  baseUrl: string | null;
};

type HealthPayload = {
  status: string;
  openai_key_set: boolean;
};

type JobEvent = {
  type: string;
  message?: string;
  result?: unknown;
  error?: string;
  progress?: number;
};

const ACTION_LABELS: Record<ActionType, string> = {
  extract: 'Extract Pitch Text',
  'midi-export': 'Export MIDI',
  'openai-edit': 'OpenAI MIDI Edit',
};

function createMessage(role: ChatMessage['role'], text: string): ChatMessage {
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    role,
    text,
  };
}

function formatResult(result: unknown): string {
  if (typeof result === 'string') {
    return result;
  }
  try {
    return JSON.stringify(result, null, 2);
  } catch {
    return String(result);
  }
}

export function App() {
  const [activeTab, setActiveTab] = useState<TabType>('chat');
  const [action, setAction] = useState<ActionType>('extract');
  const [filePath, setFilePath] = useState('');
  const [instructions, setInstructions] = useState('');
  const [messages, setMessages] = useState<ChatMessage[]>([
    createMessage('assistant', 'Music Clipboard AI ready. Choose an action, pick a file, and run.'),
  ]);
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [backendStatus, setBackendStatus] = useState<BackendStatus>({
    running: false,
    port: null,
    baseUrl: null,
  });
  const [health, setHealth] = useState<HealthPayload | null>(null);
  const [platform, setPlatform] = useState('unknown');

  const subscriptionsRef = useRef<Array<() => void>>([]);

  const selectedActionLabel = useMemo(() => ACTION_LABELS[action], [action]);

  const addMessage = (role: ChatMessage['role'], text: string) => {
    setMessages((prev) => [...prev, createMessage(role, text)]);
  };

  const refreshStatus = async () => {
    const status = await window.musicClipboard.backend.status();
    setBackendStatus(status);
  };

  const refreshHealth = async () => {
    try {
      const response = await window.musicClipboard.api.request('GET', '/api/v1/health');
      setHealth(response.data as HealthPayload);
    } catch (error) {
      addMessage('system', `Health check failed: ${String(error)}`);
    }
  };

  useEffect(() => {
    void (async () => {
      try {
        const status = await window.musicClipboard.backend.start();
        setBackendStatus(status);
        setPlatform(await window.musicClipboard.app.getPlatform());
        await refreshHealth();
      } catch (error) {
        addMessage('system', `Backend startup failed: ${String(error)}`);
      }
    })();

    return () => {
      for (const unsubscribe of subscriptionsRef.current) {
        unsubscribe();
      }
      subscriptionsRef.current = [];
    };
  }, []);

  const pickFile = async () => {
    const selected = await window.musicClipboard.files.pick();
    if (selected) {
      setFilePath(selected);
    }
  };

  const runAction = async () => {
    if (!filePath.trim()) {
      addMessage('system', 'Pick a source file first.');
      return;
    }
    if (action === 'openai-edit' && !instructions.trim()) {
      addMessage('system', 'OpenAI edit requires instructions.');
      return;
    }

    setIsSubmitting(true);

    try {
      let endpoint = '/api/v1/extract';
      let payload: Record<string, unknown> = { file_path: filePath.trim() };

      if (action === 'midi-export') {
        endpoint = '/api/v1/midi/export';
      }
      if (action === 'openai-edit') {
        endpoint = '/api/v1/ai/openai/edit-midi';
        payload = {
          file_path: filePath.trim(),
          instructions: instructions.trim(),
        };
      }

      addMessage('user', `${selectedActionLabel}\n${filePath.trim()}${action === 'openai-edit' ? `\n${instructions.trim()}` : ''}`);

      const response = await window.musicClipboard.api.request('POST', endpoint, payload);
      const data = response.data as { job_id: string };
      const jobId = data.job_id;
      addMessage('assistant', `Job queued: ${jobId}`);

      const unsubscribe = await window.musicClipboard.jobs.subscribe(jobId, (event: JobEvent) => {
        if (event.type === 'log' && event.message) {
          addMessage('assistant', event.message);
          return;
        }
        if (event.type === 'status' && event.message) {
          addMessage('assistant', `Status: ${event.message}`);
          return;
        }
        if (event.type === 'progress' && typeof event.progress === 'number') {
          addMessage('assistant', `Progress: ${Math.round(event.progress * 100)}%`);
          return;
        }
        if (event.type === 'done') {
          addMessage('assistant', `Completed:\n${formatResult(event.result)}`);
          return;
        }
        if (event.type === 'error') {
          addMessage('assistant', `Failed: ${event.error || 'Unknown error'}`);
        }
      });

      subscriptionsRef.current.push(unsubscribe);
      await refreshStatus();
      await refreshHealth();
    } catch (error) {
      addMessage('assistant', `Request failed: ${String(error)}`);
    } finally {
      setIsSubmitting(false);
    }
  };

  const stopBackend = async () => {
    await window.musicClipboard.backend.stop();
    await refreshStatus();
  };

  const startBackend = async () => {
    await window.musicClipboard.backend.start();
    await refreshStatus();
    await refreshHealth();
  };

  return (
    <div className="app-shell">
      <header className="top-bar">
        <div className="brand-block">
          <h1>Music Clipboard AI</h1>
          <p>Desktop chat interface for local extraction, MIDI export, and OpenAI editing.</p>
        </div>
        <div className="status-chip" data-running={backendStatus.running}>
          {backendStatus.running ? 'Backend Online' : 'Backend Offline'}
        </div>
      </header>

      <nav className="tab-row">
        <button className={activeTab === 'chat' ? 'active' : ''} onClick={() => setActiveTab('chat')}>
          Chat
        </button>
        <button className={activeTab === 'settings' ? 'active' : ''} onClick={() => setActiveTab('settings')}>
          Settings
        </button>
      </nav>

      {activeTab === 'chat' && (
        <main className="chat-layout">
          <section className="composer-panel">
            <h2>Run Action</h2>

            <label>
              Action
              <select value={action} onChange={(event) => setAction(event.target.value as ActionType)}>
                <option value="extract">Extract Pitch Text</option>
                <option value="midi-export">Export MIDI</option>
                <option value="openai-edit">OpenAI MIDI Edit</option>
              </select>
            </label>

            <label>
              Source File
              <div className="file-row">
                <input value={filePath} onChange={(event) => setFilePath(event.target.value)} placeholder="Select .mscx/.mscz/.mid file" />
                <button type="button" onClick={pickFile}>Browse</button>
              </div>
            </label>

            <label>
              Instructions (OpenAI edit only)
              <textarea
                value={instructions}
                onChange={(event) => setInstructions(event.target.value)}
                placeholder="Example: Transpose all notes up a perfect fifth and humanize velocity"
                rows={5}
              />
            </label>

            <button className="primary" type="button" onClick={runAction} disabled={isSubmitting}>
              {isSubmitting ? 'Running...' : `Run ${selectedActionLabel}`}
            </button>
          </section>

          <section className="chat-panel">
            <h2>Session</h2>
            <div className="message-list">
              {messages.map((message) => (
                <article key={message.id} className={`message message-${message.role}`}>
                  <h3>{message.role}</h3>
                  <pre>{message.text}</pre>
                </article>
              ))}
            </div>
          </section>
        </main>
      )}

      {activeTab === 'settings' && (
        <main className="settings-layout">
          <section className="settings-card">
            <h2>Backend</h2>
            <p>Platform: <strong>{platform}</strong></p>
            <p>Running: <strong>{backendStatus.running ? 'Yes' : 'No'}</strong></p>
            <p>Base URL: <code>{backendStatus.baseUrl || 'n/a'}</code></p>
            <div className="row-actions">
              <button type="button" onClick={startBackend}>Start</button>
              <button type="button" onClick={stopBackend}>Stop</button>
              <button type="button" onClick={refreshStatus}>Refresh</button>
            </div>
          </section>

          <section className="settings-card">
            <h2>OpenAI Diagnostics</h2>
            <p>
              API key detected: <strong>{health?.openai_key_set ? 'Yes' : 'No'}</strong>
            </p>
            <p>
              Health status: <strong>{health?.status || 'Unknown'}</strong>
            </p>
            <div className="row-actions">
              <button type="button" onClick={refreshHealth}>Run Health Check</button>
            </div>
          </section>
        </main>
      )}
    </div>
  );
}
