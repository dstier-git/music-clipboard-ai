import React, { useRef, useEffect, useState, useCallback } from 'react';
import type { ChatMessage, AttachedFile, Artifact, LogLine, SSEEvent } from '../types';
import { postExtract, postMidiExport, postOpenAIEdit, subscribeJob } from '../api';

function makeId(): string {
  return Math.random().toString(36).slice(2, 10);
}

function logLevel(text: string): LogLine['level'] {
  const t = text.toLowerCase();
  if (t.includes('error') || t.includes('fail')) return 'error';
  if (t.includes('done') || t.includes('success') || t.includes('complete')) return 'success';
  if (t.includes('info') || t.startsWith('[')) return 'info';
  return 'plain';
}

function artifactIcon(type: Artifact['type']): string {
  if (type === 'midi') return '🎵';
  if (type === 'text') return '📄';
  return '📁';
}

// ── Sub-components ────────────────────────────────────────────────────────────
function MessageBubble({ msg }: { msg: ChatMessage }) {
  return (
    <div className={`message ${msg.role}`}>
      {msg.role !== 'system' && (
        <div className="message-header">
          <span className="message-role">{msg.role === 'assistant' ? 'System' : 'You'}</span>
          {msg.loading && <span className="spinner" />}
        </div>
      )}
      {msg.text && <div className="message-body">{msg.text}</div>}
      {msg.logs && msg.logs.length > 0 && (
        <div className="log-block">
          {msg.logs.map((l, i) => (
            <div key={i} className={`log-line ${l.level}`}>{l.text}</div>
          ))}
        </div>
      )}
      {msg.artifacts && msg.artifacts.map((a, i) => (
        <div key={i} className="artifact">
          <span className="artifact-icon">{artifactIcon(a.type)}</span>
          <div className="artifact-info">
            <div className="artifact-name">{a.name}</div>
            <div className="artifact-path">{a.path}</div>
          </div>
        </div>
      ))}
    </div>
  );
}

// ── Main chat view ────────────────────────────────────────────────────────────
interface Props {
  backendOnline: boolean;
}

export default function ChatView({ backendOnline }: Props) {
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: 'welcome',
      role: 'system',
      text: 'Attach a .mscx or .mscz file, then choose an action below.',
    },
  ]);
  const [attachedFiles, setAttachedFiles] = useState<AttachedFile[]>([]);
  const [inputText, setInputText] = useState('');
  const [busy, setBusy] = useState(false);
  const messagesEndRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  const addMessage = useCallback((msg: Omit<ChatMessage, 'id'>) => {
    const id = makeId();
    setMessages(prev => [...prev, { ...msg, id }]);
    return id;
  }, []);

  const updateMessage = useCallback((id: string, patch: Partial<ChatMessage>) => {
    setMessages(prev => prev.map(m => m.id === id ? { ...m, ...patch } : m));
  }, []);

  const appendLog = useCallback((id: string, line: LogLine) => {
    setMessages(prev => prev.map(m =>
      m.id === id ? { ...m, logs: [...(m.logs ?? []), line] } : m
    ));
  }, []);

  // ── Job runner ──────────────────────────────────────────────────────────────
  function runJob(
    msgId: string,
    starter: () => Promise<{ job_id: string }>
  ): void {
    setBusy(true);
    starter()
      .then(({ job_id }) => {
        updateMessage(msgId, { jobId: job_id });
        const unsub = subscribeJob(
          job_id,
          (ev: SSEEvent) => {
            if (ev.type === 'log' && ev.message) {
              appendLog(msgId, { text: ev.message, level: logLevel(ev.message) });
            } else if (ev.type === 'done') {
              const result = ev.result as Record<string, unknown> | undefined;
              const artifacts: Artifact[] = [];
              if (result?.output_path) {
                const p = result.output_path as string;
                const name = p.split('/').pop() ?? p;
                const type = name.endsWith('.mid') || name.endsWith('.midi') ? 'midi' : 'text';
                artifacts.push({ name, path: p, type });
              }
              updateMessage(msgId, {
                loading: false,
                text: 'Done.',
                artifacts: artifacts.length ? artifacts : undefined,
              });
            } else if (ev.type === 'error') {
              updateMessage(msgId, {
                loading: false,
                text: `Error: ${ev.error ?? 'Unknown error'}`,
              });
            }
          },
          () => { setBusy(false); },
          (err) => {
            updateMessage(msgId, { loading: false, text: `Connection error: ${err.message}` });
            setBusy(false);
          }
        );
        // Cleanup handled by SSE close events above
        void unsub;
      })
      .catch((err: Error) => {
        updateMessage(msgId, { loading: false, text: `Failed to start job: ${err.message}` });
        setBusy(false);
      });
  }

  // ── Actions ─────────────────────────────────────────────────────────────────
  function doExtract() {
    const f = attachedFiles[0];
    if (!f) return;
    addMessage({ role: 'user', text: `Extract pitches from ${f.name}` });
    const id = addMessage({ role: 'assistant', text: 'Starting extraction...', logs: [], loading: true });
    runJob(id, () => postExtract(f.path));
  }

  function doMidiExport() {
    const f = attachedFiles[0];
    if (!f) return;
    addMessage({ role: 'user', text: `Export MIDI from ${f.name}` });
    const id = addMessage({ role: 'assistant', text: 'Starting MIDI export...', logs: [], loading: true });
    runJob(id, () => postMidiExport(f.path));
  }

  function doOpenAIEdit() {
    const f = attachedFiles[0];
    if (!f) return;
    const instructions = inputText.trim() || 'Transpose up by a major third';
    addMessage({ role: 'user', text: `OpenAI edit MIDI: "${instructions}"` });
    const id = addMessage({ role: 'assistant', text: 'Sending to OpenAI...', logs: [], loading: true });
    setInputText('');
    runJob(id, () => postOpenAIEdit(f.path, instructions));
  }

  function handleSend() {
    if (!inputText.trim()) return;
    addMessage({ role: 'user', text: inputText.trim() });
    setInputText('');
  }

  function handleKeyDown(e: React.KeyboardEvent<HTMLTextAreaElement>) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  }

  function handleFileChange(e: React.ChangeEvent<HTMLInputElement>) {
    const files = Array.from(e.target.files ?? []);
    const newFiles: AttachedFile[] = files.map(f => ({
      name: f.name,
      path: (f as File & { path?: string }).path ?? f.name,
      size: f.size,
    }));
    setAttachedFiles(prev => [...prev, ...newFiles]);
    if (fileInputRef.current) fileInputRef.current.value = '';
  }

  function removeFile(name: string) {
    setAttachedFiles(prev => prev.filter(f => f.name !== name));
  }

  const hasScore = attachedFiles.some(f =>
    f.name.endsWith('.mscx') || f.name.endsWith('.mscz')
  );

  const hasMidi = attachedFiles.some(f =>
    f.name.endsWith('.mid') || f.name.endsWith('.midi')
  );

  return (
    <div className="chat-container">
      <div className="chat-messages">
        {messages.map(m => <MessageBubble key={m.id} msg={m} />)}
        <div ref={messagesEndRef} />
      </div>

      <div className="input-area">
        {attachedFiles.length > 0 && (
          <div className="attached-files">
            {attachedFiles.map(f => (
              <div key={f.name} className="attached-file">
                <span>{f.name}</span>
                <button onClick={() => removeFile(f.name)} title="Remove">×</button>
              </div>
            ))}
          </div>
        )}

        {/* Quick-action chips */}
        <div className="action-chips">
          <button
            className="chip"
            onClick={doExtract}
            disabled={busy || !backendOnline || !hasScore}
            title={!hasScore ? 'Attach a .mscx or .mscz file first' : ''}
          >
            Extract pitches
          </button>
          <button
            className="chip"
            onClick={doMidiExport}
            disabled={busy || !backendOnline || !hasScore}
            title={!hasScore ? 'Attach a .mscx or .mscz file first' : ''}
          >
            Export MIDI
          </button>
          <button
            className="chip"
            onClick={doOpenAIEdit}
            disabled={busy || !backendOnline || (!hasScore && !hasMidi)}
            title={(!hasScore && !hasMidi) ? 'Attach a score or MIDI file first' : ''}
          >
            OpenAI edit MIDI
          </button>
        </div>

        <div className="input-row">
          <textarea
            ref={textareaRef}
            className="chat-input"
            rows={1}
            placeholder={
              !backendOnline
                ? 'Backend offline — check Settings'
                : attachedFiles.length === 0
                ? 'Attach a file or type instructions...'
                : 'Type OpenAI edit instructions, or use action buttons above...'
            }
            value={inputText}
            onChange={e => setInputText(e.target.value)}
            onKeyDown={handleKeyDown}
            disabled={busy}
          />
          <div className="input-actions">
            <button
              className="icon-btn"
              onClick={() => fileInputRef.current?.click()}
              title="Attach file"
            >
              📎
            </button>
            <button
              className="icon-btn send-btn"
              onClick={handleSend}
              disabled={busy || !inputText.trim()}
              title="Send"
            >
              ↑
            </button>
          </div>
        </div>

        <input
          ref={fileInputRef}
          type="file"
          accept=".mscx,.mscz,.mid,.midi"
          multiple
          style={{ display: 'none' }}
          onChange={handleFileChange}
        />
      </div>
    </div>
  );
}
