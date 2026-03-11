import { Fragment, useEffect, useMemo, useRef, useState, type DragEvent } from 'react';

type ActionType = 'extract' | 'midi-export' | 'openai-edit';
type TabType = 'chat' | 'settings';
type ChatAttachmentKind = 'midi' | 'text' | 'file';

type ChatAttachment = {
  filePath: string;
  fileName: string;
  kind: ChatAttachmentKind;
};

type ChatMessage = {
  id: string;
  role: 'system' | 'user' | 'assistant';
  text: string;
  attachment?: ChatAttachment;
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

type ProgramInfo = {
  program_id: string;
  label: string;
};

type HotkeyDefault = {
  normalized: string;
  label: string;
};

type GlobalHotkeyStatus = {
  global_hotkey: string;
  global_hotkey_label: string;
  selected_program?: string;
  registered: boolean;
  available: boolean;
  provider: string | null;
  error: string | null;
  request_monitor_running: boolean;
  request_file: string;
  last_trigger_source: string | null;
  last_trigger_at: string | null;
  last_trigger_result: {
    ok: boolean;
    code: string;
    message: string;
  } | null;
};

type HotkeySettingsPayload = {
  selected_program: string;
  visible_programs: string[];
  custom_hotkeys: Record<string, string>;
  default_hotkeys_by_program: Record<string, HotkeyDefault>;
  effective_hotkey: string | null;
  effective_hotkey_label: string;
  effective_hotkey_error: string | null;
  global_hotkey: string;
  program_order: string[];
  program_labels: Record<string, string>;
  programs: ProgramInfo[];
  global_hotkey_status: GlobalHotkeyStatus;
};

type TriggerSaveSelectionResult = {
  ok: boolean;
  code: string;
  message: string;
  logs: string[];
  program_id?: string;
  program_label?: string;
  effective_hotkey?: string;
  effective_hotkey_label?: string;
};

type HotkeySettingsDraft = {
  selectedProgram: string;
  visiblePrograms: string[];
  customHotkeys: Record<string, string>;
};

const ACTION_LABELS: Record<ActionType, string> = {
  extract: 'Extract Pitch Text',
  'midi-export': 'Export MIDI',
  'openai-edit': 'OpenAI MIDI Edit',
};

function createMessage(
  role: ChatMessage['role'],
  text: string,
  attachment?: ChatAttachment,
): ChatMessage {
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    role,
    text,
    attachment,
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

function getFileName(filePath: string): string {
  const normalized = filePath.replace(/\\/g, '/');
  const segments = normalized.split('/');
  return segments[segments.length - 1] || filePath;
}

function getAttachmentKind(fileName: string): ChatAttachmentKind {
  const lower = fileName.toLowerCase();
  if (lower.endsWith('.mid') || lower.endsWith('.midi')) {
    return 'midi';
  }
  if (lower.endsWith('.txt')) {
    return 'text';
  }
  return 'file';
}

function getOutputAttachment(result: unknown): ChatAttachment | undefined {
  if (!result || typeof result !== 'object') {
    return undefined;
  }

  const outputPath = (result as { output_path?: unknown }).output_path;
  if (typeof outputPath !== 'string' || !outputPath.trim()) {
    return undefined;
  }

  const filePath = outputPath.trim();
  const fileName = getFileName(filePath);
  return {
    filePath,
    fileName,
    kind: getAttachmentKind(fileName),
  };
}

function toSettingsDraft(settings: HotkeySettingsPayload): HotkeySettingsDraft {
  return {
    selectedProgram: settings.selected_program,
    visiblePrograms: [...settings.visible_programs],
    customHotkeys: { ...settings.custom_hotkeys },
  };
}

export function App() {
  const [activeTab, setActiveTab] = useState<TabType>('chat');
  const [action, setAction] = useState<ActionType>('openai-edit');
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

  const [hotkeySettings, setHotkeySettings] = useState<HotkeySettingsPayload | null>(null);
  const [hotkeyDraft, setHotkeyDraft] = useState<HotkeySettingsDraft | null>(null);
  const [isSavingHotkeys, setIsSavingHotkeys] = useState(false);
  const [isTriggeringSaveSelection, setIsTriggeringSaveSelection] = useState(false);
  const [settingsNotice, setSettingsNotice] = useState('');

  const subscriptionsRef = useRef<Array<() => void>>([]);

  const selectedActionLabel = useMemo(() => ACTION_LABELS[action], [action]);

  const addMessage = (role: ChatMessage['role'], text: string, attachment?: ChatAttachment) => {
    setMessages((prev) => [...prev, createMessage(role, text, attachment)]);
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

  const refreshHotkeySettings = async ({ syncDraft = false }: { syncDraft?: boolean } = {}) => {
    try {
      const response = await window.musicClipboard.hotkeys.getSettings();
      const payload = response.data as HotkeySettingsPayload;
      setHotkeySettings(payload);
      if (syncDraft || hotkeyDraft === null) {
        setHotkeyDraft(toSettingsDraft(payload));
      }
      return payload;
    } catch (error) {
      addMessage('system', `Failed to load hotkey settings: ${String(error)}`);
      return null;
    }
  };

  useEffect(() => {
    void (async () => {
      try {
        const status = await window.musicClipboard.backend.start();
        setBackendStatus(status);
        setPlatform(await window.musicClipboard.app.getPlatform());
        await refreshHealth();
        await refreshHotkeySettings({ syncDraft: true });
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const pickFile = async () => {
    const selected = await window.musicClipboard.files.pick();
    if (selected) {
      setFilePath(selected);
    }
  };

  const handleAttachmentDragStart = (event: DragEvent<HTMLButtonElement>, attachment: ChatAttachment) => {
    const dragFilePath = attachment.filePath.trim();
    if (!dragFilePath) {
      event.preventDefault();
      return;
    }

    const ghostImage = new Image();
    event.dataTransfer.setDragImage(ghostImage, 0, 0);
    event.dataTransfer.effectAllowed = 'copy';

    window.musicClipboard.files.startDrag(dragFilePath);
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

      const targetProgramLabel = hotkeySettings?.program_labels[hotkeySettings.selected_program] || 'Target App';
      addMessage(
        'user',
        `${selectedActionLabel}\n${filePath.trim()}\nTarget: ${targetProgramLabel}${action === 'openai-edit' ? `\n${instructions.trim()}` : ''}`,
      );

      const response = await window.musicClipboard.api.request('POST', endpoint, payload);
      const data = response.data as { job_id: string };
      const jobId = data.job_id;
      addMessage('assistant', `Job queued: ${jobId}`);

      let unsubscribe: () => void = () => undefined;

      unsubscribe = await window.musicClipboard.jobs.subscribe(jobId, (event: JobEvent) => {
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
          const attachment = getOutputAttachment(event.result);
          const completionText = attachment ? 'Completed.' : `Completed:\n${formatResult(event.result)}`;
          addMessage('assistant', completionText, attachment);
          unsubscribe();
          return;
        }
        if (event.type === 'error') {
          addMessage('assistant', `Failed: ${event.error || 'Unknown error'}`);
          unsubscribe();
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

  const triggerSaveSelection = async () => {
    if (!hotkeySettings) {
      addMessage('system', 'Hotkey settings are not loaded yet.');
      return;
    }

    setIsTriggeringSaveSelection(true);
    try {
      const response = await window.musicClipboard.hotkeys.triggerSaveSelection(hotkeySettings.selected_program);
      const result = response.data as TriggerSaveSelectionResult;
      addMessage('assistant', result.message || 'Triggered save/export.');
      if (Array.isArray(result.logs)) {
        for (const entry of result.logs) {
          addMessage('assistant', entry);
        }
      }
    } catch (error: unknown) {
      const text = String(error);
      addMessage('assistant', `Save/export trigger failed: ${text}`);
    } finally {
      await refreshHotkeySettings();
      setIsTriggeringSaveSelection(false);
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
    await refreshHotkeySettings();
  };

  const updateSelectedProgram = async (selectedProgram: string) => {
    if (!hotkeySettings) {
      return;
    }

    try {
      const response = await window.musicClipboard.hotkeys.updateSettings({ selected_program: selectedProgram });
      const payload = response.data as HotkeySettingsPayload;
      setHotkeySettings(payload);
      setHotkeyDraft((current) => ({
        selectedProgram: payload.selected_program,
        visiblePrograms: [...(current?.visiblePrograms || payload.visible_programs)],
        customHotkeys: { ...(current?.customHotkeys || payload.custom_hotkeys) },
      }));
      setSettingsNotice('Target app updated.');
    } catch (error) {
      addMessage('system', `Failed to update target app: ${String(error)}`);
    }
  };

  const toggleProgramVisibility = (programId: string) => {
    setHotkeyDraft((current) => {
      if (!current) {
        return current;
      }
      const currentlyVisible = current.visiblePrograms.includes(programId);
      const nextVisible = currentlyVisible
        ? current.visiblePrograms.filter((item) => item !== programId)
        : [...current.visiblePrograms, programId];

      return {
        ...current,
        visiblePrograms: nextVisible,
      };
    });
  };

  const updateCustomHotkeyDraft = (programId: string, value: string) => {
    setHotkeyDraft((current) => {
      if (!current) {
        return current;
      }
      return {
        ...current,
        customHotkeys: {
          ...current.customHotkeys,
          [programId]: value,
        },
      };
    });
  };

  const resetCustomHotkeyDraft = (programId: string) => {
    updateCustomHotkeyDraft(programId, '');
  };

  const saveHotkeySettings = async () => {
    if (!hotkeyDraft) {
      return;
    }

    setIsSavingHotkeys(true);
    setSettingsNotice('');

    try {
      const selectedProgram = hotkeyDraft.visiblePrograms.includes(hotkeyDraft.selectedProgram)
        ? hotkeyDraft.selectedProgram
        : (hotkeyDraft.visiblePrograms[0] || hotkeyDraft.selectedProgram);

      const response = await window.musicClipboard.hotkeys.updateSettings({
        selected_program: selectedProgram,
        visible_programs: hotkeyDraft.visiblePrograms,
        custom_hotkeys: hotkeyDraft.customHotkeys,
      });
      const payload = response.data as HotkeySettingsPayload;
      setHotkeySettings(payload);
      setHotkeyDraft(toSettingsDraft(payload));
      setSettingsNotice('Settings saved.');
      addMessage('assistant', 'Program visibility and custom hotkeys saved.');
    } catch (error) {
      setSettingsNotice(`Failed to save settings: ${String(error)}`);
      addMessage('system', `Failed to save hotkey settings: ${String(error)}`);
    } finally {
      setIsSavingHotkeys(false);
    }
  };

  const reloadGlobalHotkey = async () => {
    try {
      await window.musicClipboard.hotkeys.reloadGlobal();
      await refreshHotkeySettings();
      setSettingsNotice('Global hotkey registration reloaded.');
    } catch (error) {
      setSettingsNotice(`Failed to reload global hotkey: ${String(error)}`);
      addMessage('system', `Failed to reload global hotkey: ${String(error)}`);
    }
  };

  const visibleProgramIds = hotkeySettings?.visible_programs || hotkeySettings?.program_order || [];
  const selectedProgramLabel = hotkeySettings
    ? hotkeySettings.program_labels[hotkeySettings.selected_program] || hotkeySettings.selected_program
    : 'Target App';

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
                <option value="openai-edit">OpenAI MIDI Edit</option>
                <option value="extract">Extract Pitch Text</option>
                <option value="midi-export">Export MIDI</option>
              </select>
            </label>

            <label>
              Target App
              <select
                value={hotkeySettings?.selected_program || ''}
                onChange={(event) => void updateSelectedProgram(event.target.value)}
                disabled={!hotkeySettings}
              >
                {visibleProgramIds.map((programId) => (
                  <option key={programId} value={programId}>
                    {hotkeySettings?.program_labels[programId] || programId}
                  </option>
                ))}
              </select>
            </label>

            <button
              className="secondary"
              type="button"
              onClick={triggerSaveSelection}
              disabled={!hotkeySettings || isTriggeringSaveSelection}
            >
              {isTriggeringSaveSelection
                ? 'Triggering...'
                : `Trigger Save/Export in ${selectedProgramLabel}`}
            </button>

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
              {messages.map((message) => {
                const attachment = message.attachment;
                const canDrag = Boolean(attachment?.filePath.trim());
                return (
                  <article key={message.id} className={`message message-${message.role}`}>
                    <h3>{message.role}</h3>
                    <div className="message-body">
                      {message.text ? <pre>{message.text}</pre> : null}
                      {attachment ? (
                        <div className="message-file">
                          <button
                            type="button"
                            className={`file-chip file-chip-${attachment.kind}${canDrag ? '' : ' is-disabled'}`}
                            draggable={canDrag}
                            disabled={!canDrag}
                            onDragStart={canDrag ? (event) => handleAttachmentDragStart(event, attachment) : undefined}
                          >
                            <span className="file-chip-name">{attachment.fileName}</span>
                            <span className="file-chip-hint">Drag to DAW or app</span>
                          </button>
                          <p className="file-path">{attachment.filePath}</p>
                        </div>
                      ) : null}
                    </div>
                  </article>
                );
              })}
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

          <section className="settings-card settings-wide">
            <h2>Program Visibility & Hotkeys</h2>
            <p className="settings-help">
              Control which target apps appear in Run Action and override save/export shortcuts per app.
              Leave custom hotkey blank to use platform default.
            </p>

            {!hotkeySettings || !hotkeyDraft ? (
              <p>Loading settings...</p>
            ) : (
              <>
                <div className="hotkey-table">
                  <div className="hotkey-header">Program</div>
                  <div className="hotkey-header">Show</div>
                  <div className="hotkey-header">Custom Hotkey</div>
                  <div className="hotkey-header">Default</div>
                  <div className="hotkey-header">Action</div>

                  {hotkeySettings.program_order.map((programId) => (
                    <Fragment key={programId}>
                      <div className="hotkey-cell">{hotkeySettings.program_labels[programId] || programId}</div>
                      <div className="hotkey-cell">
                        <input
                          type="checkbox"
                          checked={hotkeyDraft.visiblePrograms.includes(programId)}
                          onChange={() => toggleProgramVisibility(programId)}
                        />
                      </div>
                      <div className="hotkey-cell">
                        <input
                          value={hotkeyDraft.customHotkeys[programId] || ''}
                          onChange={(event) => updateCustomHotkeyDraft(programId, event.target.value)}
                          placeholder="e.g. cmd+shift+s"
                        />
                      </div>
                      <div className="hotkey-cell muted">
                        {hotkeySettings.default_hotkeys_by_program[programId]?.label || 'No default'}
                      </div>
                      <div className="hotkey-cell">
                        <button type="button" onClick={() => resetCustomHotkeyDraft(programId)}>Reset</button>
                      </div>
                    </Fragment>
                  ))}
                </div>

                <div className="row-actions">
                  <button type="button" onClick={saveHotkeySettings} disabled={isSavingHotkeys}>
                    {isSavingHotkeys ? 'Saving...' : 'Save Settings'}
                  </button>
                </div>
              </>
            )}

            {settingsNotice ? <p className="settings-notice">{settingsNotice}</p> : null}
          </section>

          <section className="settings-card settings-wide">
            <h2>Global Hotkey Status</h2>
            <p>
              Global hotkey: <strong>{hotkeySettings?.global_hotkey_status.global_hotkey_label || hotkeySettings?.global_hotkey || 'Unknown'}</strong>
            </p>
            <p>
              Registered: <strong>{hotkeySettings?.global_hotkey_status.registered ? 'Yes' : 'No'}</strong>
              {' · '}
              Provider: <strong>{hotkeySettings?.global_hotkey_status.provider || 'n/a'}</strong>
              {' · '}
              Request monitor: <strong>{hotkeySettings?.global_hotkey_status.request_monitor_running ? 'Running' : 'Stopped'}</strong>
            </p>
            <p>
              Last trigger: <strong>{hotkeySettings?.global_hotkey_status.last_trigger_at || 'Never'}</strong>
              {' · '}
              Source: <strong>{hotkeySettings?.global_hotkey_status.last_trigger_source || 'n/a'}</strong>
            </p>
            {hotkeySettings?.global_hotkey_status.error ? (
              <p className="settings-error">{hotkeySettings.global_hotkey_status.error}</p>
            ) : null}
            <div className="row-actions">
              <button type="button" onClick={reloadGlobalHotkey}>Reload Global Hotkey</button>
              <button type="button" onClick={() => void refreshHotkeySettings()}>Refresh Status</button>
            </div>
          </section>
        </main>
      )}
    </div>
  );
}
