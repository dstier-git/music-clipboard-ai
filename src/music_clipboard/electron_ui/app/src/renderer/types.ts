// ── Electron bridge ───────────────────────────────────────────────────────────
export interface ElectronAPI {
  backend: {
    status(): Promise<{ running: boolean; port: number }>;
    getConfig(): Promise<{ port: number; token: string }>;
  };
  app: {
    getVersion(): string;
    getPlatform(): string;
  };
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}

// ── API / Job types ───────────────────────────────────────────────────────────
export type JobStatus = 'queued' | 'running' | 'done' | 'error';

export interface Job {
  job_id: string;
  status: JobStatus;
  created_at: string;
  updated_at: string;
  result?: unknown;
  error?: string;
}

export interface SSEEvent {
  type: 'log' | 'progress' | 'done' | 'error';
  message?: string;
  result?: unknown;
  error?: string;
}

// ── Message types for chat UI ─────────────────────────────────────────────────
export type MessageRole = 'user' | 'assistant' | 'system';

export interface LogLine {
  text: string;
  level: 'info' | 'error' | 'success' | 'plain';
}

export interface Artifact {
  name: string;
  path: string;
  type: 'text' | 'midi' | 'file';
}

export interface ChatMessage {
  id: string;
  role: MessageRole;
  text?: string;
  logs?: LogLine[];
  artifacts?: Artifact[];
  jobId?: string;
  loading?: boolean;
}

// ── Attached file ─────────────────────────────────────────────────────────────
export interface AttachedFile {
  name: string;
  path: string;
  size: number;
}
