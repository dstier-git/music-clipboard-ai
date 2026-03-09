import { contextBridge, ipcRenderer } from 'electron';

// ── Types shared with renderer ───────────────────────────────────────────────
export interface BackendConfig {
  port: number;
  token: string;
}

export interface BackendStatus {
  running: boolean;
  port: number;
}

// ── Exposed API surface ───────────────────────────────────────────────────────
contextBridge.exposeInMainWorld('electronAPI', {
  backend: {
    status: (): Promise<BackendStatus> => ipcRenderer.invoke('backend:status'),
    getConfig: (): Promise<BackendConfig> => ipcRenderer.invoke('backend:get-config'),
  },
  app: {
    getVersion: (): string => process.versions.electron ?? 'unknown',
    getPlatform: (): string => process.platform,
  },
});
