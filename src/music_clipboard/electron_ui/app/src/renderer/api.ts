import type { Job, SSEEvent } from './types';

let _port = 0;
let _token = '';

export async function initAPI(): Promise<void> {
  const cfg = await window.electronAPI.backend.getConfig();
  _port = cfg.port;
  _token = cfg.token;
}

function base(): string {
  return `http://127.0.0.1:${_port}`;
}

function headers(extra?: Record<string, string>): Record<string, string> {
  return {
    'Content-Type': 'application/json',
    'X-API-Token': _token,
    ...extra,
  };
}

async function request<T>(method: string, path: string, body?: unknown): Promise<T> {
  const res = await fetch(`${base()}${path}`, {
    method,
    headers: headers(),
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });

  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error((err as { detail?: string }).detail ?? res.statusText);
  }

  return res.json() as Promise<T>;
}

// ── Endpoints ─────────────────────────────────────────────────────────────────
export async function checkHealth(): Promise<boolean> {
  try {
    const res = await fetch(`${base()}/api/v1/health`);
    return res.ok;
  } catch {
    return false;
  }
}

export async function postExtract(filePath: string): Promise<{ job_id: string }> {
  return request('POST', '/api/v1/extract', { file_path: filePath });
}

export async function postMidiExport(filePath: string, measureRange?: [number, number]): Promise<{ job_id: string }> {
  return request('POST', '/api/v1/midi/export', { file_path: filePath, measure_range: measureRange });
}

export async function postOpenAIEdit(midiPath: string, instructions: string): Promise<{ job_id: string }> {
  return request('POST', '/api/v1/ai/openai/edit-midi', { midi_path: midiPath, instructions });
}

export async function getJob(jobId: string): Promise<Job> {
  return request('GET', `/api/v1/jobs/${jobId}`);
}

export function subscribeJob(
  jobId: string,
  onEvent: (ev: SSEEvent) => void,
  onDone: () => void,
  onError: (err: Error) => void
): () => void {
  // EventSource doesn't support custom headers — pass token as query param
  const url = `${base()}/api/v1/jobs/${jobId}/events?token=${encodeURIComponent(_token)}`;
  const es = new EventSource(url);

  es.addEventListener('message', (e) => {
    try {
      const ev = JSON.parse(e.data) as SSEEvent;
      onEvent(ev);
      if (ev.type === 'done' || ev.type === 'error') {
        es.close();
        onDone();
      }
    } catch {
      /* ignore parse errors */
    }
  });

  es.addEventListener('error', () => {
    es.close();
    onError(new Error('SSE connection failed'));
  });

  return () => es.close();
}
