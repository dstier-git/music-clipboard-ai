import { app, BrowserWindow, ipcMain } from 'electron';
import path from 'node:path';
import { spawn, ChildProcess } from 'node:child_process';
import crypto from 'node:crypto';
import net from 'node:net';
import started from 'electron-squirrel-startup';

if (started) app.quit();

// ── Backend state ────────────────────────────────────────────────────────────
let backendProcess: ChildProcess | null = null;
let backendPort = 0;
const authToken = crypto.randomBytes(32).toString('hex');

function findFreePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const srv = net.createServer();
    srv.listen(0, '127.0.0.1', () => {
      const addr = srv.address() as net.AddressInfo;
      srv.close(() => resolve(addr.port));
    });
    srv.on('error', reject);
  });
}

function projectRoot(): string {
  // app lives at src/music_clipboard/electron_ui/app
  // project root is 4 levels up from __dirname when running from source
  // When packaged the path is different; rely on ELECTRON_IS_DEV or resource path
  if (process.env.NODE_ENV === 'development' || MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    // __dirname = app/.vite/build  →  6 levels up = project root
    return path.resolve(__dirname, '..', '..', '..', '..', '..', '..');
  }
  return path.dirname(app.getPath('exe'));
}

async function startBackend(port: number): Promise<void> {
  const root = projectRoot();
  const venvPython = path.join(root, 'venv', 'bin', 'python');
  const backendMain = path.join(root, 'src', 'music_clipboard', 'electron_ui', 'backend', 'main.py');

  const env = {
    ...process.env,
    BACKEND_PORT: String(port),
    API_TOKEN: authToken,
    PYTHONPATH: path.join(root, 'src'),
  };

  const pythonExe = require('node:fs').existsSync(venvPython) ? venvPython : 'python3';

  return new Promise((resolve, reject) => {
    backendProcess = spawn(pythonExe, [backendMain], { env, stdio: ['ignore', 'pipe', 'pipe'] });

    const timeout = setTimeout(() => reject(new Error('Backend startup timed out')), 15_000);

    backendProcess.stdout?.on('data', (chunk: Buffer) => {
      const line = chunk.toString();
      console.log('[backend]', line.trim());
      if (line.includes('Application startup complete') || line.includes('READY')) {
        clearTimeout(timeout);
        resolve();
      }
    });

    backendProcess.stderr?.on('data', (chunk: Buffer) => {
      const line = chunk.toString();
      console.error('[backend:err]', line.trim());
      if (line.includes('Application startup complete') || line.includes('READY')) {
        clearTimeout(timeout);
        resolve();
      }
    });

    backendProcess.on('error', (err) => {
      clearTimeout(timeout);
      reject(err);
    });

    backendProcess.on('exit', (code) => {
      if (code !== 0 && code !== null) {
        clearTimeout(timeout);
        reject(new Error(`Backend exited with code ${code}`));
      }
    });
  });
}

function stopBackend(): void {
  if (backendProcess && !backendProcess.killed) {
    backendProcess.kill('SIGTERM');
    backendProcess = null;
  }
}

// ── IPC handlers ─────────────────────────────────────────────────────────────
ipcMain.handle('backend:status', () => ({
  running: backendProcess !== null && !backendProcess.killed,
  port: backendPort,
}));

ipcMain.handle('backend:get-config', () => ({
  port: backendPort,
  token: authToken,
}));

// ── Window ───────────────────────────────────────────────────────────────────
function createWindow(): void {
  const isDev = Boolean(MAIN_WINDOW_VITE_DEV_SERVER_URL);

  const mainWindow = new BrowserWindow({
    width: 1100,
    height: 780,
    minWidth: 800,
    minHeight: 600,
    titleBarStyle: 'hiddenInset',
    backgroundColor: '#0f0f13',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      devTools: isDev,
    },
  });

  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL);
    mainWindow.webContents.openDevTools();
  } else {
    mainWindow.loadFile(path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`));
  }
}

// ── App lifecycle ─────────────────────────────────────────────────────────────
app.whenReady().then(async () => {
  try {
    backendPort = await findFreePort();
    await startBackend(backendPort);
    console.log(`[main] Backend ready on port ${backendPort}`);
  } catch (err) {
    console.error('[main] Backend failed to start:', err);
    // Continue anyway — renderer shows backend-down state
  }

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  stopBackend();
  if (process.platform !== 'darwin') app.quit();
});

app.on('will-quit', () => {
  stopBackend();
});

// Type declarations for Vite-injected globals
declare const MAIN_WINDOW_VITE_DEV_SERVER_URL: string;
declare const MAIN_WINDOW_VITE_NAME: string;
