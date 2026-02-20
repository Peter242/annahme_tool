const path = require('path');
const { spawn, spawnSync } = require('child_process');
const readline = require('readline');

class ComWorkerClient {
  constructor(rootDir) {
    this.rootDir = rootDir;
    this.child = null;
    this.rl = null;
    this.pending = new Map();
    this.seq = 0;
    this.starting = null;
    this.disposed = false;
  }

  async start() {
    if (this.disposed) {
      throw new Error('COM worker client is disposed');
    }
    if (this.child) {
      return;
    }
    if (this.starting) {
      await this.starting;
      return;
    }

    this.starting = new Promise((resolve, reject) => {
      const scriptPath = path.join(this.rootDir, 'scripts', 'com_worker.ps1');
      const child = spawn('powershell.exe', [
        '-NoLogo',
        '-NoProfile',
        '-NonInteractive',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        scriptPath,
        '-ParentPid',
        String(process.pid),
      ], {
        windowsHide: true,
        detached: false,
        stdio: ['pipe', 'pipe', 'pipe'],
      });

      const failStart = (error) => {
        try {
          child.kill();
        } catch (_error) {
          // ignore
        }
        reject(error);
      };

      child.once('error', (error) => {
        failStart(new Error(`COM worker start failed: ${error.message}`));
      });

      child.stdout.setEncoding('utf8');
      child.stderr.setEncoding('utf8');

      const rl = readline.createInterface({ input: child.stdout });
      rl.on('line', (line) => {
        const trimmed = String(line || '').trim();
        if (!trimmed) {
          return;
        }
        let parsed;
        try {
          parsed = JSON.parse(trimmed);
        } catch (_error) {
          return;
        }
        const id = Number.parseInt(String(parsed.id), 10);
        if (!Number.isFinite(id)) {
          return;
        }
        const pending = this.pending.get(id);
        if (!pending) {
          return;
        }
        this.pending.delete(id);
        clearTimeout(pending.timeout);
        pending.resolve(parsed);
      });

      child.stderr.on('data', (chunk) => {
        const text = String(chunk || '').trim();
        if (text) {
          console.error(`[com-worker:stderr] ${text}`);
        }
      });

      child.on('exit', (code, signal) => {
        this.child = null;
        if (this.rl) {
          this.rl.removeAllListeners();
          this.rl.close();
          this.rl = null;
        }
        const reason = `COM worker exited (code=${String(code)} signal=${String(signal || '')})`;
        for (const [id, pending] of this.pending.entries()) {
          clearTimeout(pending.timeout);
          pending.reject(new Error(reason));
          this.pending.delete(id);
        }
      });

      this.child = child;
      this.rl = rl;
      resolve();
    });

    try {
      await this.starting;
    } finally {
      this.starting = null;
    }
  }

  async stop() {
    this.disposed = true;
    if (!this.child) {
      return;
    }
    const child = this.child;
    this.child = null;
    if (this.rl) {
      this.rl.removeAllListeners();
      this.rl.close();
      this.rl = null;
    }
    for (const [id, pending] of this.pending.entries()) {
      clearTimeout(pending.timeout);
      pending.reject(new Error('COM worker stopped'));
      this.pending.delete(id);
    }
    try {
      child.stdin.end();
    } catch (_error) {
      // ignore
    }
    try {
      child.kill();
    } catch (_error) {
      // ignore
    }
    if (child.pid) {
      try {
        spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], {
          stdio: 'ignore',
        });
      } catch (_error) {
        // ignore
      }
    }
  }

  async restart() {
    console.warn('[worker-client] restarting com worker process');
    const wasDisposed = this.disposed;
    await this.stop();
    this.disposed = wasDisposed ? true : false;
    if (!this.disposed) {
      await this.start();
    }
  }

  async request(payload, options = {}) {
    const timeoutMs = Number.isFinite(options.timeoutMs) ? options.timeoutMs : 20000;
    const retryOnFailure = options.retryOnFailure !== false;

    const execute = async () => {
      await this.start();
      if (!this.child || !this.child.stdin) {
        throw new Error('COM worker is not running');
      }
      const id = ++this.seq;
      const message = JSON.stringify({ id, payload });

      return new Promise((resolve, reject) => {
        const timeout = setTimeout(() => {
          this.pending.delete(id);
          reject(new Error(`COM worker timeout after ${timeoutMs}ms`));
        }, timeoutMs);

        this.pending.set(id, { resolve, reject, timeout });
        this.child.stdin.write(`${message}\n`, 'utf8', (error) => {
          if (!error) {
            return;
          }
          const pending = this.pending.get(id);
          if (!pending) {
            return;
          }
          clearTimeout(pending.timeout);
          this.pending.delete(id);
          pending.reject(new Error(`COM worker write failed: ${error.message}`));
        });
      });
    };

    try {
      return await execute();
    } catch (error) {
      if (!retryOnFailure) {
        throw error;
      }
      console.warn(`[worker-client] request failed, retrying once: ${error.message}`);
      await this.restart();
      return execute();
    }
  }
}

let singleton = null;
let shutdownHooksInstalled = false;

function installShutdownHooks() {
  if (shutdownHooksInstalled) {
    return;
  }
  shutdownHooksInstalled = true;

  const stopClient = async () => {
    if (!singleton) {
      return;
    }
    const client = singleton;
    singleton = null;
    try {
      await client.stop();
    } catch (_error) {
      // ignore shutdown errors
    }
  };

  process.once('SIGINT', () => {
    stopClient().finally(() => process.exit(0));
  });
  process.once('SIGTERM', () => {
    stopClient().finally(() => process.exit(0));
  });
  process.once('exit', () => {
    if (!singleton) {
      return;
    }
    try {
      singleton.stop();
    } catch (_error) {
      // ignore
    }
  });
}

function getComWorkerClient(rootDir) {
  if (!singleton) {
    singleton = new ComWorkerClient(rootDir);
    installShutdownHooks();
  }
  return singleton;
}

module.exports = {
  getComWorkerClient,
};
