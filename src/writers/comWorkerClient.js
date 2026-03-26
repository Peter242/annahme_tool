const path = require('path');
const { spawn, spawnSync } = require('child_process');
const readline = require('readline');
const WORKBOOK_READONLY_USER_MESSAGE = 'Annahme.xlsx ist schreibgeschützt oder von einem anderen Benutzer gesperrt. Bitte in Excel schreibbar öffnen oder Sperre lösen.';

class ComWorkerClient {
  constructor(rootDir) {
    this.rootDir = rootDir;
    this.child = null;
    this.rl = null;
    this.pending = new Map();
    this.seq = 0;
    this.starting = null;
    this.disposed = false;
    this.consecutiveTimeouts = 0;
    this.logRing = [];
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

    this.starting = this.spawnAndPing(true);

    try {
      await this.starting;
    } finally {
      this.starting = null;
    }
  }

  async spawnAndPing(allowRetry) {
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

    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');

    child.stdout.on('data', (chunk) => {
      const text = String(chunk || '');
      if (!text) {
        return;
      }
      this.pushLogLine(
        'raw',
        `[stdout-chunk] bytes=${Buffer.byteLength(text, 'utf8')} head=${this.buildLogHead(text, 200)}`,
      );
    });

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
        this.pushLogLine('stdout', trimmed);
        return;
      }
      if (parsed && typeof parsed === 'object' && parsed.type) {
        if (parsed.type === 'step') {
          this.pushLogLine('raw', `[step ${String(parsed.ts || '').trim()}] ${String(parsed.msg || '').trim()}`.trim());
          return;
        }
        if (parsed.type === 'worker') {
          this.pushLogLine('raw', `[worker] ${String(parsed.msg || '').trim()}`.trim());
          return;
        }
        if (parsed.type !== 'response') {
          this.pushLogLine('stdout', trimmed);
          return;
        }
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
      const text = String(chunk || '');
      const lines = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
      for (const line of lines) {
        this.pushLogLine('stderr', line);
        console.error(`[com-worker:stderr] ${line}`);
      }
    });

    if (child.stdin) {
      child.stdin.on('error', (error) => {
        this.pushLogLine('raw', `[stdin-error] msg=${String(error?.message || error || '').trim()}`);
      });
      child.stdin.on('close', () => {
        this.pushLogLine('raw', '[stdin-close]');
      });
      child.stdin.on('finish', () => {
        this.pushLogLine('raw', '[stdin-finish]');
      });
    }

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

    const failStart = async (error) => {
      try {
        if (rl) {
          rl.removeAllListeners();
          rl.close();
        }
      } catch (_error) {
        // ignore
      }
      try {
        child.kill();
      } catch (_error) {
        // ignore
      }
      this.child = null;
      this.rl = null;
      if (allowRetry) {
        return this.spawnAndPing(false);
      }
      throw error;
    };

    try {
      await new Promise((resolve, reject) => {
        child.once('error', (error) => {
          reject(new Error(`COM worker start failed: ${error.message}`));
        });
        setImmediate(resolve);
      });

      this.child = child;
      this.rl = rl;
      const pingResponse = await this.sendRequest({ __ping: true }, { timeoutMs: 2000 });
      if (!pingResponse || pingResponse.ok !== true || pingResponse.status !== 'ready') {
        throw new Error('unexpected ping response');
      }
    } catch (error) {
      return failStart(error);
    }
  }

  async stop() {
    this.disposed = true;
    this.consecutiveTimeouts = 0;
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
    this.consecutiveTimeouts = 0;
    await this.stop();
    this.disposed = wasDisposed ? true : false;
    if (!this.disposed) {
      await this.start();
    }
  }

  async request(payload, options = {}) {
    const timeoutMs = this.resolveTimeoutMs(payload, options);
    const retryOnFailure = options.retryOnFailure !== false;
    const operationName = this.getOperationName(payload);

    const execute = async () => {
      await this.start();
      return this.sendRequest(payload, { timeoutMs, operationName });
    };

    try {
      const parsed = await execute();
      if (parsed && typeof parsed === 'object') {
        const parsedDebug = parsed.debug && typeof parsed.debug === 'object'
          ? { ...parsed.debug }
          : {};
        parsedDebug.logs = this.getRecentLogLines(200);
        parsedDebug.operation = operationName;
        parsed.debug = parsedDebug;
      }
      if (parsed && parsed.ok === false && String(parsed.errorCode || '') === 'WORKBOOK_READONLY') {
        const readonlyError = new Error(`${WORKBOOK_READONLY_USER_MESSAGE} | code=WORKBOOK_READONLY`);
        readonlyError.code = 'WORKBOOK_READONLY';
        readonlyError.userMessage = WORKBOOK_READONLY_USER_MESSAGE;
        readonlyError.debug = { ...parsed };
        throw readonlyError;
      }
      this.consecutiveTimeouts = 0;
      return parsed;
    } catch (error) {
      if (error && String(error.code || '') === 'WORKBOOK_READONLY') {
        throw error;
      }
      if (this.isTimeoutError(error)) {
        this.consecutiveTimeouts += 1;
        error.debug = {
          ...(error.debug || {}),
          operation: operationName,
          timeoutMs,
          lastLogs: this.getRecentLogLines(40),
        };
        if (this.consecutiveTimeouts >= 2) {
          console.warn('[worker-client] second consecutive timeout detected, restarting worker');
          this.consecutiveTimeouts = 0;
          await this.restart();
        }
        throw error;
      }
      error.debug = {
        ...(error.debug || {}),
        operation: operationName,
        timeoutMs,
        lastLogs: this.getRecentLogLines(40),
      };
      this.consecutiveTimeouts = 0;
      if (!retryOnFailure) {
        throw error;
      }
      console.warn(`[worker-client] request failed, retrying once: ${error.message}`);
      await this.restart();
      return execute();
    }
  }

  sendRequest(payload, options = {}) {
    const timeoutMs = Number.isFinite(options.timeoutMs) ? options.timeoutMs : 20000;
    const operationName = String(options.operationName || 'request');
    if (!this.child || !this.child.stdin) {
      throw new Error('COM worker is not running');
    }
    const id = ++this.seq;
    const message = JSON.stringify({ id, payload });
    this.pushLogLine(
      'raw',
      `[stdin-write] id=${id} op=${operationName} bytes=${Buffer.byteLength(message, 'utf8')} head=${this.buildLogHead(message, 200)}`,
    );

    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        const timeoutError = new Error(`COM worker timeout after ${timeoutMs}ms (${operationName})`);
        timeoutError.code = 'COM_WORKER_TIMEOUT';
        timeoutError.debug = {
          operation: operationName,
          timeoutMs,
          lastLogs: this.getRecentLogLines(40),
        };
        reject(timeoutError);
      }, timeoutMs);

      this.pending.set(id, { resolve, reject, timeout });
      this.child.stdin.write(`${message}\n`, 'utf8', (error) => {
        if (!error) {
          this.pushLogLine('raw', `[stdin-write-ok] id=${id}`);
          return;
        }
        const pending = this.pending.get(id);
        if (!pending) {
          return;
        }
        this.pushLogLine('raw', `[stdin-write-error] id=${id} msg=${String(error.message || '').trim()}`);
        clearTimeout(pending.timeout);
        this.pending.delete(id);
        pending.reject(new Error(`COM worker write failed: ${error.message}`));
      });
    });
  }

  resolveTimeoutMs(payload, options = {}) {
    const requestedTimeoutMs = Number.isFinite(options.timeoutMs) ? options.timeoutMs : 0;
    if (payload && payload.__ping === true) {
      return 2000;
    }
    if (payload && payload.__readSheetState === true) {
      return Math.max(requestedTimeoutMs, 8000);
    }
    return Math.max(requestedTimeoutMs, 120000);
  }

  isTimeoutError(error) {
    return String(error?.code || '') === 'COM_WORKER_TIMEOUT'
      || String(error?.message || '').includes('COM worker timeout after');
  }

  getOperationName(payload) {
    if (payload && payload.__ping === true) {
      return 'ping';
    }
    if (payload && payload.__readSheetState === true) {
      return 'sheetCacheInit';
    }
    return 'commit';
  }

  pushLogLine(stream, line) {
    this.logRing.push(
      stream === 'stdout' || stream === 'stderr'
        ? `[${stream}] ${String(line || '')}`
        : String(line || ''),
    );
    if (this.logRing.length > 200) {
      this.logRing.splice(0, this.logRing.length - 200);
    }
  }

  getRecentLogLines(count = 40) {
    const take = Math.max(0, Number.parseInt(String(count), 10) || 0);
    if (take < 1) {
      return [];
    }
    return this.logRing.slice(-take);
  }

  buildLogHead(text, maxLen = 200) {
    const normalized = String(text || '').replace(/\r?\n/g, ' ').trim();
    if (normalized.length <= maxLen) {
      return normalized;
    }
    return `${normalized.slice(0, maxLen)}...`;
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
