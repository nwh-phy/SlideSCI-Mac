import test from 'node:test';
import assert from 'node:assert/strict';
import { spawn, execFileSync } from 'node:child_process';
import { mkdtemp, mkdir, writeFile, readFile, readdir, rm, utimes, access, chmod } from 'node:fs/promises';
import { closeSync, openSync } from 'node:fs';
import http from 'node:http';
import https from 'node:https';
import net from 'node:net';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const server = fileURLToPath(new URL('../server.mjs', import.meta.url));
const exists = file => access(file).then(() => true, () => false);

async function freePort() {
  const probe = net.createServer();
  await new Promise(resolve => probe.listen(0, '127.0.0.1', resolve));
  const { port } = probe.address();
  await new Promise(resolve => probe.close(resolve));
  return port;
}
function request(url, { method = 'GET', headers = {}, body, ca } = {}) {
  const transport = url.startsWith('https:') ? https : http;
  return new Promise((resolve, reject) => {
    const call = transport.request(url, { method, headers, ca, servername: 'localhost', agent: false }, response => {
      let text = '';
      response.setEncoding('utf8');
      response.on('data', chunk => { text += chunk; });
      response.on('end', () => resolve({ status: response.statusCode, text }));
    });
    call.on('error', reject);
    call.end(body);
  });
}
function firstEvent(url) {
  return new Promise((resolve, reject) => {
    const call = http.get(url, response => {
      let text = '';
      response.setEncoding('utf8');
      response.on('data', chunk => {
        text += chunk;
        const frame = /event: (\w+)\ndata: (.*)\n\n/.exec(text);
        if (frame) { call.destroy(); resolve({ type: frame[1], data: JSON.parse(frame[2]) }); }
      });
    });
    call.on('error', error => { if (error.code !== 'ECONNRESET') reject(error); });
  });
}

test('service lifecycle: loopback listeners, hello, removed grant route, cache sweep and clean exit', { timeout: 60000 }, async () => {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-service-'));
  try {
    const tls = path.join(home, 'tls');
    await mkdir(tls);
    execFileSync('/usr/bin/openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-days', '2', '-subj', '/CN=localhost',
      '-addext', 'subjectAltName=DNS:localhost,IP:127.0.0.1,IP:::1', '-keyout', path.join(tls, 'key.pem'), '-out', path.join(tls, 'cert.pem')], { stdio: 'ignore' });
    const ca = await readFile(path.join(tls, 'cert.pem'));
    const cache = path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/tmp/SlideSCI-Mac');
    await mkdir(cache, { recursive: true });
    const stale = path.join(cache, 'old-export.pdf'), video = path.join(cache, 'old-clip.mp4'), fresh = path.join(cache, 'fresh.png');
    for (const file of [stale, video, fresh]) await writeFile(file, 'x');
    // A host cache the service may not write (macOS refused access to that app's data).
    const wpsCache = path.join(home, 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/SlideSCI-Mac');
    await mkdir(wpsCache, { recursive: true });
    await chmod(wpsCache, 0o500);
    const twoDaysAgo = new Date(Date.now() - 2 * 86400000);
    for (const file of [stale, video]) await utimes(file, twoDaysAgo, twoDaysAgo);

    const httpPort = await freePort(), httpsPort = await freePort();
    const child = spawn(process.execPath, [server, '--cert', path.join(tls, 'cert.pem'), '--key', path.join(tls, 'key.pem')], {
      env: { ...process.env, HOME: home, SLIDESCI_HTTP_PORT: String(httpPort), SLIDESCI_HTTPS_PORT: String(httpsPort), SLIDESCI_ACCEPTANCE: '' },
      stdio: ['ignore', 'pipe', 'pipe']
    });
    let output = '';
    child.stdout.on('data', chunk => { output += chunk; });
    child.stderr.on('data', chunk => { output += chunk; });
    const exited = new Promise(resolve => child.on('exit', code => resolve(code)));
    try {
      const deadline = Date.now() + 15000;
      while (!(output.includes(`127.0.0.1:${httpPort}`) && output.includes(`(::1)`) && output.includes('(127.0.0.1)'))) {
        if (Date.now() > deadline || child.exitCode !== null) throw new Error('service did not start:\n' + output);
        await new Promise(resolve => setTimeout(resolve, 50));
      }
      const base = `http://127.0.0.1:${httpPort}`;
      const session = JSON.parse((await request(base + '/api/session')).text);
      assert.match(session.instance, /^[0-9a-f-]{36}$/);
      assert.deepEqual(session.warnings, [], 'passive startup never probes host containers or raises permission prompts');
      for (const address of ['127.0.0.1', '[::1]']) {
        const secure = JSON.parse((await request(`https://${address}:${httpsPort}/api/session`, { ca })).text);
        assert.equal(secure.token, session.token, `HTTPS on ${address} serves the same session`);
      }
      const grant = await request(base + '/api/files/grant', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-SlideSCI-Token': session.token }, body: JSON.stringify({ paths: ['/etc/hosts'] }) });
      assert.equal(grant.status, 404, 'arbitrary path grants are no longer exposed');
      const aiRoute = await request(base + '/api/ai/profiles', { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-SlideSCI-Token': session.token }, body: '{}' });
      assert.equal(aiRoute.status, 404, 'AI assistant endpoints are removed');
      const foreign = await request(base + '/api/session', { headers: { Host: 'example.com' } });
      assert.equal(foreign.status, 403);
      const hello = await firstEvent(`${base}/api/events?token=${session.token}&host=powerpoint&client=pane-test-0001`);
      assert.deepEqual(hello, { type: 'hello', data: { instance: session.instance, leases: [] } });
      assert.equal(await exists(stale), true, 'startup does not touch PowerPoint container files before an advanced action');
      assert.equal(await exists(video), true);
      assert.equal(await exists(fresh), true);
      const workDirs = (await readdir(os.tmpdir())).filter(name => name.startsWith(`slidesci-mac-${child.pid}-`));
      assert.equal(workDirs.length, 1);
      child.kill('SIGTERM');
      const code = await Promise.race([exited, new Promise(resolve => setTimeout(() => resolve('timeout'), 5000))]);
      assert.equal(code, 0, 'SIGTERM ends the service promptly:\n' + output);
      assert.equal(await exists(path.join(os.tmpdir(), workDirs[0])), false, 'the work folder is removed on exit');
    } finally {
      if (child.exitCode === null) child.kill('SIGKILL');
    }
  } finally {
    await chmod(path.join(home, 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/SlideSCI-Mac'), 0o700).catch(() => {});
    await rm(home, { recursive: true, force: true });
  }
});

test('under launchd a taken port exits cleanly, and the service rotates its own log', { timeout: 60000 }, async () => {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-supervisor-'));
  const blocker = net.createServer();
  try {
    await new Promise(resolve => blocker.listen(0, '127.0.0.1', resolve));
    const taken = blocker.address().port;
    const exitCode = env => new Promise(resolve => {
      const child = spawn(process.execPath, [server], { env: { ...process.env, HOME: home, SLIDESCI_HTTP_PORT: String(taken), SLIDESCI_ACCEPTANCE: '', ...env }, stdio: 'ignore' });
      child.on('exit', code => resolve(code));
    });
    // launchd restarts a non-zero exit every 10 s, which never frees a port held by another program.
    assert.equal(await exitCode({ SLIDESCI_SUPERVISOR: 'launchd' }), 0);
    assert.equal(await exitCode({ SLIDESCI_SUPERVISOR: '' }), 1, 'a manual start still reports the failure');

    const log = path.join(home, 'service.log');
    await writeFile(log, 'x'.repeat(5 * 1024 * 1024 + 16));
    const fd = openSync(log, 'a');
    const httpPort = await freePort();
    const child = spawn(process.execPath, [server], {
      env: { ...process.env, HOME: home, SLIDESCI_HTTP_PORT: String(httpPort), SLIDESCI_ACCEPTANCE: '', SLIDESCI_LOG_FILE: log },
      stdio: ['ignore', fd, fd]
    });
    closeSync(fd);
    const exited = new Promise(resolve => child.on('exit', code => resolve(code)));
    try {
      const deadline = Date.now() + 15000;
      while (!(await readFile(log, 'utf8')).includes(`127.0.0.1:${httpPort}`)) {
        if (Date.now() > deadline || child.exitCode !== null) throw new Error('service did not start');
        await new Promise(resolve => setTimeout(resolve, 50));
      }
      const current = await readFile(log, 'utf8');
      assert.ok(current.length < 4096 && !current.includes('xxxx'), 'the oversized log is truncated before startup lines are written');
      assert.equal((await readFile(log + '.1', 'utf8')).length, 5 * 1024 * 1024 + 16, 'the previous log is kept once');
    } finally {
      child.kill('SIGTERM');
      await exited;
    }
  } finally {
    await new Promise(resolve => blocker.close(resolve));
    await rm(home, { recursive: true, force: true });
  }
});
