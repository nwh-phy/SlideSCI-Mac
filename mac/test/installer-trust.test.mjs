import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, mkdir, writeFile, readFile, readdir, rm, chmod } from 'node:fs/promises';
import { X509Certificate } from 'node:crypto';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const installer = fileURLToPath(new URL('../scripts/distribution.mjs', import.meta.url));
// Stand-in for /usr/bin/security: records trust per SHA-1 fingerprint in a folder, never touches a keychain.
const fakeSecurity = `#!/usr/bin/env node
const fs = require('node:fs'), path = require('node:path'), { X509Certificate } = require('node:crypto');
const state = process.env.FAKE_SECURITY_STATE, args = process.argv.slice(2);
fs.appendFileSync(path.join(state, 'calls.log'), args.join(' ') + '\\n');
const print = file => new X509Certificate(fs.readFileSync(file)).fingerprint.replaceAll(':', '');
const mark = fp => path.join(state, 'trusted-' + fp);
if (args[0] === 'verify-cert') process.exit(fs.existsSync(mark(print(args[args.indexOf('-c') + 1]))) ? 0 : 1);
if (args[0] === 'add-trusted-cert') {
  if (process.env.FAKE_SECURITY_CANCEL) { console.error('SecTrustSettingsSetTrustSettings: The authorization was canceled by the user.'); process.exit(1); }
  fs.writeFileSync(mark(print(args.at(-1))), ''); process.exit(0);
}
if (args[0] === 'delete-certificate') {
  const fp = args[args.indexOf('-Z') + 1];
  if (!fs.existsSync(mark(fp))) { console.error('Unable to delete certificate matching "' + fp + '"'); process.exit(44); }
  fs.rmSync(mark(fp)); process.exit(0);
}
process.exit(2);
`;

async function fixture() {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-trust-'));
  const application = path.join(home, 'Library/Application Support/SlideSCI-Mac');
  const stateDir = path.join(home, 'Library/Application Support/SlideSCI-Mac-Installer');
  const fake = path.join(home, 'fake-security'), securityState = path.join(home, 'security');
  for (const directory of [application, stateDir, securityState]) await mkdir(directory, { recursive: true });
  await writeFile(path.join(application, 'BUILD-INFO.json'), '{}');
  await writeFile(path.join(stateDir, 'installation.json'), JSON.stringify({ application, registrations: [] }));
  await writeFile(fake, fakeSecurity);
  await chmod(fake, 0o755);
  const runInstaller = (command, extra = [], env = {}) => execFileSync(process.execPath, [installer, command, '--profile', home, ...extra], {
    env: { ...process.env, SLIDESCI_SECURITY_TOOL: fake, FAKE_SECURITY_STATE: securityState, ...env }, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
  });
  const tls = path.join(stateDir, 'tls');
  const current = async () => new X509Certificate(await readFile(path.join(tls, 'localhost.crt'))).fingerprint.replaceAll(':', '');
  const trusted = async () => (await readdir(securityState)).filter(name => name.startsWith('trusted-')).map(name => name.slice(8));
  const calls = async () => (await readFile(path.join(securityState, 'calls.log'), 'utf8').catch(() => '')).trim().split('\n').filter(Boolean);
  return { home, runInstaller, current, trusted, calls, tls };
}

test('installer trusts its localhost certificate only with consent, swaps trust on renewal and removes it on uninstall', async () => {
  const { home, runInstaller, current, trusted, calls, tls } = await fixture();
  try {
    let output = runInstaller('renew-certificate', ['--trust-certificate', 'no']);
    assert.match(output, /PowerPoint 需要信任本机证书/);
    assert.deepEqual(await trusted(), [], 'declining never calls add-trusted-cert');
    assert.ok(!(await calls()).some(line => line.startsWith('add-trusted-cert')));

    output = runInstaller('renew-certificate', ['--trust-certificate', 'yes'], { FAKE_SECURITY_CANCEL: '1' });
    assert.match(output, /未完成证书信任：.*canceled/);
    assert.deepEqual(await trusted(), []);

    output = runInstaller('renew-certificate', ['--trust-certificate', 'yes']);
    assert.match(output, /已信任 PowerPoint 使用的本机证书/);
    const first = await current();
    assert.deepEqual(await trusted(), [first]);
    const added = (await calls()).filter(line => line.startsWith('add-trusted-cert')).at(-1);
    assert.match(added, /-r trustRoot -p ssl -s localhost -k .*login\.keychain-db .*localhost\.crt$/, 'trust is limited to SSL for localhost');
    assert.equal(JSON.parse(await readFile(path.join(tls, 'trust.json'), 'utf8')).fingerprint, first);

    runInstaller('renew-certificate', ['--trust-certificate', 'yes']);
    const second = await current();
    assert.notEqual(second, first);
    assert.deepEqual(await trusted(), [second], 'renewal removes the old certificate and its trust');
    assert.ok((await calls()).includes(`delete-certificate -Z ${first} -t ${path.join(home, 'Library/Keychains/login.keychain-db')}`));

    output = runInstaller('uninstall');
    assert.match(output, /已移除旧证书及其信任设置/);
    assert.deepEqual(await trusted(), [], 'uninstall removes the trust it added');
  } finally { await rm(home, { recursive: true, force: true }); }
});

test('an isolated profile never asks to change the signed-in user\'s trust settings', async () => {
  const { home, calls } = await fixture();
  try {
    const output = execFileSync(process.execPath, [installer, 'renew-certificate', '--profile', home, '--trust-certificate', 'yes'], {
      env: { ...process.env, SLIDESCI_SECURITY_TOOL: '' }, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
    });
    assert.match(output, /PowerPoint 需要信任本机证书/);
    assert.doesNotMatch(output, /授权窗口中输入/);
    assert.deepEqual(await calls(), [], 'the stand-in tool is not used, and the real one only checks trust');
  } finally { await rm(home, { recursive: true, force: true }); }
});
