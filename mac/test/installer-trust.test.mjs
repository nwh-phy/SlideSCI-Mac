import test from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtemp, mkdir, writeFile, readFile, readdir, rm, chmod, symlink } from 'node:fs/promises';
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

test('Start configures LaunchAgent passively without probing host containers, and Uninstall removes it', async () => {
  const { home, runInstaller } = await fixture();
  const cache = path.join(home, 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp/SlideSCI-Mac');
  const launchState = path.join(home, 'launchctl-state');
  const fakeLaunchctl = path.join(home, 'fake-launchctl');
  const fakeLaunchctlScript = `#!/usr/bin/env node
const fs = require('node:fs'), path = require('node:path');
const state = process.env.FAKE_LAUNCHCTL_STATE, args = process.argv.slice(2);
fs.appendFileSync(path.join(state, 'calls.log'), args.join(' ') + '\\n');
const loaded = path.join(state, 'loaded');
if (args[0] === 'print') {
  if (!fs.existsSync(loaded)) process.exit(113);
  // A job that exited cleanly stays loaded without a pid.
  console.log(fs.existsSync(path.join(state, 'exited')) ? '\\tstate = not running' : '\\tstate = running\\n\\tpid = 4242');
  process.exit(0);
}
if (args[0] === 'bootstrap') { fs.writeFileSync(loaded, args.at(-1)); process.exit(0); }
if (args[0] === 'bootout') { if (!fs.existsSync(loaded)) process.exit(3); fs.rmSync(loaded); process.exit(0); }
process.exit(2);
`;
  try {
    await mkdir(launchState, { recursive: true });
    await writeFile(fakeLaunchctl, fakeLaunchctlScript);
    await chmod(fakeLaunchctl, 0o755);
    runInstaller('renew-certificate', ['--trust-certificate', 'no']);
    await mkdir(cache, { recursive: true });
    await chmod(cache, 0o500);
    const output = runInstaller('start', ['--trust-certificate', 'no'], { SLIDESCI_LAUNCHCTL: fakeLaunchctl, FAKE_LAUNCHCTL_STATE: launchState });
    assert.doesNotMatch(output, /拒绝“终端”访问/, 'Start stays passive and does not probe host containers');
    assert.match(output, /登录自动运行/);
    const plist = path.join(home, 'Library/LaunchAgents/com.slidesci.mac.service.plist');
    const plistText = await readFile(plist, 'utf8');
    assert.match(plistText, /<key>RunAtLoad<\/key><true\/>/);
    assert.match(plistText, /com\.slidesci\.mac\.service/);
    const stopOut = runInstaller('stop', [], { SLIDESCI_LAUNCHCTL: fakeLaunchctl, FAKE_LAUNCHCTL_STATE: launchState });
    assert.match(stopOut, /已停止 SlideSCI 后台自启服务/);
    assert.match(stopOut, /取消登录自动运行/);
    assert.equal(await readFile(plist, 'utf8').then(() => true, () => false), false, 'Stop removes the login item so the service stays stopped after the next login');
    runInstaller('uninstall', [], { SLIDESCI_LAUNCHCTL: fakeLaunchctl, FAKE_LAUNCHCTL_STATE: launchState });
    assert.equal(await readFile(plist, 'utf8').then(() => true, () => false), false, 'uninstall removes LaunchAgent plist');
  } finally { await chmod(cache, 0o700).catch(() => {}); await rm(home, { recursive: true, force: true }); }
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

async function makePackageFixture(directory, { version, manifestXml }) {
  await mkdir(path.join(directory, 'runtime/bin'), { recursive: true });
  await mkdir(path.join(directory, 'mac/scripts'), { recursive: true });
  await mkdir(path.join(directory, 'mac/office'), { recursive: true });
  await writeFile(path.join(directory, 'BUILD-INFO.json'), JSON.stringify({ version }, null, 2));
  await writeFile(path.join(directory, 'runtime/bin/node'), '#!/bin/sh\nexit 0\n');
  await chmod(path.join(directory, 'runtime/bin/node'), 0o755);
  await writeFile(path.join(directory, 'mac/scripts/distribution.mjs'), await readFile(installer, 'utf8'));
  await writeFile(path.join(directory, 'mac/office/manifest.xml'), manifestXml);
  await symlink(path.resolve(path.dirname(installer), '../node_modules'), path.join(directory, 'mac/node_modules'));
}

function makeFakeLaunchctl() {
  return `#!/usr/bin/env node
const fs = require('node:fs'), path = require('node:path');
const state = process.env.FAKE_LAUNCHCTL_STATE, args = process.argv.slice(2);
fs.appendFileSync(path.join(state, 'calls.log'), args.join(' ') + '\\n');
const loaded = path.join(state, 'loaded');
if (args[0] === 'print') {
  if (!fs.existsSync(loaded)) process.exit(113);
  // A job that exited cleanly stays loaded without a pid.
  console.log(fs.existsSync(path.join(state, 'exited')) ? '\\tstate = not running' : '\\tstate = running\\n\\tpid = 4242');
  process.exit(0);
}
if (args[0] === 'bootstrap') {
  if (process.env.FAKE_LAUNCHCTL_FAIL_BOOTSTRAP) { console.error('Bootstrap failed: 5: Input/output error'); process.exit(5); }
  if (process.env.FAKE_LAUNCHCTL_FAIL_VERSION) {
    const application = path.join(state, '../Library/Application Support/SlideSCI-Mac');
    if (require(path.join(application, 'BUILD-INFO.json')).version === process.env.FAKE_LAUNCHCTL_FAIL_VERSION) {
      console.error('Bootstrap failed for the new version'); process.exit(5);
    }
  }
  fs.writeFileSync(loaded, args.at(-1)); process.exit(0);
}
if (args[0] === 'bootout') {
  if (process.env.FAKE_LAUNCHCTL_FAIL_BOOTOUT) { console.error('Boot-out failed: 5: Input/output error'); process.exit(5); }
  if (!fs.existsSync(loaded)) process.exit(3);
  fs.rmSync(loaded); process.exit(0);
}
process.exit(2);
`;
}

const priorManifest = '<OfficeApp><Version>0.9.0.0</Version></OfficeApp>';
const v1Manifest = '<OfficeApp><Version>1.1.2.0</Version></OfficeApp>';
const v2Manifest = '<OfficeApp><Version>1.2.0.0</Version></OfficeApp>';

async function upgradeEnvironment(prefix) {
  const home = await mkdtemp(path.join(os.tmpdir(), prefix));
  const wefDir = path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef');
  const securityState = path.join(home, 'security');
  const launchState = path.join(home, 'launchctl');
  const fakeSecurityTool = path.join(home, 'fake-security');
  const fakeLaunchctl = path.join(home, 'fake-launchctl');
  const v1Pkg = path.join(home, 'pkg-v1');
  const v2Pkg = path.join(home, 'pkg-v2');
  for (const directory of [wefDir, securityState, launchState]) await mkdir(directory, { recursive: true });
  await writeFile(path.join(wefDir, 'SlideSCI-Mac.xml'), priorManifest);
  await writeFile(fakeSecurityTool, fakeSecurity);
  await chmod(fakeSecurityTool, 0o755);
  await writeFile(fakeLaunchctl, makeFakeLaunchctl());
  await chmod(fakeLaunchctl, 0o755);
  await makePackageFixture(v1Pkg, { version: '0.1.2', manifestXml: v1Manifest });
  await makePackageFixture(v2Pkg, { version: '0.2.0', manifestXml: v2Manifest });
  const env = {
    ...process.env,
    SLIDESCI_SECURITY_TOOL: fakeSecurityTool,
    FAKE_SECURITY_STATE: securityState,
    SLIDESCI_LAUNCHCTL: fakeLaunchctl,
    FAKE_LAUNCHCTL_STATE: launchState,
    SLIDESCI_RUNNING_HOSTS: ''
  };
  const runPkg = (pkg, command, extra = [], overrides = {}) => execFileSync(process.execPath, [path.join(pkg, 'mac/scripts/distribution.mjs'), command, '--profile', home, ...extra], {
    env: { ...env, ...overrides }, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
  });
  const launchCalls = async () => (await readFile(path.join(launchState, 'calls.log'), 'utf8').catch(() => '')).trim().split('\n').filter(Boolean);
  const trusted = async () => (await readdir(securityState)).filter(name => name.startsWith('trusted-')).map(name => name.slice(8));
  const installation = async () => JSON.parse(await readFile(path.join(home, 'Library/Application Support/SlideSCI-Mac-Installer/installation.json'), 'utf8'));
  const manifest = () => readFile(path.join(wefDir, 'SlideSCI-Mac.xml'), 'utf8');
  const buildVersion = async () => JSON.parse(await readFile(path.join(home, 'Library/Application Support/SlideSCI-Mac/BUILD-INFO.json'), 'utf8')).version;
  return { home, wefDir, launchState, v1Pkg, v2Pkg, runPkg, launchCalls, trusted, installation, manifest, buildVersion };
}

// Builds a localhost certificate whose notAfter is in the past so certificate renewal can be exercised.
async function makeExpiredCertificate(certFile, keyFile) {
  const work = await mkdtemp(path.join(os.tmpdir(), 'slidesci-expired-'));
  try {
    await mkdir(path.join(work, 'newcerts'));
    await writeFile(path.join(work, 'index.txt'), '');
    await writeFile(path.join(work, 'serial'), '1000\n');
    await writeFile(path.join(work, 'ca.cnf'), [
      '[ca]', 'default_ca = CA_default', '[CA_default]', 'dir = .', 'database = ./index.txt',
      'new_certs_dir = ./newcerts', 'serial = ./serial', 'default_md = sha256',
      'policy = policy_any', 'x509_extensions = ext', '[policy_any]', 'commonName = supplied',
      '[ext]', 'subjectAltName = DNS:localhost,IP:127.0.0.1,IP:::1',
      'basicConstraints = critical,CA:FALSE', 'keyUsage = critical,digitalSignature,keyEncipherment',
      'extendedKeyUsage = serverAuth', ''
    ].join('\n'));
    const key = path.join(work, 'key.pem'), request = path.join(work, 'request.pem'), output = path.join(work, 'cert.pem');
    execFileSync('/usr/bin/openssl', ['req', '-new', '-newkey', 'rsa:2048', '-nodes', '-subj', '/CN=localhost', '-keyout', key, '-out', request], { stdio: 'ignore' });
    execFileSync('/usr/bin/openssl', ['ca', '-batch', '-config', path.join(work, 'ca.cnf'), '-selfsign', '-keyfile', key, '-in', request, '-out', output, '-startdate', '20200101000000Z', '-enddate', '20200102000000Z'], { cwd: work, stdio: 'ignore' });
    await writeFile(certFile, await readFile(output));
    await writeFile(keyFile, await readFile(key));
  } finally { await rm(work, { recursive: true, force: true }); }
}

function failureOutput(callback) {
  try { callback(); }
  catch (error) { return `${error.stdout || ''}${error.stderr || ''}${error.message || ''}`; }
  assert.fail('expected the command to fail');
}

test('install upgrades in place, stops old service, reuses trusted cert, clears Wef cache and preserves original prior on uninstall', async () => {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-upgrade-'));
  const wefDir = path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef');
  const wefCache = path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef');
  const securityState = path.join(home, 'security');
  const launchState = path.join(home, 'launchctl');
  const fakeSecurityTool = path.join(home, 'fake-security');
  const fakeLaunchctl = path.join(home, 'fake-launchctl');
  const v1Pkg = path.join(home, 'pkg-v1');
  const v2Pkg = path.join(home, 'pkg-v2');
  const originalPrior = '<OfficeApp><Version>0.9.0.0</Version></OfficeApp>';
  const v1Manifest = '<OfficeApp><Version>1.1.2.0</Version></OfficeApp>';
  const v2Manifest = '<OfficeApp><Version>1.2.0.0</Version></OfficeApp>';

  try {
    for (const dir of [wefDir, wefCache, securityState, launchState]) await mkdir(dir, { recursive: true });
    await writeFile(path.join(wefDir, 'SlideSCI-Mac.xml'), originalPrior);
    await writeFile(path.join(wefCache, 'stale-cache.json'), '{}');
    await writeFile(fakeSecurityTool, fakeSecurity);
    await chmod(fakeSecurityTool, 0o755);
    await writeFile(fakeLaunchctl, makeFakeLaunchctl());
    await chmod(fakeLaunchctl, 0o755);
    await makePackageFixture(v1Pkg, { version: '0.1.2', manifestXml: v1Manifest });
    await makePackageFixture(v2Pkg, { version: '0.2.0', manifestXml: v2Manifest });

    const env = {
      ...process.env,
      SLIDESCI_SECURITY_TOOL: fakeSecurityTool,
      FAKE_SECURITY_STATE: securityState,
      SLIDESCI_LAUNCHCTL: fakeLaunchctl,
      FAKE_LAUNCHCTL_STATE: launchState,
      SLIDESCI_RUNNING_HOSTS: ''
    };
    const runPkg = (pkg, cmd, extra = []) => execFileSync(process.execPath, [path.join(pkg, 'mac/scripts/distribution.mjs'), cmd, '--profile', home, ...extra], {
      env, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
    });

    runPkg(v1Pkg, 'install', ['--trust-certificate', 'yes', '--clear-wef-cache', 'no']);
    assert.equal(await readFile(path.join(wefDir, 'SlideSCI-Mac.xml'), 'utf8'), v1Manifest);

    const upgradeOut = runPkg(v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no', '--clear-wef-cache', 'yes']);
    assert.match(upgradeOut, /SlideSCI 已升级到 0\.2\.0/);
    assert.match(upgradeOut, /已清理 PowerPoint 加载项缓存（Wef）/);
    assert.match(upgradeOut, /SlideSCI-Mac\.previous-/);

    const appSupport = await readdir(path.join(home, 'Library/Application Support'));
    assert.ok(appSupport.some(name => name.startsWith('SlideSCI-Mac.previous-')), 'old install directory is backed up');
    assert.equal(await readFile(path.join(home, 'Library/Application Support/SlideSCI-Mac/BUILD-INFO.json'), 'utf8').then(t => JSON.parse(t).version), '0.2.0');
    assert.equal(await readFile(path.join(wefDir, 'SlideSCI-Mac.xml'), 'utf8'), v2Manifest);
    assert.equal(await readFile(path.join(wefCache, 'stale-cache.json'), 'utf8').then(() => true, () => false), false, 'Wef cache cleared on upgrade when PowerPoint is closed');

    const secCalls = (await readFile(path.join(securityState, 'calls.log'), 'utf8')).trim().split('\n');
    assert.equal(secCalls.filter(line => line.startsWith('add-trusted-cert')).length, 1, 'upgrade reuses existing trusted certificate without prompting again');

    const launchCalls = (await readFile(path.join(launchState, 'calls.log'), 'utf8')).trim().split('\n');
    assert.deepEqual(launchCalls.filter(line => line.startsWith('boot')), [
      'bootstrap gui/' + process.getuid() + ' ' + path.join(home, 'Library/LaunchAgents/com.slidesci.mac.service.plist'),
      'bootout gui/' + process.getuid() + '/com.slidesci.mac.service',
      'bootstrap gui/' + process.getuid() + ' ' + path.join(home, 'Library/LaunchAgents/com.slidesci.mac.service.plist')
    ]);

    runPkg(v2Pkg, 'uninstall');
    assert.equal(await readFile(path.join(wefDir, 'SlideSCI-Mac.xml'), 'utf8'), originalPrior, 'uninstall after upgrade still restores original pre-SlideSCI manifest');
  } finally { await rm(home, { recursive: true, force: true }); }
});

test('upgrade protects running PowerPoint Wef cache, respects declined upgrade and rejects tampered registrations', async () => {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-upgrade-guard-'));
  const wefDir = path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef');
  const wefCache = path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef');
  const securityState = path.join(home, 'security');
  const launchState = path.join(home, 'launchctl');
  const fakeSecurityTool = path.join(home, 'fake-security');
  const fakeLaunchctl = path.join(home, 'fake-launchctl');
  const v1Pkg = path.join(home, 'pkg-v1');
  const v2Pkg = path.join(home, 'pkg-v2');
  const v1Manifest = '<OfficeApp><Version>1.1.2.0</Version></OfficeApp>';
  const v2Manifest = '<OfficeApp><Version>1.2.0.0</Version></OfficeApp>';

  try {
    for (const dir of [wefDir, wefCache, securityState, launchState]) await mkdir(dir, { recursive: true });
    await writeFile(path.join(wefCache, 'stale-cache.json'), '{}');
    await writeFile(fakeSecurityTool, fakeSecurity);
    await chmod(fakeSecurityTool, 0o755);
    await writeFile(fakeLaunchctl, makeFakeLaunchctl());
    await chmod(fakeLaunchctl, 0o755);
    await makePackageFixture(v1Pkg, { version: '0.1.2', manifestXml: v1Manifest });
    await makePackageFixture(v2Pkg, { version: '0.2.0', manifestXml: v2Manifest });

    const runPkg = (pkg, extra = [], running = '') => execFileSync(process.execPath, [path.join(pkg, 'mac/scripts/distribution.mjs'), 'install', '--profile', home, ...extra], {
      env: { ...process.env, SLIDESCI_SECURITY_TOOL: fakeSecurityTool, FAKE_SECURITY_STATE: securityState, SLIDESCI_LAUNCHCTL: fakeLaunchctl, FAKE_LAUNCHCTL_STATE: launchState, SLIDESCI_RUNNING_HOSTS: running },
      encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
    });

    runPkg(v1Pkg, ['--trust-certificate', 'no', '--clear-wef-cache', 'no']);

    assert.throws(() => runPkg(v2Pkg, ['--upgrade', 'no', '--trust-certificate', 'no']), /已取消原地升级/);
    assert.equal(JSON.parse(await readFile(path.join(home, 'Library/Application Support/SlideSCI-Mac/BUILD-INFO.json'), 'utf8')).version, '0.1.2');

    const runningOut = runPkg(v2Pkg, ['--upgrade', 'yes', '--trust-certificate', 'no'], 'PowerPoint');
    assert.match(runningOut, /检测到 PowerPoint 正在运行/);
    assert.match(runningOut, /未自动清理 Wef 加载项缓存/);
    assert.equal(await readFile(path.join(wefCache, 'stale-cache.json'), 'utf8'), '{}', 'Wef cache is kept while PowerPoint is running');

    await writeFile(path.join(wefDir, 'SlideSCI-Mac.xml'), '<OfficeApp><Tampered/></OfficeApp>');
    assert.throws(() => runPkg(v2Pkg, ['--upgrade', 'yes', '--trust-certificate', 'no']), /PowerPoint 注册已被另行修改/);
  } finally { await rm(home, { recursive: true, force: true }); }
});

test('Stop fails and blocks upgrade or uninstall before changes when bootout fails while print reports loaded', async () => {
  const fixture = await upgradeEnvironment('slidesci-stop-guard-');
  try {
    fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'no', '--clear-wef-cache', 'no']);
    const installationBefore = await fixture.installation();
    const bootstrapsBefore = (await fixture.launchCalls()).filter(line => line.startsWith('bootstrap ')).length;
    const failBootout = { FAKE_LAUNCHCTL_FAIL_BOOTOUT: '1' };

    let output = failureOutput(() => fixture.runPkg(fixture.v1Pkg, 'stop', [], failBootout));
    assert.match(output, /无法停止 SlideSCI 后台自启服务/);
    assert.equal(await fixture.buildVersion(), '0.1.2');

    output = failureOutput(() => fixture.runPkg(fixture.v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no'], failBootout));
    assert.match(output, /无法停止 SlideSCI 后台自启服务/);
    assert.equal(await fixture.buildVersion(), '0.1.2', 'upgrade leaves the installed app untouched');
    assert.equal(await fixture.manifest(), v1Manifest, 'upgrade leaves registrations untouched');
    assert.deepEqual(await fixture.installation(), installationBefore, 'upgrade leaves the installation record untouched');
    assert.equal((await fixture.launchCalls()).filter(line => line.startsWith('bootstrap ')).length, bootstrapsBefore);

    output = failureOutput(() => fixture.runPkg(fixture.v1Pkg, 'uninstall', [], failBootout));
    assert.match(output, /无法停止 SlideSCI 后台自启服务/);
    assert.equal(await fixture.buildVersion(), '0.1.2', 'uninstall leaves the installed app in place');
    assert.equal(await fixture.manifest(), v1Manifest, 'uninstall leaves registrations in place');
    assert.equal(await readFile(path.join(fixture.home, 'Library/LaunchAgents/com.slidesci.mac.service.plist'), 'utf8').then(() => true, () => false), true);
  } finally { await rm(fixture.home, { recursive: true, force: true }); }
});

test('a failed upgrade restores the previous app, registrations, record and running service', async () => {
  const fixture = await upgradeEnvironment('slidesci-upgrade-rollback-');
  const unreadable = path.join(fixture.v2Pkg, 'unreadable');
  try {
    await mkdir(unreadable);
    await writeFile(path.join(unreadable, 'payload.txt'), 'payload');
    await chmod(unreadable, 0);
    fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'yes', '--clear-wef-cache', 'no']);
    const bootstrapsBefore = (await fixture.launchCalls()).filter(line => line.startsWith('bootstrap ')).length;

    const output = failureOutput(() => fixture.runPkg(fixture.v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no']));
    assert.match(output, /EACCES|Permission denied/);
    assert.match(output, /已恢复升级前正在运行的 SlideSCI 后台服务/);
    assert.equal(await fixture.buildVersion(), '0.1.2');
    assert.equal(await fixture.manifest(), v1Manifest);
    assert.equal((await fixture.installation()).version, '0.1.2');
    const launchCalls = (await fixture.launchCalls()).filter(line => line.startsWith('boot'));
    assert.deepEqual(launchCalls.slice(bootstrapsBefore), [
      `bootout gui/${process.getuid()}/com.slidesci.mac.service`,
      `bootstrap gui/${process.getuid()} ${path.join(fixture.home, 'Library/LaunchAgents/com.slidesci.mac.service.plist')}`
    ], 'the service stopped for the upgrade is started again after rollback');
  } finally { await chmod(unreadable, 0o700).catch(() => {}); await rm(fixture.home, { recursive: true, force: true }); }
});

test('upgrade rolls back if the new version cannot start the previously running service', async () => {
  const fixture = await upgradeEnvironment('slidesci-upgrade-start-failure-');
  try {
    fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'no']);
    const prior = await fixture.installation();
    const output = failureOutput(() => fixture.runPkg(fixture.v2Pkg, 'install',
      ['--upgrade', 'yes', '--trust-certificate', 'no'], { FAKE_LAUNCHCTL_FAIL_VERSION: '0.2.0' }));
    assert.match(output, /Bootstrap failed for the new version/);
    assert.doesNotMatch(output, /已升级到/);
    assert.equal(await fixture.buildVersion(), '0.1.2');
    assert.equal(await fixture.manifest(), v1Manifest);
    assert.deepEqual(await fixture.installation(), prior);
    await readFile(path.join(fixture.launchState, 'loaded'));
  } finally { await rm(fixture.home, { recursive: true, force: true }); }
});

test('rollback restoration failures are reported together with the original upgrade error', async () => {
  const fixture = await upgradeEnvironment('slidesci-upgrade-rollback-error-');
  try {
    fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'no', '--clear-wef-cache', 'no']);
    const plist = path.join(fixture.home, 'Library/LaunchAgents/com.slidesci.mac.service.plist');
    await rm(plist);
    await mkdir(path.join(plist, 'blocked'), { recursive: true });

    const output = failureOutput(() => fixture.runPkg(fixture.v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no']));
    assert.match(output, /EISDIR|ENOTEMPTY|directory/i, 'the original write failure is kept');
    assert.match(output, /回滚未完全完成/);
    assert.match(output, /重新启动后台服务/);
    assert.equal(await fixture.buildVersion(), '0.1.2');
    assert.equal(await fixture.manifest(), v1Manifest);
    assert.equal((await fixture.installation()).version, '0.1.2');
  } finally { await rm(fixture.home, { recursive: true, force: true }); }
});

test('Start repairs certificate trust without restarting a service left running after an unanswered prompt', async () => {
  const cases = [
    { name: 'declined', args: ['--trust-certificate', 'no'], env: {} },
    { name: 'canceled', args: ['--trust-certificate', 'yes'], env: { FAKE_SECURITY_CANCEL: '1' } }
  ];
  for (const item of cases) {
    const fixture = await upgradeEnvironment(`slidesci-trust-repair-${item.name}-`);
    try {
      fixture.runPkg(fixture.v1Pkg, 'install', [...item.args, '--clear-wef-cache', 'no'], item.env);
      assert.deepEqual(await fixture.trusted(), [], `${item.name} authorization leaves the certificate untrusted`);
      const loaded = await readFile(path.join(fixture.launchState, 'loaded'), 'utf8');
      const bootCalls = (await fixture.launchCalls()).filter(line => line.startsWith('boot'));

      const output = fixture.runPkg(fixture.v1Pkg, 'start', ['--trust-certificate', 'yes']);
      assert.match(output, /已信任 PowerPoint 使用的本机证书/);
      assert.match(output, /已经运行/);
      const current = new X509Certificate(await readFile(path.join(fixture.home, 'Library/Application Support/SlideSCI-Mac-Installer/tls/localhost.crt'))).fingerprint.replaceAll(':', '');
      assert.deepEqual(await fixture.trusted(), [current]);
      assert.equal(await readFile(path.join(fixture.launchState, 'loaded'), 'utf8'), loaded, 'the running agent is not replaced');
      assert.deepEqual((await fixture.launchCalls()).filter(line => line.startsWith('boot')), bootCalls, 'trusting the certificate does not restart the service');
    } finally { await rm(fixture.home, { recursive: true, force: true }); }
  }
});

test('Start renews an expired certificate and restarts the running service so the renewed TLS is served', async () => {
  const fixture = await upgradeEnvironment('slidesci-expired-active-');
  try {
    fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'yes', '--clear-wef-cache', 'no']);
    const tls = path.join(fixture.home, 'Library/Application Support/SlideSCI-Mac-Installer/tls');
    await makeExpiredCertificate(path.join(tls, 'localhost.crt'), path.join(tls, 'localhost.key'));
    const expired = new X509Certificate(await readFile(path.join(tls, 'localhost.crt')));

    const output = fixture.runPkg(fixture.v1Pkg, 'start', ['--trust-certificate', 'yes']);
    assert.match(output, /已过期，正在生成新证书/);
    const renewed = new X509Certificate(await readFile(path.join(tls, 'localhost.crt')));
    assert.notEqual(renewed.fingerprint, expired.fingerprint);
    assert.ok(Date.parse(renewed.validTo) > Date.now());
    assert.equal((await readdir(tls)).filter(name => name.startsWith('retired-')).length, 1);
    const launchCalls = (await fixture.launchCalls()).filter(line => line.startsWith('boot'));
    assert.deepEqual(launchCalls.slice(-2), [
      `bootout gui/${process.getuid()}/com.slidesci.mac.service`,
      `bootstrap gui/${process.getuid()} ${path.join(fixture.home, 'Library/LaunchAgents/com.slidesci.mac.service.plist')}`
    ], 'the old service is stopped before the renewed certificate is served');
    assert.ok((await fixture.trusted()).includes(renewed.fingerprint.replaceAll(':', '')));
  } finally { await rm(fixture.home, { recursive: true, force: true }); }
});

test('Start leaves the WPS service running if the replacement TLS pair cannot be staged', async () => {
  const fixture = await upgradeEnvironment('slidesci-expired-stage-failure-');
  const tls = path.join(fixture.home, 'Library/Application Support/SlideSCI-Mac-Installer/tls');
  const cert = path.join(tls, 'localhost.crt'), key = path.join(tls, 'localhost.key');
  try {
    fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'no']);
    await makeExpiredCertificate(cert, key);
    const oldPair = await Promise.all([readFile(cert), readFile(key)]);
    const boots = (await fixture.launchCalls()).filter(line => line.startsWith('boot'));
    await chmod(tls, 0o500);
    assert.match(failureOutput(() => fixture.runPkg(fixture.v1Pkg, 'start')), /EACCES|EPERM/);
    assert.deepEqual((await fixture.launchCalls()).filter(line => line.startsWith('boot')), boots);
    assert.deepEqual(await Promise.all([readFile(cert), readFile(key)]), oldPair);
    await readFile(path.join(fixture.launchState, 'loaded'));
  } finally {
    await chmod(tls, 0o700);
    await rm(fixture.home, { recursive: true, force: true });
  }
});

test('upgrades keep one previous folder, ask before clearing the Wef cache, and Start relaunches an exited login item', async () => {
  const fixture = await upgradeEnvironment('slidesci-upgrade-hygiene-');
  const { home, launchState, v1Pkg, v2Pkg, runPkg, launchCalls } = fixture;
  const wefCache = path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef');
  const support = path.join(home, 'Library/Application Support');
  const previous = async () => (await readdir(support)).filter(name => name.startsWith('SlideSCI-Mac.previous-'));
  try {
    await mkdir(wefCache, { recursive: true });
    await writeFile(path.join(wefCache, 'stale-cache.json'), '{}');
    runPkg(v1Pkg, 'install', ['--trust-certificate', 'no']);
    // Without an answer (no terminal) the shared Wef cache is kept and the manual command is shown.
    const first = runPkg(v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no']);
    assert.match(first, /未清理 Wef 加载项缓存/);
    assert.equal(await readFile(path.join(wefCache, 'stale-cache.json'), 'utf8'), '{}', 'other add-ins keep their cache without consent');
    assert.equal((await previous()).length, 1);
    await new Promise(resolve => setTimeout(resolve, 5));
    const second = runPkg(v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no']);
    assert.match(second, /已删除更早的旧版本备份/);
    const kept = await previous();
    assert.equal(kept.length, 1, 'only the latest previous version is kept');
    assert.ok(second.includes(path.join(support, kept[0])), 'the kept folder is the one this upgrade backed up');

    // A login item that exited cleanly (for example, its port was taken at login) stays loaded but not running.
    await writeFile(path.join(launchState, 'exited'), '');
    const before = (await launchCalls()).filter(line => line.startsWith('boot')).length;
    const restart = runPkg(v2Pkg, 'start', ['--trust-certificate', 'no']);
    assert.doesNotMatch(restart, /已经运行/);
    assert.match(restart, /已退出，正在重新启动/);
    assert.deepEqual((await launchCalls()).filter(line => line.startsWith('boot')).slice(before).map(line => line.split(' ')[0]), ['bootout', 'bootstrap']);
    const plist = await readFile(path.join(home, 'Library/LaunchAgents/com.slidesci.mac.service.plist'), 'utf8');
    assert.match(plist, /<key>SLIDESCI_SUPERVISOR<\/key><string>launchd<\/string>/);
    assert.match(plist, /<key>SLIDESCI_LOG_FILE<\/key><string>[^<]*service\.log<\/string>/);
  } finally { await rm(home, { recursive: true, force: true }); }
});

test('a browser-downloaded package installs without the quarantine mark, so Gatekeeper does not kill the native helper', async () => {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-quarantine-'));
  const pkg = path.join(home, 'downloaded package');
  const quarantined = ['BUILD-INFO.json', 'runtime/bin/node', 'mac/bin/slidesci-native', 'mac/scripts/distribution.mjs', 'mac/office/manifest.xml'];
  const mark = file => execFileSync('/usr/bin/xattr', ['-p', 'com.apple.quarantine', file], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] }).trim();
  const marked = file => { try { mark(file); return true; } catch { return false; } };
  try {
    await mkdir(path.join(home, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef'), { recursive: true });
    await makePackageFixture(pkg, { version: '0.2.1', manifestXml: '<OfficeApp><Version>1.2.0.0</Version></OfficeApp>' });
    await mkdir(path.join(pkg, 'mac/bin'), { recursive: true });
    await writeFile(path.join(pkg, 'mac/bin/slidesci-native'), '#!/bin/sh\nexit 0\n');
    await chmod(path.join(pkg, 'mac/bin/slidesci-native'), 0o755);
    // What Safari and Archive Utility leave on every extracted file (the linked node_modules stays untouched).
    for (const file of quarantined) execFileSync('/usr/bin/xattr', ['-w', 'com.apple.quarantine', `0083;${Math.floor(Date.now() / 1000).toString(16)};Safari;`, path.join(pkg, file)]);
    assert.ok(quarantined.every(file => marked(path.join(pkg, file))), 'the fixture is quarantined like a download');

    const output = execFileSync(process.execPath, [path.join(pkg, 'mac/scripts/distribution.mjs'), 'install', '--profile', home, '--trust-certificate', 'no'], {
      env: { ...process.env, SLIDESCI_RUNNING_HOSTS: '' }, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe']
    });
    assert.match(output, /SlideSCI 已安装/);
    const application = path.join(home, 'Library/Application Support/SlideSCI-Mac');
    for (const file of quarantined) assert.equal(marked(path.join(application, file)), false, `${file} is installed without quarantine`);
    assert.ok(marked(path.join(pkg, 'mac/bin/slidesci-native')), 'the downloaded folder itself is left as it was');
  } finally { await rm(home, { recursive: true, force: true }); }
});

test('a host container macOS refuses to open stops install and upgrade before any change, instead of counting as missing', async () => {
  const fixture = await upgradeEnvironment('slidesci-denied-container-');
  const data = path.join(fixture.home, 'Library/Containers/com.microsoft.Powerpoint/Data');
  const jsaddons = path.join(fixture.home, 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons');
  const publish = path.join(jsaddons, 'publish.xml');
  const installed = () => readFile(path.join(fixture.home, 'Library/Application Support/SlideSCI-Mac/BUILD-INFO.json')).then(() => true, () => false);
  try {
    await mkdir(jsaddons, { recursive: true });
    await writeFile(publish, '<jsplugins/>');
    // What a "Don't Allow" on macOS's app-data prompt looks like to the installer.
    await chmod(data, 0);
    let output = failureOutput(() => fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'no']));
    assert.match(output, /macOS 拒绝“终端”访问 PowerPoint 的数据，安装器无法读写其中的 SlideSCI 加载项注册/);
    assert.doesNotMatch(output, /尚未找到 PowerPoint 或 WPS 的容器/);
    assert.equal(await readFile(publish, 'utf8'), '<jsplugins/>', 'WPS is not registered on its own');
    assert.equal(await installed(), false, 'nothing is installed');

    await chmod(data, 0o700);
    fixture.runPkg(fixture.v1Pkg, 'install', ['--trust-certificate', 'no']);
    const before = await fixture.installation();
    assert.deepEqual(before.registrations.map(item => item.host).sort(), ['powerpoint', 'wps']);
    await chmod(data, 0);
    output = failureOutput(() => fixture.runPkg(fixture.v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no']));
    assert.match(output, /macOS 拒绝“终端”访问 PowerPoint 的数据/);
    assert.deepEqual(await fixture.installation(), before, 'the record keeps both registrations');
    assert.equal(await fixture.buildVersion(), '0.1.2');
    assert.equal((await fixture.launchCalls()).filter(line => line.startsWith('bootout')).length, 0, 'the running service is left alone');
  } finally { await chmod(data, 0o700).catch(() => {}); await rm(fixture.home, { recursive: true, force: true }); }
});
