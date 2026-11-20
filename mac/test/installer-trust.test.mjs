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
if (args[0] === 'print') process.exit(fs.existsSync(loaded) ? 0 : 113);
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
if (args[0] === 'print') process.exit(fs.existsSync(loaded) ? 0 : 113);
if (args[0] === 'bootstrap') { fs.writeFileSync(loaded, args.at(-1)); process.exit(0); }
if (args[0] === 'bootout') { if (!fs.existsSync(loaded)) process.exit(3); fs.rmSync(loaded); process.exit(0); }
process.exit(2);
`;
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

    const upgradeOut = runPkg(v2Pkg, 'install', ['--upgrade', 'yes', '--trust-certificate', 'no']);
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
