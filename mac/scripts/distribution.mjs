#!/usr/bin/env node
import * as fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import https from 'node:https';
import { createSecureContext } from 'node:tls';
import { lookup } from 'node:dns/promises';
import { fileURLToPath } from 'node:url';
import { createHash, randomUUID, X509Certificate } from 'node:crypto';
import { spawn, execFileSync } from 'node:child_process';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const root = fileURLToPath(new URL('../../', import.meta.url));
const action = process.argv[2];
const options = process.argv.slice(3);
function option(name, fallback) { const i = options.indexOf(name); return i < 0 ? fallback : options[i + 1]; }
const profile = path.resolve(option('--profile', os.homedir()));
const application = path.join(profile, 'Library/Application Support/SlideSCI-Mac');
const state = path.join(profile, 'Library/Application Support/SlideSCI-Mac-Installer');
const recordFile = path.join(state, 'installation.json');
const pidFile = path.join(state, 'service.json');
const launchAgentLabel = 'com.slidesci.mac.service';
const launchAgentFile = path.join(profile, 'Library/LaunchAgents', `${launchAgentLabel}.plist`);
const launchctlTool = process.env.SLIDESCI_LAUNCHCTL || '/bin/launchctl';
const launchctlAvailable = () => profile === path.resolve(os.homedir()) || Boolean(process.env.SLIDESCI_LAUNCHCTL);
const launchDomain = () => `gui/${process.getuid()}`;
const xmlEscape = value => String(value).replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
function launchAgentPlist() {
  const args = [
    path.join(application, 'runtime/bin/node'),
    path.join(application, 'mac/server.mjs'),
    '--cert', path.join(state, 'tls/localhost.crt'),
    '--key', path.join(state, 'tls/localhost.key')
  ];
  const logFile = path.join(state, 'service.log');
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">',
    '<plist version="1.0">',
    '<dict>',
    `  <key>Label</key><string>${launchAgentLabel}</string>`,
    '  <key>ProgramArguments</key>',
    '  <array>',
    ...args.map(arg => `    <string>${xmlEscape(arg)}</string>`),
    '  </array>',
    `  <key>WorkingDirectory</key><string>${xmlEscape(application)}</string>`,
    '  <key>RunAtLoad</key><true/>',
    '  <key>KeepAlive</key><dict><key>SuccessfulExit</key><false/></dict>',
    `  <key>StandardOutPath</key><string>${xmlEscape(logFile)}</string>`,
    `  <key>StandardErrorPath</key><string>${xmlEscape(logFile)}</string>`,
    '  <key>EnvironmentVariables</key>',
    '  <dict>',
    '    <key>NODE_ENV</key><string>production</string>',
    '    <key>SLIDESCI_SUPERVISOR</key><string>launchd</string>',
    `    <key>SLIDESCI_LOG_FILE</key><string>${xmlEscape(logFile)}</string>`,
    '  </dict>',
    '</dict>',
    '</plist>',
    ''
  ].join('\n');
}
async function writeLaunchAgent() {
  await atomic(launchAgentFile, launchAgentPlist(), 0o644);
}
// A loaded job is not necessarily a running one: after a clean exit (a taken port at login, or the
// process ended in Activity Monitor) launchd keeps the job loaded without restarting it.
function launchAgentState() {
  if (!launchctlAvailable()) return { loaded: false, running: false };
  try {
    const output = execFileSync(launchctlTool, ['print', `${launchDomain()}/${launchAgentLabel}`], { encoding: 'utf8', stdio: ['ignore', 'pipe', 'ignore'] });
    return { loaded: true, running: /^\s*state = running\s*$/m.test(output) || /^\s*pid = \d+\s*$/m.test(output) };
  } catch { return { loaded: false, running: false }; }
}
const launchAgentLoaded = () => launchAgentState().loaded;
async function waitForListening(certificate) {
  const deadline = Date.now() + 15000;
  while (Date.now() < deadline) {
    try {
      const response = await fetch('http://127.0.0.1:3000/api/session', { signal: AbortSignal.timeout(1000) });
      const session = await response.json();
      const tlsSession = await httpsSession(certificate);
      if (response.ok && session.acceptance === false && tlsSession.token === session.token) return session;
    } catch { /* retry until startup deadline */ }
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error('服务启动超时，请检查 service.log。');
}
function bootoutLaunchAgent() {
  if (!launchctlAvailable()) return false;
  try { execFileSync(launchctlTool, ['bootout', `${launchDomain()}/${launchAgentLabel}`], { stdio: 'ignore' }); return true; }
  catch {
    try { execFileSync(launchctlTool, ['bootout', launchDomain(), launchAgentFile], { stdio: 'ignore' }); return true; }
    catch { return false; }
  }
}
const stamp = () => new Date().toISOString().replace(/[:.]/g, '-');
const hash = data => createHash('sha256').update(data).digest('hex');
const exists = async file => fs.access(file).then(() => true, () => false);
const json = async file => JSON.parse(await fs.readFile(file, 'utf8'));
function run(file, args, cwd = root) { return execFileSync(file, args, { cwd, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 8 * 1024 * 1024 }); }
async function atomic(file, data, mode = 0o600) {
  await fs.mkdir(path.dirname(file), { recursive: true, mode: 0o700 });
  const temporary = `${file}.${randomUUID()}.tmp`;
  try { await fs.writeFile(temporary, data, { mode, flag: 'wx' }); await fs.rename(temporary, file); }
  finally { await fs.rm(temporary, { force: true }); }
}
function parse(text) {
  const errors = [];
  const document = new DOMParser({ onError: (level, message) => { if (level !== 'warning') errors.push(message); } }).parseFromString(text, 'application/xml');
  if (errors.length) throw new Error('加载项注册 XML 无效，停止修改。');
  return document;
}
const serialize = node => new XMLSerializer().serializeToString(node);
async function files(directory, prefix = '') {
  const result = [];
  for (const item of await fs.readdir(directory, { withFileTypes: true })) {
    const relative = path.join(prefix, item.name), file = path.join(directory, item.name);
    if (item.isSymbolicLink()) throw new Error(`分发源不允许符号链接：${relative}`);
    if (item.isDirectory()) result.push(...await files(file, relative));
    else if (item.isFile()) result.push(relative);
  }
  return result.sort();
}
function processSignature(pid) {
  if (!Number.isInteger(pid) || pid <= 1) return null;
  try { return run('/bin/ps', ['-p', String(pid), '-o', 'lstart=', '-o', 'command=']).trim() || null; }
  catch { return null; }
}
function ownsService(signature) {
  return typeof signature === 'string' && signature.includes(path.join(application, 'runtime/bin/node'))
    && signature.includes(path.join(application, 'mac/server.mjs'));
}
// The signature holds the start time and command line, so a mismatch proves the recorded
// service is gone (reboot, logout, crash or PID reuse). Such a record is dropped without
// signalling the process that now owns the PID.
async function serviceRecord() {
  if (!await exists(pidFile)) return null;
  let record = null;
  try { record = await json(pidFile); } catch { /* An unreadable record cannot identify a running service. */ }
  const current = record ? processSignature(Number(record.pid)) : null;
  if (record?.signature && current === record.signature && ownsService(current)) return record;
  await fs.rm(pidFile, { force: true });
  console.log('已清除过期的服务记录（电脑重启、服务退出或进程号已被其他程序使用）；未结束任何进程。');
  return null;
}
// Stop.command also removes the login item, or launchd would start the service again at the next
// login; Start.command writes it back. Upgrades and certificate renewal keep it (keepAutostart).
async function stop({ keepAutostart = false } = {}) {
  const hadAgent = launchAgentLoaded();
  if (hadAgent && !bootoutLaunchAgent()) {
    throw new Error('无法停止 SlideSCI 后台自启服务；未修改安装目录或加载项注册，升级与卸载已中止。请稍后重试或重启电脑。');
  }
  const hadAutostart = !keepAutostart && await exists(launchAgentFile);
  if (hadAutostart) await fs.rm(launchAgentFile, { force: true });
  const record = await serviceRecord();
  if (!record) {
    if (hadAgent) console.log(hadAutostart ? '已停止 SlideSCI 后台自启服务，并取消登录自动运行；需要时运行 Start.command。' : '已停止 SlideSCI 后台自启服务。');
    else console.log(hadAutostart ? 'SlideSCI 本机服务未在运行；已取消登录自动运行。' : 'SlideSCI 本机服务未在运行。');
    return;
  }
  process.kill(-record.pid, 'SIGTERM');
  const deadline = Date.now() + 10000;
  while (Date.now() < deadline) {
    try { process.kill(record.pid, 0); }
    catch { await fs.rm(pidFile, { force: true }); console.log('SlideSCI 本机服务已停止。'); return; }
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error('服务尚未退出；保留记录，不强制结束。');
}
// The service writes into each host's own cache; macOS asks before Terminal touches another app's
// data. Probing here makes that prompt appear in this window and names a refusal plainly.
const accessHelp = names => `macOS 拒绝“终端”访问 ${names.join('、')} 的数据，SlideSCI 的插入、格式与导出等功能将无法使用。出现“访问其他 App 的数据”提示时请选择允许；如果之前拒绝过，请在“系统设置 → 隐私与安全性 → 完全磁盘访问权限”中打开“终端”，然后重新运行此命令。`;
async function deniedHosts() {
  const denied = [];
  for (const [name, id] of [['PowerPoint', 'com.microsoft.Powerpoint'], ['WPS', 'com.kingsoft.wpsoffice.mac']]) {
    const container = path.join(profile, 'Library/Containers', id, 'Data');
    if (!await exists(container)) continue;
    const directory = path.join(container, 'tmp/SlideSCI-Mac'), probe = path.join(directory, `.slidesci-access-${process.pid}`);
    try { await fs.mkdir(directory, { recursive: true, mode: 0o700 }); await fs.writeFile(probe, ''); await fs.rm(probe, { force: true }); }
    catch (error) { if (['EPERM', 'EACCES'].includes(error.code)) denied.push(name); }
  }
  return denied;
}
function permissionMessage(error) {
  if (!['EPERM', 'EACCES'].includes(error?.code)) return error?.message;
  const where = String(error.path || error.message);
  const names = /com\.microsoft\.Powerpoint/i.test(where) ? ['PowerPoint'] : /kingsoft/i.test(where) ? ['WPS'] : null;
  return names ? accessHelp(names) : error.message;
}
async function reserve(port, host) {
  const server = net.createServer(socket => socket.destroy());
  await new Promise((resolve, reject) => server.once('error', reject).listen(port, host, resolve));
  return () => new Promise(resolve => server.close(resolve));
}
function httpsSession(certificate) {
  return new Promise((resolve, reject) => {
    const request = https.get('https://localhost:3443/api/session', { ca: certificate, agent: false, timeout: 1000 }, response => {
      let text = '';
      response.setEncoding('utf8');
      response.on('data', chunk => { text += chunk; if (text.length > 65536) request.destroy(new Error('HTTPS 会话响应过大。')); });
      response.on('end', () => {
        try { if (response.statusCode !== 200) throw new Error('HTTPS 会话尚未就绪。'); resolve(JSON.parse(text)); }
        catch (error) { reject(error); }
      });
      response.on('error', reject);
    });
    request.on('timeout', () => request.destroy(new Error('HTTPS 会话超时。')));
    request.on('error', reject);
  });
}
async function start({ launchAgent = launchctlAvailable(), trust = true } = {}) {
  if (!await exists(recordFile) || !await exists(path.join(application, 'BUILD-INFO.json'))) throw new Error('请先运行 Install.command。');
  const agent = launchAgentState();
  if (agent.loaded && !agent.running) {
    // Unload the exited job so it can be bootstrapped again below, after the port check.
    if (!bootoutLaunchAgent()) throw new Error('无法重置已退出的 SlideSCI 后台自启服务，请稍后重试或重启电脑。');
    console.log('SlideSCI 后台自启服务已退出，正在重新启动。');
  }
  let running = Boolean(await serviceRecord()) || agent.running;
  const { days } = await certificateStatus();
  if (days <= 0) {
    console.log('PowerPoint 使用的本机证书已过期，正在生成新证书。');
    await replaceCertificate(running);
    running = false;
    trust = false; // Renewal already attempted trust for the new certificate.
  } else if (days <= CERTIFICATE_RENEW_DAYS) {
    console.log(`提醒：PowerPoint 使用的本机证书将在 ${Math.ceil(days)} 天后到期，到期后 PowerPoint 无法加载 SlideSCI。`);
    console.log('可运行安装目录中的 RenewCertificate.command 提前生成新证书。');
  }
  const trusted = trust ? await trustCertificate() : certificateTrusted(path.join(tlsDirectory(), 'localhost.crt'));
  if (running) { console.log('SlideSCI 本机服务已经运行。'); return; }
  const certificate = await fs.readFile(path.join(state, 'tls/localhost.crt'));
  if (!process.env.SLIDESCI_LAUNCHCTL) {
    const reservations = [];
    try {
      const addresses = [...new Set((await lookup('localhost', { all: true })).map(value => value.address))];
      if (!addresses.length || addresses.some(address => !['127.0.0.1', '::1'].includes(address))) throw new Error('localhost 必须仅解析到回环地址。');
      reservations.push(await reserve(3000, '127.0.0.1'));
      for (const address of addresses) reservations.push(await reserve(3443, address));
    }
    catch (error) { for (const close of reservations) await close(); throw new Error(`本机端口 3000 或 3443 已被占用，未启动其他服务：${error.code || error.message}`); }
    for (const close of reservations) await close();
  }
  const logFile = path.join(state, 'service.log');
  // Keep one previous log so a long-lived installation never grows the log without bound.
  const logSize = await fs.stat(logFile).then(item => item.size, () => 0);
  if (logSize > 5 * 1024 * 1024) await fs.rename(logFile, logFile + '.1');
  if (launchAgent) {
    await writeLaunchAgent();
    execFileSync(launchctlTool, ['bootstrap', launchDomain(), launchAgentFile], { stdio: ['ignore', 'pipe', 'pipe'] });
    if (!process.env.SLIDESCI_LAUNCHCTL) await waitForListening(certificate);
    console.log('SlideSCI 后台服务已启动并设为登录自动运行：WPS http://127.0.0.1:3000；PowerPoint https://localhost:3443。');
    if (!trusted) console.log('本机证书尚未被信任：PowerPoint 暂时无法加载 SlideSCI，WPS 可以正常使用。');
    return;
  }
  const log = await fs.open(logFile, 'a', 0o600);
  const child = spawn(path.join(application, 'runtime/bin/node'), [path.join(application, 'mac/server.mjs'), '--cert', path.join(state, 'tls/localhost.crt'), '--key', path.join(state, 'tls/localhost.key')], {
    cwd: application, detached: true, stdio: ['ignore', log.fd, log.fd],
    env: { ...process.env, NODE_OPTIONS: '', NODE_ENV: 'production', SLIDESCI_ACCEPTANCE: '', SLIDESCI_NATIVE_DEBUG: '', SLIDESCI_HTTP_PORT: '', SLIDESCI_HTTPS_PORT: '', SLIDESCI_SUPERVISOR: '', SLIDESCI_LOG_FILE: logFile }
  });
  let spawnError;
  child.on('error', error => { spawnError = error; });
  child.unref(); await log.close();
  try {
    const deadline = Date.now() + 15000;
    while (Date.now() < deadline) {
      if (spawnError) throw spawnError;
      if (child.exitCode !== null) throw new Error('服务启动失败，请检查安装器状态目录中的 service.log。');
      try {
        const response = await fetch('http://127.0.0.1:3000/api/session', { signal: AbortSignal.timeout(1000) });
        const session = await response.json();
        const tlsSession = await httpsSession(certificate);
        const owners = [3000, 3443].map(port => run('/usr/sbin/lsof', ['-nP', `-iTCP:${port}`, '-sTCP:LISTEN', '-Fp']).trim().split('\n'));
        if (response.ok && session.acceptance === false && tlsSession.token === session.token && owners.every(values => values.includes(`p${child.pid}`))) {
          const signature = run('/bin/ps', ['-p', String(child.pid), '-o', 'lstart=', '-o', 'command=']).trim();
          await atomic(pidFile, JSON.stringify({ pid: child.pid, signature, root: application }));
          console.log('SlideSCI 已启动：WPS http://127.0.0.1:3000；PowerPoint https://localhost:3443。');
          if (!trusted) console.log('本机证书尚未被信任：PowerPoint 暂时无法加载 SlideSCI，WPS 可以正常使用。');
          return;
        }
      } catch { /* The process is checked again until the startup deadline. */ }
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    throw new Error('服务启动超时，请检查 service.log。');
  } catch (error) { try { process.kill(-child.pid, 'SIGTERM'); } catch {} throw error; }
}
// macOS rejects TLS server certificates valid for more than 825 days.
const CERTIFICATE_DAYS = 825;
const CERTIFICATE_RENEW_DAYS = 30;
const tlsDirectory = () => path.join(state, 'tls');
async function certificate(directory = tlsDirectory()) {
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const cert = path.join(directory, 'localhost.crt'), key = path.join(directory, 'localhost.key');
  if (await exists(cert) && await exists(key)) return false;
  if (await exists(cert) || await exists(key)) throw new Error('本机 TLS 文件不完整，请先恢复成对的证书与私钥。');
  const config = path.join(directory, 'openssl.cnf');
  await atomic(config, '[req]\ndistinguished_name=dn\nx509_extensions=extensions\nprompt=no\n[dn]\nCN=localhost\n[extensions]\nsubjectAltName=DNS:localhost,IP:127.0.0.1,IP:::1\nbasicConstraints=critical,CA:FALSE\nkeyUsage=critical,digitalSignature,keyEncipherment\nextendedKeyUsage=serverAuth\n');
  run('/usr/bin/openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-days', String(CERTIFICATE_DAYS), '-config', config, '-keyout', key, '-out', cert]);
  await fs.chmod(key, 0o600); await fs.chmod(cert, 0o644);
  return true;
}
async function certificateStatus() {
  const file = path.join(tlsDirectory(), 'localhost.crt');
  if (!await exists(file) || !await exists(path.join(tlsDirectory(), 'localhost.key'))) throw new Error('缺少 PowerPoint 使用的本机证书，请运行 RenewCertificate.command 重新生成。');
  const expires = Date.parse(new X509Certificate(await fs.readFile(file)).validTo);
  if (!Number.isFinite(expires)) throw new Error('无法读取本机证书的有效期，请运行 RenewCertificate.command 重新生成。');
  return { expires, days: (expires - Date.now()) / 86400000 };
}
// The retired pair is kept beside the new one; the old trust setting is removed, the new one asked for.
async function replaceCertificate(running) {
  const directory = tlsDirectory();
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const stage = await fs.mkdtemp(path.join(directory, '.renew-'));
  const retired = path.join(directory, `retired-${stamp()}`);
  const names = ['localhost.crt', 'localhost.key'];
  const backedUp = [], promoted = [];
  try {
    await certificate(stage);
    createSecureContext({ cert: await fs.readFile(path.join(stage, names[0])), key: await fs.readFile(path.join(stage, names[1])) });
    await fs.mkdir(retired, { mode: 0o700 });
    for (const name of names) {
      const original = path.join(directory, name);
      if (await exists(original)) {
        await fs.copyFile(original, path.join(retired, name), fs.constants.COPYFILE_EXCL);
        backedUp.push(name);
      }
      await fs.rename(path.join(stage, name), original);
      promoted.push(name);
    }
    // The running server keeps its old TLS context until the replacement pair is ready.
    if (running) await stop({ keepAutostart: true });
  } catch (error) {
    const recovery = [];
    for (const name of promoted.reverse()) {
      const original = path.join(directory, name);
      try {
        if (backedUp.includes(name)) await fs.copyFile(path.join(retired, name), original);
        else await fs.rm(original, { force: true });
      } catch (restoreError) { recovery.push(lastLine(restoreError)); }
    }
    if (recovery.length) throw new Error(`${lastLine(error)}；恢复旧证书失败：${recovery.join('；')}`, { cause: error });
    throw error;
  } finally { await fs.rm(stage, { recursive: true, force: true }); }
  const { expires } = await certificateStatus();
  console.log(`已生成新的本机证书，有效期至 ${new Date(expires).toLocaleDateString('zh-CN')}；旧证书移至 ${retired}。`);
  await untrustCertificate(path.join(retired, 'localhost.crt'));
  if (await trustCertificate()) console.log('请保存工作并重开 PowerPoint。');
}
async function renewCertificate() {
  if (!await exists(recordFile)) throw new Error('请先运行 Install.command。');
  const running = Boolean(await serviceRecord() || launchAgentLoaded());
  await replaceCertificate(running);
  if (running) await start({ trust: false });
  else console.log('需要使用时运行 Start.command。');
}
// PowerPoint loads the pane only over HTTPS that macOS trusts. With the user's consent the installer
// lets macOS show its own authorization prompt and trusts this one localhost leaf certificate for SSL
// only; it never creates a certificate authority. --trust-certificate yes|no answers the question
// for managed installs; an isolated --profile never touches the signed-in user's trust settings.
const securityTool = process.env.SLIDESCI_SECURITY_TOOL || '/usr/bin/security';
const trustRecord = () => path.join(tlsDirectory(), 'trust.json');
const loginKeychain = () => path.join(profile, 'Library/Keychains/login.keychain-db');
const trustAvailable = () => profile === path.resolve(os.homedir()) || Boolean(process.env.SLIDESCI_SECURITY_TOOL);
const fingerprint = pem => new X509Certificate(pem).fingerprint.replaceAll(':', '');
function certificateTrusted(file) {
  try { execFileSync(securityTool, ['verify-cert', '-c', file, '-p', 'ssl', '-n', 'localhost', '-L', '-q'], { stdio: 'ignore' }); return true; }
  catch { return false; }
}
function runningHosts() {
  if (process.env.SLIDESCI_RUNNING_HOSTS !== undefined) {
    return process.env.SLIDESCI_RUNNING_HOSTS.split(',').map(item => item.trim()).filter(Boolean);
  }
  if (profile !== path.resolve(os.homedir())) return [];
  try {
    const comms = run('/bin/ps', ['-axo', 'comm=']);
    const hosts = [];
    if (/Microsoft PowerPoint/i.test(comms)) hosts.push('PowerPoint');
    if (/wpsoffice/i.test(comms)) hosts.push('WPS');
    return hosts;
  } catch {
    return [];
  }
}
async function consent(question, optionName = '--trust-certificate') {
  const choice = option(optionName, null);
  if (choice === 'yes') return true;
  if (choice === 'no' || !process.stdin.isTTY) return false;
  const { createInterface } = await import('node:readline/promises');
  const prompt = createInterface({ input: process.stdin, output: process.stdout });
  try { return ['', 'y', 'yes', '是'].includes((await prompt.question(question + ' [Y/n] ')).trim().toLowerCase()); }
  finally { prompt.close(); }
}
// Ask once per run: a renewal that restarts the service must not repeat a question just declined.
let trustDeclined = false;
function lastLine(error) { return String(error?.stderr || error?.message || error).trim().split('\n').at(-1); }
async function trustCertificate() {
  const file = path.join(tlsDirectory(), 'localhost.crt');
  if (certificateTrusted(file)) return true;
  const manual = () => console.log(`PowerPoint 需要信任本机证书才能加载 SlideSCI（WPS 不受影响）。可再次运行 Start.command 按提示信任，或用“钥匙串访问”导入 ${file} 并将其 SSL 设为始终信任。`);
  if (!trustAvailable() || trustDeclined) { manual(); return false; }
  if (!await consent('PowerPoint 只通过受 macOS 信任的 HTTPS 加载 SlideSCI。是否现在信任本机生成的 localhost 证书？它只用于本机 localhost 连接，私钥不离开这台 Mac；macOS 会弹出授权窗口，需要输入登录密码。')) {
    trustDeclined = true; manual(); return false;
  }
  console.log('请在 macOS 弹出的授权窗口中输入登录密码……');
  try { execFileSync(securityTool, ['add-trusted-cert', '-r', 'trustRoot', '-p', 'ssl', '-s', 'localhost', '-k', loginKeychain(), file], { stdio: ['ignore', 'pipe', 'pipe'] }); }
  catch (error) { trustDeclined = true; console.log('未完成证书信任：' + lastLine(error)); manual(); return false; }
  if (!certificateTrusted(file)) { console.log('macOS 仍未信任该证书。'); manual(); return false; }
  await atomic(trustRecord(), JSON.stringify({ fingerprint: fingerprint(await fs.readFile(file)), keychain: loginKeychain(), trustedAt: new Date().toISOString() }, null, 2));
  console.log('已信任 PowerPoint 使用的本机证书。');
  return true;
}
// Removes only what trustCertificate added: the recorded certificate's keychain copy and its trust setting.
async function untrustCertificate(file) {
  if (!trustAvailable() || !await exists(trustRecord())) return;
  let record;
  try { record = await json(trustRecord()); } catch { return; }
  const pem = await fs.readFile(file).catch(() => null);
  if (!record?.fingerprint || (pem && fingerprint(pem) !== record.fingerprint)) return;
  console.log('正在移除安装器添加的证书信任，macOS 可能弹出授权窗口……');
  try {
    execFileSync(securityTool, ['delete-certificate', '-Z', record.fingerprint, '-t', record.keychain || loginKeychain()], { stdio: ['ignore', 'pipe', 'pipe'] });
    await fs.rm(trustRecord(), { force: true });
    console.log('已移除旧证书及其信任设置。');
  } catch (error) {
    console.log(`未能移除旧证书信任：${lastLine(error)}。可在“钥匙串访问”中删除 SHA-1 指纹为 ${record.fingerprint} 的 localhost 证书。`);
  }
}
async function registrations() {
  const result = [];
  const power = path.join(profile, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents');
  if (await exists(power)) {
    const target = path.join(power, 'wef/SlideSCI-Mac.xml');
    const prior = await exists(target) ? await fs.readFile(target, 'utf8') : null;
    const value = await fs.readFile(path.join(root, 'mac/office/manifest.xml'), 'utf8');
    result.push({ host: 'powerpoint', target, prior, before: prior, value, installedHash: hash(value) });
  }
  const wps = path.join(profile, 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons');
  if (await exists(wps)) {
    const target = path.join(wps, 'publish.xml');
    const before = await exists(target) ? await fs.readFile(target, 'utf8') : null;
    const document = parse(before ?? '<jsplugins/>'), parent = document.documentElement;
    if (parent.tagName !== 'jsplugins') throw new Error('WPS 注册根节点未知，未修改。');
    const entries = Array.from(parent.childNodes).filter(node => node.nodeType === 1 && node.getAttribute('name') === 'slidesci-mac-dev');
    if (entries.length > 1) throw new Error('存在重复的 SlideSCI WPS 注册，停止安装。');
    const prior = entries[0] ? serialize(entries[0]) : null;
    const entry = document.createElement('jspluginonline');
    for (const [name, value] of Object.entries({ name: 'slidesci-mac-dev', type: 'wpp', url: 'http://127.0.0.1:3000/wps/', debug: '', enable: 'enable_dev', install: 'null' })) entry.setAttribute(name, value);
    if (entries[0]) parent.replaceChild(entry, entries[0]); else parent.appendChild(entry);
    result.push({ host: 'wps', target, prior, before, value: serialize(document), installedEntry: serialize(entry) });
  }
  return result;
}
async function restoreRegistration(item, checkOnly = false) {
  if (!await exists(item.target)) return;
  const current = await fs.readFile(item.target, 'utf8');
  if (item.host === 'powerpoint') {
    if (hash(current) !== item.installedHash) throw new Error('PowerPoint 注册已被另行修改，保留当前文件及备份。');
    if (checkOnly) return;
    if (item.prior === null) await fs.rm(item.target); else await atomic(item.target, item.prior);
  } else {
    const document = parse(current), parent = document.documentElement;
    const entries = Array.from(parent.childNodes).filter(node => node.nodeType === 1 && node.getAttribute('name') === 'slidesci-mac-dev');
    if (entries.length !== 1 || serialize(entries[0]) !== item.installedEntry) throw new Error('WPS 的 SlideSCI 注册已被另行修改，保留当前文件及备份。');
    if (checkOnly) return;
    if (item.prior === null) parent.removeChild(entries[0]);
    else parent.replaceChild(document.importNode(parse(item.prior).documentElement, true), entries[0]);
    await atomic(item.target, serialize(document));
  }
}
async function clearWefCache(activeHosts) {
  const wefCache = path.join(profile, 'Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef');
  if (option('--clear-wef-cache', null) === 'no' || !await exists(wefCache)) return;
  if (activeHosts.includes('PowerPoint')) {
    console.log('PowerPoint 正在运行，未自动清理 Wef 加载项缓存。若重开 PowerPoint 后在“开发人员加载项”中看不到新版 SlideSCI Mac，请先 Cmd+Q 完全退出 PowerPoint 并执行：');
    console.log('rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef');
    return;
  }
  // The folder also caches other add-ins (they are fetched again), so it is cleared only on consent.
  if (!await consent('PowerPoint 可能仍缓存旧版 SlideSCI 功能区。是否清理 PowerPoint 加载项缓存（Wef 文件夹，其他加载项的缓存会在下次打开时重新下载）？', '--clear-wef-cache')) {
    console.log('未清理 Wef 加载项缓存。若重开 PowerPoint 后仍显示旧版或找不到插件，请完全退出 PowerPoint 后执行：');
    console.log('rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef');
    return;
  }
  try {
    await fs.rm(wefCache, { recursive: true, force: true });
    console.log('已清理 PowerPoint 加载项缓存（Wef），重开 PowerPoint 后将加载新版功能区。');
  } catch (error) {
    console.log(`未能自动清理 PowerPoint 加载项缓存（${lastLine(error)}）。若重开后仍显示旧版或找不到插件，请完全退出 PowerPoint 后执行：`);
    console.log('rm -rf ~/Library/Containers/com.microsoft.Powerpoint/Data/Library/Caches/Microsoft/Office/16.0/Wef');
  }
}
// Each upgrade keeps exactly one previous program folder for rollback; older ones are removed.
async function prunePreviousVersions(keep) {
  const directory = path.dirname(application), prefix = `${path.basename(application)}.previous-`;
  const removed = [];
  for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
    const file = path.join(directory, entry.name);
    if (!entry.isDirectory() || !entry.name.startsWith(prefix) || file === keep) continue;
    try { await fs.rm(file, { recursive: true, force: true }); removed.push(entry.name); }
    catch (error) { console.log(`未能删除更早的旧版本备份 ${file}：${lastLine(error)}`); }
  }
  if (removed.length) console.log(`已删除更早的旧版本备份：${removed.join('、')}`);
}
async function install() {
  if (process.platform !== 'darwin' || process.arch !== 'arm64') throw new Error('此安装包仅适用于 Apple Silicon Mac。Intel Mac 请使用源码开发流程；尚未实机验证。');
  if (path.resolve(root) === application) throw new Error('请从解压后的发行目录运行安装器。');
  if (!await exists(path.join(root, 'BUILD-INFO.json')) || !await exists(path.join(root, 'runtime/bin/node'))) throw new Error('请先构建并解压发行 ZIP，再运行 Install.command。');

  const incomingVersion = (await json(path.join(root, 'BUILD-INFO.json'))).version;
  const existingRecord = await exists(recordFile) ? await json(recordFile) : null;
  if (existingRecord) {
    if (existingRecord.application !== application) throw new Error('已有安装记录的路径不匹配，停止升级。');
    const oldVersion = existingRecord.version || '未知版本';
    if (!await consent(`检测到已安装的 SlideSCI ${oldVersion}，是否原地升级到 ${incomingVersion}？（将自动停止旧服务，保留素材、设置与已信任证书）`, '--upgrade')) {
      throw new Error('已取消原地升级；旧版本保持不变（非交互模式可传 --upgrade yes）。');
    }
    for (const item of existingRecord.registrations || []) await restoreRegistration(item, true);
  } else {
    if (await exists(application)) throw new Error('目标目录已有未登记内容，未覆盖。');
    if (await serviceRecord() || launchAgentLoaded()) throw new Error('SlideSCI 服务仍在运行，请先运行 Stop.command。');
  }

  const planned = await registrations();
  if (!planned.length) throw new Error('尚未找到 PowerPoint 或 WPS 的容器。请先安装并打开宿主一次。');
  const plannedNames = new Set(planned.map(item => item.host === 'powerpoint' ? 'PowerPoint' : 'WPS'));
  const activeHosts = runningHosts().filter(name => plannedNames.has(name));
  if (activeHosts.length) {
    console.log(`检测到 ${activeHosts.join('、')} 正在运行：安装/升级完成后请先保存工作并完全退出（Cmd+Q）再重开，新加载项才会生效。`);
  }
  const originalPriors = new Map((existingRecord?.registrations || []).map(item => [item.host, item.prior]));
  const powerPointManifestChanged = planned.some(item => item.host === 'powerpoint' && item.before !== null && item.before !== item.value);
  let previousService = null;
  let wroteLaunchAgent = false;
  const priorAgentFile = existingRecord ? await fs.lstat(launchAgentFile).catch(error => {
    if (error.code !== 'ENOENT') throw error;
    return null;
  }) : null;
  const priorLaunchAgent = priorAgentFile?.isFile() ? await fs.readFile(launchAgentFile) : null;

  await fs.mkdir(state, { recursive: true, mode: 0o700 });
  const stage = `${application}.install-${randomUUID()}`;
  await fs.mkdir(path.dirname(application), { recursive: true, mode: 0o700 });
  const installed = [];
  let moved = false;
  let previousApp = null;
  let previousRecordBackup = null;
  let preserveActiveInstall = false;
  try {
    if (existingRecord) {
      const running = await serviceRecord() ? 'detached' : launchAgentLoaded() ? 'launchAgent' : null;
      if (running) { await stop({ keepAutostart: true }); previousService = running; }
    }
    await fs.cp(root, stage, { recursive: true, errorOnExist: true, force: false });
    await certificate();
    if (existingRecord && await exists(application)) {
      previousApp = `${application}.previous-${stamp()}`;
      await fs.rename(application, previousApp);
    }
    await fs.rename(stage, application); moved = true;
    for (const { before, value, ...item } of planned) {
      const current = await exists(item.target) ? await fs.readFile(item.target, 'utf8') : null;
      if (current !== before) throw new Error('安装期间加载项注册已变化，停止覆盖并恢复本次写入。');
      if (before !== null) await atomic(path.join(state, `backups/${stamp()}-${item.host}.xml`), before);
      await atomic(item.target, value); installed.push(item);
    }
    if (existingRecord) {
      previousRecordBackup = path.join(state, `backups/${stamp()}-installation.json`);
      await atomic(previousRecordBackup, await fs.readFile(recordFile, 'utf8'));
    }
    const savedRegistrations = existingRecord
      ? installed.map(item => originalPriors.has(item.host) ? { ...item, prior: originalPriors.get(item.host) } : item)
      : installed;
    await atomic(recordFile, JSON.stringify({ application, version: incomingVersion, registrations: savedRegistrations, launchAgent: launchAgentFile }, null, 2));
    await writeLaunchAgent();
    wroteLaunchAgent = true;
    if (installed.some(item => item.host === 'powerpoint')) await trustCertificate();
    if (launchctlAvailable()) {
      try {
        await start();
      } catch (error) {
        if (previousService) {
          if (launchAgentLoaded()) {
            try { await stop({ keepAutostart: true }); }
            catch (stopError) {
              preserveActiveInstall = true;
              throw new Error(`${lastLine(error)}；新版服务未能停止，已保留当前安装：${lastLine(stopError)}`, { cause: error });
            }
          }
          throw error;
        }
        console.log('自动启动后台服务失败（稍后可手动运行 Start.command）：' + lastLine(error));
      }
    } else {
      console.log('已生成登录自启配置（隔离测试目录未向系统 launchd 注册）。');
    }
    if (powerPointManifestChanged) await clearWefCache(activeHosts);
    if (previousApp) await prunePreviousVersions(previousApp);
    console.log((existingRecord ? `SlideSCI 已升级到 ${incomingVersion}。注册宿主：` : 'SlideSCI 已安装。注册宿主：') + installed.map(item => item.host).join('、'));
    if (previousApp) console.log(`旧版本程序已备份至：${previousApp}`);
    if (installed.some(item => item.host === 'powerpoint')) {
      console.log('PowerPoint 启用位置：开始 → 加载项 → 更多加载项 → 开发人员加载项 → 双击 SlideSCI Mac。');
    }
    if (activeHosts.length) {
      console.log(`提醒：请保存工作并 Cmd+Q 完全退出 ${activeHosts.join('、')} 后重开，以加载最新插件。`);
    }
  } catch (error) {
    if (preserveActiveInstall) throw error;
    const rollbackErrors = [];
    for (const item of installed.reverse()) {
      try { await restoreRegistration(item); }
      catch (rollbackError) { rollbackErrors.push(`加载项注册（${item.host}）：${lastLine(rollbackError)}`); }
    }
    if (moved) {
      try { await fs.rename(application, `${application}.failed-${stamp()}`); }
      catch (rollbackError) { rollbackErrors.push(`移出失败的程序目录：${lastLine(rollbackError)}`); }
    }
    if (previousApp && await exists(previousApp) && !await exists(application)) {
      try { await fs.rename(previousApp, application); }
      catch (rollbackError) { rollbackErrors.push(`恢复旧版程序目录：${lastLine(rollbackError)}`); }
    }
    if (previousRecordBackup && await exists(previousRecordBackup)) {
      try { await atomic(recordFile, await fs.readFile(previousRecordBackup, 'utf8')); }
      catch (rollbackError) { rollbackErrors.push(`恢复安装记录：${lastLine(rollbackError)}`); }
    }
    if (wroteLaunchAgent && !existingRecord) {
      // A failed fresh install must not leave a login item pointing at the removed program.
      if (launchAgentLoaded() && !bootoutLaunchAgent()) rollbackErrors.push('停止新装的登录自启服务失败');
      try { await fs.rm(launchAgentFile, { force: true }); }
      catch (rollbackError) { rollbackErrors.push(`删除登录自启配置：${lastLine(rollbackError)}`); }
    }
    if (wroteLaunchAgent && existingRecord) {
      try {
        if (priorLaunchAgent) await atomic(launchAgentFile, priorLaunchAgent, 0o644);
        else await fs.rm(launchAgentFile, { force: true });
      } catch (rollbackError) { rollbackErrors.push(`恢复登录自启配置：${lastLine(rollbackError)}`); }
    }
    if (previousService) {
      try { await start({ launchAgent: previousService === 'launchAgent', trust: false }); console.log('已恢复升级前正在运行的 SlideSCI 后台服务。'); }
      catch (rollbackError) { rollbackErrors.push(`重新启动后台服务：${lastLine(rollbackError)}`); }
    }
    if (rollbackErrors.length) throw new Error(`${error.message}；回滚未完全完成：${rollbackErrors.join('；')}`);
    throw error;
  } finally { await fs.rm(stage, { recursive: true, force: true }); }
}
async function uninstall() {
  if (!await exists(recordFile)) { console.log('未找到此安装器的安装记录。'); return; }
  const record = await json(recordFile);
  if (record.application !== application) throw new Error('安装路径不匹配，停止卸载。');
  for (const item of record.registrations) await restoreRegistration(item, true);
  await stop();
  if (await exists(launchAgentFile)) await fs.rm(launchAgentFile, { force: true });
  for (const item of [...record.registrations].reverse()) await restoreRegistration(item);
  if (await exists(application)) await fs.rename(application, `${application}.uninstalled-${stamp()}`);
  await fs.rename(recordFile, path.join(state, `uninstalled-${stamp()}.json`));
  await untrustCertificate(path.join(tlsDirectory(), 'localhost.crt'));
  console.log('已移除本安装的加载项注册、自启配置并恢复原有记录。程序移入相邻 .uninstalled 备份目录。');
  console.log('素材、设置和安装器备份均保留；证书文件保留在安装器状态目录，重装时会再次询问是否信任。');
}
async function prunePackages(directory) {
  const excluded = new Set(['.bin', 'test', 'tests', '__tests__', 'fixtures', 'examples', 'example', 'benchmark', 'benchmarks', 'coverage']);
  for (const item of await fs.readdir(directory, { withFileTypes: true })) {
    const file = path.join(directory, item.name);
    if (excluded.has(item.name)) await fs.rm(file, { recursive: true, force: true });
    else if (item.isDirectory()) await prunePackages(file);
  }
}
async function build() {
  if (process.platform !== 'darwin' || process.arch !== 'arm64') throw new Error('发行构建目前要求 Apple Silicon Mac。');
  const version = (await json(path.join(root, 'mac/package.json'))).version;
  const nodeVersion = '24.22.3';
  const nodeSha = 'ff6e17d4874b76d272212b908cc020d6e815fba5a6c7eaaa06add725840307fd';
  const filename = `node-v${nodeVersion}-darwin-arm64.tar.gz`;
  const output = path.resolve(option('--output', path.join(root, 'artifacts/distribution')));
  const cache = path.join(output, 'cache');
  await fs.mkdir(cache, { recursive: true });
  const archive = path.join(cache, filename);
  if (!await exists(archive)) {
    const response = await fetch(`https://nodejs.org/dist/v${nodeVersion}/${filename}`, { signal: AbortSignal.timeout(600000) });
    if (!response.ok) throw new Error('Node.js 官方下载失败：' + response.status);
    await atomic(archive, Buffer.from(await response.arrayBuffer()));
  }
  if (hash(await fs.readFile(archive)) !== nodeSha) throw new Error('Node.js 归档 SHA-256 不匹配，停止构建。');
  const name = `SlideSCI-Mac-${version}-darwin-arm64`;
  const temporary = await fs.mkdtemp(path.join(output, '.build-'));
  const payload = path.join(temporary, name);
  await fs.mkdir(path.join(payload, 'mac'), { recursive: true });
  try {
    for (const directory of ['src', 'lib', 'native', 'assets', 'wps', 'office']) await fs.cp(path.join(root, 'mac', directory), path.join(payload, 'mac', directory), { recursive: true });
    await fs.mkdir(path.join(payload, 'mac/scripts'));
    await fs.copyFile(path.join(root, 'mac/scripts/distribution.mjs'), path.join(payload, 'mac/scripts/distribution.mjs'));
    for (const file of ['package.json', 'package-lock.json', 'server.mjs', 'taskpane.html', 'taskpane.css', 'DISTRIBUTION.md']) await fs.copyFile(path.join(root, 'mac', file), path.join(payload, 'mac', file));
    await fs.cp(path.join(root, 'SlideSCI/latex-converter'), path.join(payload, 'SlideSCI/latex-converter'), { recursive: true });
    await fs.copyFile(path.join(root, 'LICENSE'), path.join(payload, 'LICENSE'));
    await fs.copyFile(path.join(root, 'mac/DISTRIBUTION.md'), path.join(payload, 'INSTALL.md'));
    const sourceFiles = (await files(payload)).filter(file => !file.endsWith('INSTALL.md'));
    const sourceManifest = Object.fromEntries(await Promise.all(sourceFiles.map(async file => [file.split(path.sep).join('/'), hash(await fs.readFile(path.join(payload, file)))])));
    await atomic(path.join(payload, 'SOURCE-MANIFEST.json'), JSON.stringify(sourceManifest, null, 2), 0o644);
    let baseCommit = null;
    try { baseCommit = run('git', ['rev-parse', 'HEAD']).trim(); } catch {}
    await atomic(path.join(payload, 'BUILD-INFO.json'), JSON.stringify({ version, builtAt: new Date().toISOString(), platform: 'darwin', architecture: 'arm64', node: nodeVersion, nodeArchiveSha256: nodeSha, gitBaseCommit: baseCommit, source: 'Packaged source snapshot; the Git base commit does not identify this snapshot.', sourceManifestSha256: hash(JSON.stringify(sourceManifest, null, 2)), signed: false, notarized: false, coreAcceptance: 'See release acceptance notes; build success does not certify host behavior.' }, null, 2), 0o644);
    const extracted = path.join(temporary, 'node');
    await fs.mkdir(extracted);
    run('/usr/bin/tar', ['-xzf', archive, '-C', extracted, '--strip-components=1', `node-v${nodeVersion}-darwin-arm64/bin/node`, `node-v${nodeVersion}-darwin-arm64/LICENSE`]);
    await fs.mkdir(path.join(payload, 'runtime/bin'), { recursive: true });
    await fs.copyFile(path.join(extracted, 'bin/node'), path.join(payload, 'runtime/bin/node'));
    await fs.copyFile(path.join(extracted, 'LICENSE'), path.join(payload, 'runtime/NODE-LICENSE'));
    await fs.chmod(path.join(payload, 'runtime/bin/node'), 0o755);
    console.log('正在安装锁定的生产依赖。');
    run('npm', ['ci', '--omit=dev', '--ignore-scripts', '--registry=https://registry.npmjs.org', '--cache', path.join(cache, 'npm')], path.join(payload, 'mac'));
    await prunePackages(path.join(payload, 'mac/node_modules'));
    await fs.mkdir(path.join(payload, 'mac/bin'));
    console.log('正在编译随包分发的原生辅助程序。');
    run('/usr/bin/xcrun', ['swiftc', '-O', '-swift-version', '5', '-target', 'arm64-apple-macos14.0', '-debug-prefix-map', `${root}=/SlideSCI`, path.join(payload, 'mac/native/SlideSCI.swift'), '-o', path.join(payload, 'mac/bin/slidesci-native'), '-framework', 'AppKit', '-framework', 'Security', '-framework', 'PDFKit', '-framework', 'CoreImage', '-framework', 'ApplicationServices', '-framework', 'AVFoundation', '-framework', 'UniformTypeIdentifiers']);
    await fs.chmod(path.join(payload, 'mac/bin/slidesci-native'), 0o755);
    for (const [name, command] of Object.entries({ Install: 'install', Start: 'start', Stop: 'stop', Uninstall: 'uninstall', RenewCertificate: 'renew-certificate' })) {
      const text = '#!/bin/bash\nset -euo pipefail\nif [[ "$(uname -s)" != Darwin || "$(uname -m)" != arm64 ]]; then echo "此安装包仅适用于 Apple Silicon Mac。"; exit 1; fi\nTASK_ROOT="$(cd "$(dirname "$0")" && pwd)"\nexec "$TASK_ROOT/runtime/bin/node" "$TASK_ROOT/mac/scripts/distribution.mjs" ' + command + ' "$@"\n';
      await atomic(path.join(payload, `${name}.command`), text, 0o755);
    }
    for (const file of await files(payload)) {
      const data = await fs.readFile(path.join(payload, file));
      if (data.includes(Buffer.from(os.homedir() + path.sep))) throw new Error(`分发文件包含个人路径：${file}`);
      if (/-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----\r?\n[A-Za-z0-9+/]{32}/.test(data.toString('latin1'))) throw new Error(`分发文件包含私钥：${file}`);
    }
    const zip = path.join(output, name + '.zip');
    if (await exists(zip)) throw new Error('输出 ZIP 已存在，未覆盖：' + path.basename(zip));
    run('/usr/bin/ditto', ['-c', '-k', '--keepParent', '--norsrc', payload, zip]);
    await atomic(zip + '.sha256', hash(await fs.readFile(zip)) + '  ' + path.basename(zip) + '\n', 0o644);
    console.log('已生成：' + zip);
  } finally { await fs.rm(temporary, { recursive: true, force: true }); }
}
try {
  if (action === 'build') await build();
  else if (action === 'install') await install();
  else if (action === 'start') await start();
  else if (action === 'stop') await stop();
  else if (action === 'uninstall') await uninstall();
  else if (action === 'renew-certificate') await renewCertificate();
  else throw new Error('使用：distribution.mjs build|install|start|stop|uninstall|renew-certificate [--profile 隔离测试目录] [--output 构建目录] [--trust-certificate yes|no] [--upgrade yes|no] [--clear-wef-cache yes|no]');
} catch (error) { console.error(permissionMessage(error)); process.exitCode = 1; }
