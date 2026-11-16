#!/usr/bin/env node
import * as fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import https from 'node:https';
import { lookup } from 'node:dns/promises';
import { fileURLToPath } from 'node:url';
import { createHash, randomUUID } from 'node:crypto';
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
async function stop() {
  if (!await exists(pidFile)) return;
  const record = await json(pidFile);
  let signature;
  try { signature = run('/bin/ps', ['-p', String(record.pid), '-o', 'lstart=', '-o', 'command=']).trim(); }
  catch { await fs.rm(pidFile); return; }
  if (!record.signature || signature !== record.signature || !signature.includes(path.join(application, 'runtime/bin/node')) || !signature.includes(path.join(application, 'mac/server.mjs'))) {
    throw new Error('服务 PID 已不属于本安装，拒绝结束该进程。');
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
async function start() {
  if (!await exists(recordFile) || !await exists(path.join(application, 'BUILD-INFO.json'))) throw new Error('请先运行 Install.command。');
  if (await exists(pidFile)) {
    const record = await json(pidFile);
    try {
      if (run('/bin/ps', ['-p', String(record.pid), '-o', 'lstart=', '-o', 'command=']).trim() === record.signature) {
        console.log('SlideSCI 本机服务已经运行。'); return;
      }
    } catch { /* Only a recorded dead process is discarded. */ }
    throw new Error('服务记录过期或不匹配，请先运行 Stop.command 检查。');
  }
  const certificate = await fs.readFile(path.join(state, 'tls/localhost.crt'));
  const reservations = [];
  try {
    const addresses = [...new Set((await lookup('localhost', { all: true })).map(value => value.address))];
    if (!addresses.length || addresses.some(address => !['127.0.0.1', '::1'].includes(address))) throw new Error('localhost 必须仅解析到回环地址。');
    reservations.push(await reserve(3000, '127.0.0.1'));
    for (const address of addresses) reservations.push(await reserve(3443, address));
  }
  catch (error) { for (const close of reservations) await close(); throw new Error(`本机端口 3000 或 3443 已被占用，未启动其他服务：${error.code}`); }
  for (const close of reservations) await close();
  const log = await fs.open(path.join(state, 'service.log'), 'a', 0o600);
  const child = spawn(path.join(application, 'runtime/bin/node'), [path.join(application, 'mac/server.mjs'), '--cert', path.join(state, 'tls/localhost.crt'), '--key', path.join(state, 'tls/localhost.key')], {
    cwd: application, detached: true, stdio: ['ignore', log.fd, log.fd],
    env: { ...process.env, NODE_OPTIONS: '', NODE_ENV: 'production', SLIDESCI_ACCEPTANCE: '', SLIDESCI_NATIVE_DEBUG: '' }
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
          console.log('PowerPoint 需在本机信任安装时生成的 localhost 证书。'); return;
        }
      } catch { /* The process is checked again until the startup deadline. */ }
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    throw new Error('服务启动超时，请检查 service.log。');
  } catch (error) { try { process.kill(-child.pid, 'SIGTERM'); } catch {} throw error; }
}
async function certificate() {
  const directory = path.join(state, 'tls');
  await fs.mkdir(directory, { recursive: true, mode: 0o700 });
  const cert = path.join(directory, 'localhost.crt'), key = path.join(directory, 'localhost.key');
  if (await exists(cert) && await exists(key)) return;
  if (await exists(cert) || await exists(key)) throw new Error('本机 TLS 文件不完整，请先恢复成对的证书与私钥。');
  const config = path.join(directory, 'openssl.cnf');
  await atomic(config, '[req]\ndistinguished_name=dn\nx509_extensions=extensions\nprompt=no\n[dn]\nCN=localhost\n[extensions]\nsubjectAltName=DNS:localhost,IP:127.0.0.1\nbasicConstraints=critical,CA:FALSE\nkeyUsage=critical,digitalSignature,keyEncipherment\nextendedKeyUsage=serverAuth\n');
  run('/usr/bin/openssl', ['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-days', '365', '-config', config, '-keyout', key, '-out', cert]);
  await fs.chmod(key, 0o600); await fs.chmod(cert, 0o644);
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
async function install() {
  if (process.platform !== 'darwin' || process.arch !== 'arm64') throw new Error('此安装包仅适用于 Apple Silicon Mac。Intel Mac 请使用源码开发流程；尚未实机验证。');
  if (path.resolve(root) === application) throw new Error('请从解压后的发行目录运行安装器。');
  if (!await exists(path.join(root, 'BUILD-INFO.json')) || !await exists(path.join(root, 'runtime/bin/node'))) throw new Error('请先构建并解压发行 ZIP，再运行 Install.command。');
  if (await exists(recordFile)) throw new Error('已有本安装器版本，请先运行旧版本的 Uninstall.command；素材和设置会保留。');
  if (await exists(application)) throw new Error('目标目录已有未登记内容，未覆盖。');
  if (await exists(pidFile)) throw new Error('存在旧服务记录，请先运行 Stop.command。');
  const planned = await registrations();
  if (!planned.length) throw new Error('尚未找到 PowerPoint 或 WPS 的容器。请先安装并打开宿主一次。');
  await fs.mkdir(state, { recursive: true, mode: 0o700 });
  const stage = `${application}.install-${randomUUID()}`;
  await fs.mkdir(path.dirname(application), { recursive: true, mode: 0o700 });
  const installed = [];
  let moved = false;
  try {
    await fs.cp(root, stage, { recursive: true, errorOnExist: true, force: false });
    await certificate();
    await fs.rename(stage, application); moved = true;
    for (const { before, value, ...item } of planned) {
      const current = await exists(item.target) ? await fs.readFile(item.target, 'utf8') : null;
      if (current !== before) throw new Error('安装期间加载项注册已变化，停止覆盖并恢复本次写入。');
      if (before !== null) await atomic(path.join(state, `backups/${stamp()}-${item.host}.xml`), before);
      await atomic(item.target, value); installed.push(item);
    }
    await atomic(recordFile, JSON.stringify({ application, version: (await json(path.join(application, 'BUILD-INFO.json'))).version, registrations: installed }, null, 2));
    console.log('SlideSCI 已安装。注册宿主：' + installed.map(item => item.host).join('、'));
    console.log('下一步：运行安装目录中的 Start.command，保存工作并重开宿主。');
    console.log('PowerPoint 证书位于安装器状态目录 tls/localhost.crt；需自行导入并信任。安装器未修改钥匙串。');
  } catch (error) {
    for (const item of installed.reverse()) await restoreRegistration(item);
    if (moved) await fs.rename(application, `${application}.failed-${stamp()}`);
    throw error;
  } finally { await fs.rm(stage, { recursive: true, force: true }); }
}
async function uninstall() {
  if (!await exists(recordFile)) { console.log('未找到此安装器的安装记录。'); return; }
  const record = await json(recordFile);
  if (record.application !== application) throw new Error('安装路径不匹配，停止卸载。');
  for (const item of record.registrations) await restoreRegistration(item, true);
  await stop();
  for (const item of [...record.registrations].reverse()) await restoreRegistration(item);
  if (await exists(application)) await fs.rename(application, `${application}.uninstalled-${stamp()}`);
  await fs.rename(recordFile, path.join(state, `uninstalled-${stamp()}.json`));
  console.log('已移除本安装的加载项注册并恢复原有记录。程序移入相邻 .uninstalled 备份目录。');
  console.log('素材、设置、钥匙串密钥、本机证书和安装器备份均保留。');
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
    for (const [name, command] of Object.entries({ Install: 'install', Start: 'start', Stop: 'stop', Uninstall: 'uninstall' })) {
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
  else throw new Error('使用：distribution.mjs build|install|start|stop|uninstall [--profile 隔离测试目录] [--output 构建目录]');
} catch (error) { console.error(error.message); process.exitCode = 1; }
