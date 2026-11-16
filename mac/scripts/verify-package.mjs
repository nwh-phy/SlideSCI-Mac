#!/usr/bin/env node
// Verifies a built installer ZIP in an isolated profile; it never touches the signed-in user's
// PowerPoint/WPS registrations, keychain or trust settings.
//
//   node mac/scripts/verify-package.mjs <SlideSCI-Mac-x.y.z-darwin-arm64.zip> [--report report.json]
//
// The service start check needs ports 3000 and 3443; when another service owns them it is
// reported as skipped, never stopped.
import * as fs from 'node:fs/promises';
import net from 'node:net';
import https from 'node:https';
import os from 'node:os';
import path from 'node:path';
import { createHash, X509Certificate } from 'node:crypto';
import { execFileSync } from 'node:child_process';

const [zip] = process.argv.slice(2).filter(value => !value.startsWith('--'));
const reportIndex = process.argv.indexOf('--report');
if (!zip) throw new Error('用法：verify-package.mjs <安装包.zip> [--report 报告.json]');
const checks = [];
const exists = file => fs.access(file).then(() => true, () => false);
const sha256 = data => createHash('sha256').update(data).digest('hex');
async function check(name, run) {
  const started = Date.now();
  try {
    const detail = await run();
    if (detail?.skipped) { checks.push({ name, ok: null, skipped: detail.skipped, ms: Date.now() - started }); console.log(`- 跳过 ${name}：${detail.skipped}`); return; }
    checks.push({ name, ok: true, ms: Date.now() - started, ...(detail ? { detail } : {}) });
    console.log(`✓ ${name}`);
  } catch (error) {
    checks.push({ name, ok: false, ms: Date.now() - started, error: String(error.stderr || error.message || error).trim().slice(0, 2000) });
    console.log(`✗ ${name}：${String(error.message || error).split('\n')[0]}`);
  }
}
const sandbox = await fs.mkdtemp(path.join(os.tmpdir(), 'slidesci verify '));
const profile = path.join(sandbox, 'profile with spaces');
const unpacked = path.join(sandbox, 'unpacked');
await fs.mkdir(unpacked, { recursive: true });
execFileSync('/usr/bin/ditto', ['-x', '-k', zip, unpacked]);
const [payloadName] = await fs.readdir(unpacked);
const payload = path.join(unpacked, payloadName);
const node = path.join(payload, 'runtime/bin/node');
const run = (file, args, options = {}) => execFileSync(file, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], ...options });
const command = (name, ...args) => run('/bin/bash', [path.join(payload, `${name}.command`), '--profile', profile, ...args]);
const application = path.join(profile, 'Library/Application Support/SlideSCI-Mac');
const installer = path.join(profile, 'Library/Application Support/SlideSCI-Mac-Installer');
const wef = path.join(profile, 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef');
const jsaddons = path.join(profile, 'Library/Containers/com.kingsoft.wpsoffice.mac/Data/.kingsoft/wps/jsaddons');
const otherPlugin = '<jsplugins><jspluginonline name="other-addin" type="wpp" url="http://127.0.0.1:9/other/"/></jsplugins>';
let build;

await check('package layout, commands and build record', async () => {
  build = JSON.parse(await fs.readFile(path.join(payload, 'BUILD-INFO.json'), 'utf8'));
  const version = JSON.parse(await fs.readFile(path.join(payload, 'mac/package.json'), 'utf8')).version;
  if (build.version !== version || !payloadName.includes(version)) throw new Error(`版本不一致：${build.version} / ${version} / ${payloadName}`);
  for (const name of ['Install', 'Start', 'Stop', 'Uninstall', 'RenewCertificate']) {
    const stat = await fs.stat(path.join(payload, `${name}.command`));
    if (!(stat.mode & 0o111)) throw new Error(`${name}.command 不可执行`);
  }
  return { version, gitBaseCommit: build.gitBaseCommit, node: build.node };
});
await check('source manifest matches every packaged source file', async () => {
  const manifest = JSON.parse(await fs.readFile(path.join(payload, 'SOURCE-MANIFEST.json'), 'utf8'));
  for (const [file, digest] of Object.entries(manifest)) if (sha256(await fs.readFile(path.join(payload, file))) !== digest) throw new Error('哈希不一致：' + file);
  return { files: Object.keys(manifest).length };
});
await check('bundled Node runtime and production dependencies', async () => {
  const version = run(node, ['--version']).trim();
  if (version !== 'v' + build.node) throw new Error('内置 Node 版本为 ' + version);
  const script = "const sharp=(await import('sharp')).default;const png=await sharp({create:{width:2,height:2,channels:4,background:'#ff000080'}}).png().toBuffer();const raw=await sharp(png).raw().toBuffer();process.stdout.write(JSON.stringify([...raw.subarray(0,4)]));";
  const pixel = JSON.parse(run(node, ['--input-type=module', '-e', script], { cwd: path.join(payload, 'mac') }));
  if (pixel.join() !== '255,0,0,128') throw new Error('图像编解码结果异常：' + pixel);
  return { node: version };
});
await check('precompiled native helper runs without a compiler', async () => {
  const helper = path.join(payload, 'mac/bin/slidesci-native');
  const fonts = JSON.parse(run(helper, [], { input: JSON.stringify({ action: 'fonts' }), stdio: ['pipe', 'pipe', 'pipe'] }));
  if (!Array.isArray(fonts) || !fonts.length) throw new Error('字体列表为空');
  return { fonts: fonts.length };
});
await check('packaged native equation conversion', async () => {
  const script = "const { contentHandlers } = await import('./lib/content.mjs');const result = await contentHandlers.render({ kind: 'latex', source: '\\\\frac{a}{b}+x^2', width: 400, fontSize: 24, slideWidth: 960, slideHeight: 540 }, {});process.stdout.write(JSON.stringify({ pptx: result.fragment?.pptx?.length || 0, warnings: (result.warnings || []).length }));";
  const result = JSON.parse(run(node, ['--input-type=module', '-e', script], { cwd: path.join(payload, 'mac') }));
  if (!(result.pptx > 1000) || result.warnings) throw new Error('公式转换结果异常：' + JSON.stringify(result));
  return result;
});
await check('isolated install registers both hosts, keeps other add-ins and creates a local certificate', async () => {
  await fs.mkdir(wef, { recursive: true });
  await fs.mkdir(jsaddons, { recursive: true });
  await fs.writeFile(path.join(jsaddons, 'publish.xml'), otherPlugin);
  const output = command('Install', '--trust-certificate', 'no');
  const manifest = await fs.readFile(path.join(wef, 'SlideSCI-Mac.xml'), 'utf8');
  if (manifest !== await fs.readFile(path.join(payload, 'mac/office/manifest.xml'), 'utf8')) throw new Error('PowerPoint 注册内容不一致');
  const publish = await fs.readFile(path.join(jsaddons, 'publish.xml'), 'utf8');
  if (!publish.includes('other-addin') || !publish.includes('slidesci-mac-dev')) throw new Error('WPS 注册未保留其他加载项：' + publish);
  const certificate = new X509Certificate(await fs.readFile(path.join(installer, 'tls/localhost.crt')));
  const days = Math.round((Date.parse(certificate.validTo) - Date.parse(certificate.validFrom)) / 86400000);
  if (days !== 825 || !/DNS:localhost/.test(certificate.subjectAltName) || !/IP Address:0:0:0:0:0:0:0:1|::1/.test(certificate.subjectAltName)) throw new Error(`证书异常：${days} 天，${certificate.subjectAltName}`);
  if (!/PowerPoint 需要信任本机证书/.test(output)) throw new Error('拒绝信任后没有给出说明');
  return { certificateDays: days };
});
await check('installed service starts, serves HTTP and HTTPS and stops', async () => {
  const free = async (port, host) => new Promise(resolve => { const probe = net.createServer().once('error', () => resolve(false)).listen(port, host, () => probe.close(() => resolve(true))); });
  if (!(await free(3000, '127.0.0.1')) || !(await free(3443, '::1')) || !(await free(3443, '127.0.0.1'))) return { skipped: '3000/3443 端口已被其他服务占用（未结束它）' };
  const started = command('Start', '--trust-certificate', 'no');
  try {
    const session = await (await fetch('http://127.0.0.1:3000/api/session')).json();
    const ca = await fs.readFile(path.join(installer, 'tls/localhost.crt'));
    const secure = await new Promise((resolve, reject) => https.get('https://localhost:3443/api/session', { ca, agent: false }, response => {
      let text = ''; response.on('data', chunk => { text += chunk; }); response.on('end', () => resolve(JSON.parse(text)));
    }).on('error', reject));
    if (secure.token !== session.token || session.version !== build.version || session.acceptance !== false) throw new Error('会话异常');
    return { version: session.version, started: started.trim().split('\n').at(-1) };
  } finally { command('Stop'); }
});
await check('a stale service record after reboot is cleared without signalling anyone', async () => {
  await fs.writeFile(path.join(installer, 'service.json'), JSON.stringify({ pid: 1, signature: 'stale', root: application }));
  const output = command('Stop');
  if (await exists(path.join(installer, 'service.json')) || !/过期的服务记录/.test(output)) throw new Error(output);
});
await check('certificate renewal keeps the retired pair', async () => {
  const before = await fs.readFile(path.join(installer, 'tls/localhost.crt'));
  command('RenewCertificate', '--trust-certificate', 'no');
  const after = await fs.readFile(path.join(installer, 'tls/localhost.crt'));
  const retired = (await fs.readdir(path.join(installer, 'tls'))).filter(name => name.startsWith('retired-'));
  if (before.equals(after) || retired.length !== 1) throw new Error('证书未换新或旧证书未保留');
});
await check('uninstall restores registrations and keeps user data', async () => {
  const data = path.join(profile, 'Library/Application Support/SlideSCI Mac');
  await fs.mkdir(data, { recursive: true });
  await fs.writeFile(path.join(data, 'keep.txt'), 'user data');
  command('Uninstall');
  if (await exists(path.join(wef, 'SlideSCI-Mac.xml'))) throw new Error('PowerPoint 注册未移除');
  const publish = await fs.readFile(path.join(jsaddons, 'publish.xml'), 'utf8');
  if (publish.includes('slidesci-mac-dev') || !publish.includes('other-addin')) throw new Error('WPS 注册未正确恢复：' + publish);
  if (await exists(application) || !(await fs.readdir(path.dirname(application))).some(name => name.startsWith('SlideSCI-Mac.uninstalled-'))) throw new Error('程序目录未移入备份');
  if (await fs.readFile(path.join(data, 'keep.txt'), 'utf8') !== 'user data') throw new Error('用户数据被改动');
});

const report = { zip: path.basename(zip), sha256: sha256(await fs.readFile(zip)), version: build?.version, startedAt: new Date().toISOString(), checks,
  passed: checks.filter(item => item.ok === true).length, failed: checks.filter(item => item.ok === false).length, skipped: checks.filter(item => item.ok === null).length };
if (reportIndex > 0) await fs.writeFile(process.argv[reportIndex + 1], JSON.stringify(report, null, 2) + '\n');
await fs.rm(sandbox, { recursive: true, force: true });
console.log(`通过 ${report.passed}，失败 ${report.failed}，跳过 ${report.skipped}。`);
process.exitCode = report.failed ? 1 : 0;
