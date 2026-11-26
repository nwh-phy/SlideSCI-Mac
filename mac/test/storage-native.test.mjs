import test from 'node:test';
import assert from 'node:assert/strict';
import { execFile } from 'node:child_process';
import { readFile, writeFile, rm, symlink } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import path from 'node:path';
import { mkdtemp, mkdir } from 'node:fs/promises';
import os from 'node:os';
import { promisify } from 'node:util';
import { createFileServices } from '../lib/storage.mjs';
import { nativeHandlers } from '../lib/native.mjs';

// Never write into the real PowerPoint/WPS containers: run against a throwaway home whose
// containers mirror the layout the services expect.
const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-home-'));
process.env.HOME = home;
test.after(() => rm(home, { recursive: true, force: true }));

for (const host of ['wps', 'powerpoint']) test(`${host} content files retain native ownership through read and release`, async () => {
  const env = { native: { materialize: input => nativeHandlers.materialize(input) } };
  Object.assign(env, await createFileServices(env));
  let file;
  try {
    const original = Buffer.from('SlideSCI native content ownership regression');
    file = await env.files.materialize({ host, data: original.toString('base64'), name: 'ownership-regression.bin', mime: 'application/octet-stream' });
    assert.deepEqual(await readFile(file.path), original);
    assert.equal((await env.files.read(file.path)).data, original.toString('base64'));
    await nativeHandlers.execute({ host, action: 'releaseFiles', paths: [file.path] }, env);
    await assert.rejects(readFile(file.path), { code: 'ENOENT' });
  } finally {
    if (file) await rm(file.path, { force: true });
    await rm(env.workDir, { recursive: true, force: true });
  }
});

test('native cleanup refuses a file it did not issue even inside its cache', async () => {
  const env = { native: { materialize: input => nativeHandlers.materialize(input) } };
  Object.assign(env, await createFileServices(env));
  let issued, unissued;
  try {
    issued = await env.files.materialize({ host: 'wps', data: Buffer.from('issued').toString('base64'), name: 'ownership-regression.bin' });
    unissued = path.join(path.dirname(issued.path), randomUUID() + '-unissued-regression.bin');
    await writeFile(unissued, 'must remain intact', { flag: 'wx' });
    await assert.rejects(nativeHandlers.execute({ host: 'wps', action: 'releaseFiles', paths: [unissued] }, env), /Only issued files/);
    assert.equal(await readFile(unissued, 'utf8'), 'must remain intact');
    await nativeHandlers.execute({ host: 'wps', action: 'releaseFiles', paths: [issued.path] }, env);
  } finally {
    if (issued) await rm(issued.path, { force: true });
    if (unissued) await rm(unissued, { force: true });
    await rm(env.workDir, { recursive: true, force: true });
  }
});

test('a fresh service releases only durably reserved WPS PPTX files', async () => {
  const exec = promisify(execFile);
  const script = `
    import { nativeHandlers } from ${JSON.stringify(new URL('../lib/native.mjs', import.meta.url).href)};
    const input = JSON.parse(process.argv[1]);
    console.log(JSON.stringify(await nativeHandlers.execute({ ...input, host: 'wps' })));
  `;
  const runNative = async (input) => {
    const { stdout } = await exec(process.execPath, ['--input-type=module', '--eval', script, JSON.stringify(input)]);
    return JSON.parse(stdout);
  };
  const release = filename => runNative({ action: 'releaseFiles', paths: [filename] });
  const { path: reserved } = await runNative({ action: 'temporaryPath', extension: 'pptx' });
  const unissued = path.join(path.dirname(reserved), `${randomUUID()}.pptx`);
  await writeFile(unissued, 'unissued', { flag: 'wx' });

  await symlink(unissued, reserved);
  await assert.rejects(release(reserved), /Only issued files/);
  await rm(reserved);
  await writeFile(reserved, 'issued', { flag: 'wx' });
  await assert.rejects(runNative({ action: 'readFile', path: reserved }), /Choose this file/);

  await release(reserved);
  await assert.rejects(readFile(reserved), { code: 'ENOENT' });
  await assert.rejects(release(unissued), /Only issued files/);
  assert.equal(await readFile(unissued, 'utf8'), 'unissued');

  await writeFile(reserved, 'replacement', { flag: 'wx' });
  await assert.rejects(release(reserved), /Only issued files/);
  assert.equal(await readFile(reserved, 'utf8'), 'replacement');
});
