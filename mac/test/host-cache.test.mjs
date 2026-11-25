import test from 'node:test';
import assert from 'node:assert/strict';
import { access, mkdir, mkdtemp, readFile, realpath, rm, utimes, writeFile } from 'node:fs/promises';
import { randomUUID } from 'node:crypto';
import os from 'node:os';
import path from 'node:path';
import { createFileServices } from '../lib/storage.mjs';
import { hostCacheSweepsSettled, nativeHandlers } from '../lib/native.mjs';

const exists = file => access(file).then(() => true, () => false);
const ids = { powerpoint: 'com.microsoft.Powerpoint', wps: 'com.kingsoft.wpsoffice.mac' };
const twoDaysAgo = () => new Date(Date.now() - 2 * 86400000);

// Each test runs against its own throwaway home whose containers mirror the real layout, so the
// real PowerPoint/WPS data folders are never touched.
async function freshHome(t) {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-host-cache-'));
  for (const id of Object.values(ids)) await mkdir(path.join(home, 'Library/Containers', id, 'Data'), { recursive: true });
  const previous = process.env.HOME;
  process.env.HOME = home;
  t.after(async () => { process.env.HOME = previous; await rm(home, { recursive: true, force: true }); });
  return { home, cache: host => path.join(home, 'Library/Containers', ids[host], 'Data/tmp/SlideSCI-Mac') };
}
async function stale(file, text = 'stale') {
  await writeFile(file, text);
  await utimes(file, twoDaysAgo(), twoDaysAgo());
}

test('startup never touches host containers; a host cache is registered when first used', async t => {
  const { home, cache } = await freshHome(t);
  const env = { native: { materialize: input => nativeHandlers.materialize(input) } };
  Object.assign(env, await createFileServices(env));
  t.after(() => rm(env.workDir, { recursive: true, force: true }));
  for (const id of Object.values(ids)) {
    assert.equal(await exists(path.join(home, 'Library/Containers', id, 'Data/tmp')), false,
      `starting the file services leaves ${id} untouched (no macOS app-data prompt at login)`);
  }

  const file = await env.files.materialize({ host: 'powerpoint', data: Buffer.from('picture').toString('base64'), name: 'picture.png', mime: 'image/png' });
  assert.equal((await env.files.read(file.path)).data, Buffer.from('picture').toString('base64'));
  assert.equal(await exists(cache('wps')), false, 'a PowerPoint write does not touch the WPS container');

  // A file the host itself wrote into its cache (an AppleScript export, a WPS deck) is readable
  // without a picker grant, and registering that cache still leaves the other container alone.
  // The temporary home sits behind /var -> /private/var, like a home folder behind a symlink, and
  // the native layer asks with the resolved path.
  const hostWritten = path.join(cache('powerpoint'), randomUUID() + '.pdf');
  await writeFile(hostWritten, 'host export');
  const resolved = await realpath(hostWritten);
  assert.equal(await env.files.requireFile(resolved), resolved);
  assert.equal(await env.files.requireFile(hostWritten), resolved);
  assert.equal(await exists(cache('wps')), false);

  const outside = path.join(home, 'Documents/private.txt');
  await mkdir(path.dirname(outside), { recursive: true });
  await writeFile(outside, 'not granted');
  await assert.rejects(env.files.requireFile(outside), /请先通过文件选择器选择该文件/);
});

test('a host cache is swept on the first write into it, then at most once a day', async t => {
  const { home, cache } = await freshHome(t);
  await mkdir(cache('powerpoint'), { recursive: true });
  await mkdir(cache('wps'), { recursive: true });
  const oldExport = path.join(cache('powerpoint'), randomUUID() + '-slide-1.png');
  const oldDeck = path.join(cache('powerpoint'), randomUUID() + '.pptx');
  const oldFolder = path.join(cache('powerpoint'), randomUUID());
  const oldVideo = path.join(cache('powerpoint'), randomUUID() + '-clip.mp4');
  const fresh = path.join(cache('powerpoint'), randomUUID() + '-fresh.png');
  for (const file of [oldExport, oldDeck, oldVideo]) await stale(file);
  await mkdir(oldFolder);
  await stale(path.join(oldFolder, 'page.pdf'));
  await utimes(oldFolder, twoDaysAgo(), twoDaysAgo());
  await writeFile(fresh, 'fresh');
  // A WPS deck reserved by an earlier service run keeps a marker beside the user's data.
  const wpsName = randomUUID() + '.pptx';
  const wpsDeck = path.join(cache('wps'), wpsName);
  const marker = path.join(home, 'Library/Application Support/SlideSCI Mac/wps-pptx-reservations', wpsName);
  await mkdir(path.dirname(marker), { recursive: true });
  await stale(wpsDeck);
  await stale(marker, '');

  const written = await nativeHandlers.materialize({ host: 'powerpoint', data: Buffer.from('new').toString('base64'), name: 'new.png' });
  await hostCacheSweepsSettled();
  assert.equal(await exists(oldExport), false, 'a day-old export is removed');
  assert.equal(await exists(oldDeck), false, 'a day-old intermediate deck is removed');
  assert.equal(await exists(oldFolder), false, 'a day-old scratch folder is removed');
  assert.equal(await exists(oldVideo), true, 'media is kept because a host may link it');
  assert.equal(await readFile(fresh, 'utf8'), 'fresh', 'files from the last day are kept');
  assert.equal(await readFile(written.path, 'utf8'), 'new', 'the file being written is kept');
  assert.equal(await exists(wpsDeck), true, 'a PowerPoint write never sweeps the WPS cache');

  const laterExport = path.join(cache('powerpoint'), randomUUID() + '-slide-2.png');
  await stale(laterExport);
  await nativeHandlers.materialize({ host: 'powerpoint', data: Buffer.from('again').toString('base64'), name: 'again.png' });
  await hostCacheSweepsSettled();
  assert.equal(await exists(laterExport), true, 'a cache is swept at most once a day');

  await nativeHandlers.execute({ host: 'wps', action: 'temporaryPath', extension: 'png' });
  await hostCacheSweepsSettled();
  assert.equal(await exists(wpsDeck), false, 'the first WPS write sweeps the WPS cache');
  assert.equal(await exists(marker), false, 'and drops the reservation of a swept WPS deck');
});
