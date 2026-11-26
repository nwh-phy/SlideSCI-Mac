import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, readFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { nativeHandlers } from '../lib/native.mjs';

// Real-host regression: a 150 dpi BMP slide export (about 5 MB) failed with "Maximum call stack size exceeded".
test('host cache copies accept multi-megabyte base64 and still reject malformed data', async () => {
  const home = await mkdtemp(path.join(os.tmpdir(), 'slidesci-decode-'));
  const previous = process.env.HOME;
  process.env.HOME = home;
  try {
    const bytes = Buffer.alloc(12 * 1024 * 1024, 0x5a);
    const file = await nativeHandlers.materialize({ data: bytes.toString('base64'), mime: 'image/bmp', name: 'slide-1.bmp', host: 'powerpoint' });
    assert.ok((await readFile(file.path)).equals(bytes));
    for (const bad of ['abc', 'ab=c', 'a===', 'ab$d']) await assert.rejects(nativeHandlers.materialize({ data: bad, host: 'powerpoint' }), /Invalid base64/);
  } finally {
    process.env.HOME = previous;
    await rm(home, { recursive: true, force: true });
  }
});
