import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, rm } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { aiHandlers, decodeCompletion } from '../lib/ai.mjs';

test('a streamed tool call without parameters may arrive with empty arguments', async () => {
  const frames = [
    { choices: [{ delta: { tool_calls: [{ index: 0, id: 'call_1', function: { name: 'get_presentation_context', arguments: '' } }] } }] },
    { choices: [{ delta: {}, finish_reason: 'tool_calls' }] }
  ].map(value => `data: ${JSON.stringify(value)}\n\n`).join('') + 'data: [DONE]\n\n';
  const response = new Response(frames, { headers: { 'content-type': 'text/event-stream' } });
  const message = await decodeCompletion(response, { events: [] });
  assert.equal(message.tool_calls[0].function.arguments, '{}');
});

test('one damaged history file does not hide the other conversations', async () => {
  const storageDir = await mkdtemp(path.join(os.tmpdir(), 'slidesci-ai-'));
  try {
    const history = path.join(storageDir, 'ai/history');
    await mkdir(history, { recursive: true });
    await writeFile(path.join(history, 'good.json'), JSON.stringify({ id: 'good', title: '保留', updatedAt: '2026-11-16T00:00:00.000Z', messages: [] }));
    await writeFile(path.join(history, 'broken.json'), '{"id":');
    const { conversations } = await aiHandlers.history({}, { storageDir });
    assert.deepEqual(conversations.map(item => item.id), ['good']);
  } finally { await rm(storageDir, { recursive: true, force: true }); }
});
