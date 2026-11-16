import test from 'node:test';
import assert from 'node:assert/strict';
import { startBackgroundTracking, stopBackgroundTracking } from '../src/background.js';

function fakeServices({ failFirst = false } = {}) {
  const listeners = new Map();
  let calls = 0;
  const services = {
    started: [],
    stopped: [],
    subscribe(name, callback) {
      listeners.set(name, callback);
      return () => listeners.delete(name);
    },
    emit(name, payload) { listeners.get(name)?.(payload); },
    async startWatcher() {
      calls++;
      if (failFirst && calls === 1) throw new Error('无法连接 SlideSCI 本地服务');
      const id = 'lease-' + calls;
      services.started.push(id);
      return { id };
    },
    async stopWatcher(id) { services.stopped.push(id); },
    async reportError() {}
  };
  return services;
}
const adapter = { async context() { return { shapes: [] }; } };
const settle = () => new Promise(resolve => setTimeout(resolve, 10));

test('background tracking survives a service that is down at start and recovers on hello', async () => {
  const services = fakeServices({ failFirst: true });
  const statuses = [];
  const runtime = await startBackgroundTracking({ host: 'powerpoint', services, hostAdapter: adapter, onStatus: event => statuses.push(event) });
  try {
    assert.equal(runtime.stopped, false, 'a failed first lease no longer disables tracking for the session');
    assert.equal(runtime.leaseId, null);
    assert.match(statuses.at(-1).message, /自动粘贴监听暂未启动/);
    services.emit('hello', { instance: 'b', leases: [] });
    await settle();
    assert.equal(runtime.leaseId, 'lease-2');
    assert.match(statuses.at(-1).message, /已恢复/);
  } finally {
    await stopBackgroundTracking('powerpoint');
  }
  assert.deepEqual(services.stopped, ['lease-2']);
});

test('hello keeps a lease the service still holds and replaces one it lost', async () => {
  const services = fakeServices();
  const runtime = await startBackgroundTracking({ host: 'wps', services, hostAdapter: adapter });
  try {
    assert.equal(runtime.leaseId, 'lease-1');
    services.emit('hello', { instance: 'a', leases: ['lease-1'] });
    await settle();
    assert.deepEqual(services.started, ['lease-1'], 'a surviving lease is not re-acquired');
    services.emit('hello', { instance: 'restarted', leases: [] });
    await settle();
    assert.equal(runtime.leaseId, 'lease-2', 'a restarted service gets a new lease');
  } finally {
    await stopBackgroundTracking('wps');
  }
});
