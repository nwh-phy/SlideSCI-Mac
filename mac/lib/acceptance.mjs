import { randomUUID } from 'node:crypto';
export function acceptanceHandlers() {
  const jobs = new Map();
  const title = value => String(value || '').replace(/\.pptx?$/i, '').toLowerCase();
  return {
    enqueue(input) {
      if (!['wps', 'powerpoint'].includes(input.host) || !/^slidesci/i.test(input.targetTitle || '')) throw new Error('Only explicit SlideSCI test documents are allowed.');
      const id = randomUUID();
      jobs.set(id, { ...input, id, state: 'queued' });
      return { id };
    },
    next(input) {
      const job = [...jobs.values()].find(item => item.state === 'queued' && item.host === input.host && title(item.targetTitle) === title(input.title));
      if (!job) return {};
      job.state = 'running';
      return { job };
    },
    result(input) {
      const job = jobs.get(input.id);
      if (!job) throw new Error('Unknown acceptance job');
      Object.assign(job, { state: input.ok ? 'completed' : 'failed', result: input.result, error: input.error });
      return { recorded: true };
    },
    status(input) {
      const job = jobs.get(input.id);
      if (!job) throw new Error('Unknown acceptance job');
      return job;
    }
  };
}
