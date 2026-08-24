import { promises as fs } from 'fs';
import os from 'os';
import path from 'path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { GenerateClassroomInput } from '@/lib/server/classroom-generation';

let jobsDir: string;

vi.mock('@/lib/server/classroom-storage', async () => {
  const fsModule = await import('fs');
  const pathModule = await import('path');
  return {
    get CLASSROOM_JOBS_DIR() {
      return jobsDir;
    },
    ensureClassroomJobsDir: () => fsModule.promises.mkdir(jobsDir, { recursive: true }),
    writeJsonFileAtomic: async (filePath: string, data: unknown) => {
      await fsModule.promises.mkdir(pathModule.dirname(filePath), {
        recursive: true,
      });
      const temp = `${filePath}.${process.pid}.${Date.now()}.tmp`;
      await fsModule.promises.writeFile(temp, JSON.stringify(data), 'utf8');
      await fsModule.promises.rename(temp, filePath);
    },
  };
});

const KEY = 'teacheros:123e4567-e89b-42d3-a456-426614174000';
const CHECKSUM = 'f1135ac0903e24b16a03c4a7e31742215d5af8f4ba428bfe5301fcce911267e1';
const INPUT: GenerateClassroomInput = {
  requirement: '八年级语文《桃花源记》第一课时',
  enableWebSearch: false,
  enableImageGeneration: false,
  enableVideoGeneration: false,
  enableTTS: false,
  agentMode: 'default',
};

async function createOrGet(checksum = CHECKSUM) {
  const { createOrGetClassroomGenerationJob } = await import('@/lib/server/classroom-job-store');
  return createOrGetClassroomGenerationJob(KEY, checksum, INPUT);
}

describe('S2S classroom job idempotency (single replica)', () => {
  beforeEach(async () => {
    vi.resetModules();
    jobsDir = await fs.mkdtemp(path.join(os.tmpdir(), 'openmaic-s2s-jobs-'));
  });

  afterEach(async () => {
    await fs.rm(jobsDir, { recursive: true, force: true });
  });

  it('atomically creates one deterministic job for concurrent same-key requests', async () => {
    const results = await Promise.all(Array.from({ length: 10 }, () => createOrGet()));

    expect(results.filter((result) => result.outcome === 'created')).toHaveLength(1);
    expect(results.filter((result) => result.outcome === 'replayed')).toHaveLength(9);
    const jobIds = results.map((result) => {
      if (result.outcome === 'conflict') {
        throw new Error('same checksum must not conflict');
      }
      return result.job.id;
    });
    expect(new Set(jobIds)).toEqual(new Set(['s2s_380822546de379f6967f12dbe7cae042']));
    const files = await fs.readdir(jobsDir);
    expect(files).toEqual(['s2s_380822546de379f6967f12dbe7cae042.json']);
  });

  it('replays the same persisted job after a module reload', async () => {
    const first = await createOrGet();
    vi.resetModules();
    const replay = await createOrGet();

    expect(first.outcome).toBe('created');
    expect(replay.outcome).toBe('replayed');
    if (first.outcome === 'conflict' || replay.outcome === 'conflict') {
      throw new Error('same checksum must not conflict');
    }
    expect(replay.job.id).toBe(first.job.id);
  });

  it('reports a conflict when the same key is reused for another request checksum', async () => {
    await createOrGet();
    const conflict = await createOrGet('0'.repeat(64));

    expect(conflict).toEqual({
      outcome: 'conflict',
      jobId: 's2s_380822546de379f6967f12dbe7cae042',
    });
  });

  it('stores only the idempotency key hash, never the raw key', async () => {
    const created = await createOrGet();
    if (created.outcome === 'conflict') {
      throw new Error('initial creation must not conflict');
    }
    const persisted = await fs.readFile(path.join(jobsDir, `${created.job.id}.json`), 'utf8');

    expect(persisted).not.toContain(KEY);
    expect(JSON.parse(persisted).idempotency).toEqual({
      keyHash: '380822546de379f6967f12dbe7cae0429cb5435fbd6c0265aef8b97b33f08e94',
      requestChecksum: CHECKSUM,
    });
  });
});
