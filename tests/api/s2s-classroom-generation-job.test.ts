import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { NextRequest } from 'next/server';

const mocks = vi.hoisted(() => ({ readJob: vi.fn() }));
vi.mock('@/lib/server/classroom-job-store', async (importOriginal) => ({
  ...(await importOriginal<typeof import('@/lib/server/classroom-job-store')>()),
  readClassroomGenerationJob: mocks.readJob,
}));
vi.mock('@/lib/logger', () => ({
  createLogger: () => ({
    error: vi.fn(),
    warn: vi.fn(),
    info: vi.fn(),
    debug: vi.fn(),
  }),
}));

const TOKEN = 'primary-service-token-that-is-at-least-32-bytes';
const JOB_ID = 's2s_380822546de379f6967f12dbe7cae042';

async function poll(jobId = JOB_ID, authorization = `Bearer ${TOKEN}`) {
  const { GET } = await import('@/app/api/s2s/classroom-generations/[jobId]/route');
  const request = new Request(`http://openmaic:3000/api/s2s/classroom-generations/${jobId}`, {
    headers: authorization ? { Authorization: authorization } : undefined,
  });
  return GET(request as unknown as NextRequest, {
    params: Promise.resolve({ jobId }),
  });
}

function job(status: 'queued' | 'running' | 'succeeded' | 'failed') {
  return {
    id: JOB_ID,
    status,
    step:
      status === 'queued'
        ? 'queued'
        : status === 'failed'
          ? 'failed'
          : status === 'succeeded'
            ? 'completed'
            : 'generating_scenes',
    progress: status === 'succeeded' ? 100 : 40,
    message: status === 'failed' ? 'Classroom generation failed' : `Job ${status}`,
    scenesGenerated: status === 'succeeded' ? 12 : 4,
    totalScenes: 12,
    result:
      status === 'succeeded'
        ? {
            classroomId: 'classroom-1',
            url: 'http://internal/classroom/1',
            scenesCount: 12,
          }
        : undefined,
    error: status === 'failed' ? 'provider leaked secret detail' : undefined,
  };
}

describe('GET /api/s2s/classroom-generations/{jobId}', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN', TOKEN);
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN_PREVIOUS', '');
    mocks.readJob.mockReset();
  });

  afterEach(() => vi.unstubAllEnvs());

  it('authenticates before validating or loading the job', async () => {
    const response = await poll('not valid!', '');

    expect(response.status).toBe(401);
    expect(mocks.readJob).not.toHaveBeenCalled();
  });

  it('returns 400 for a non-S2S id and 404 for a missing valid S2S job', async () => {
    const invalid = await poll('browser_job');
    expect(mocks.readJob).not.toHaveBeenCalled();
    mocks.readJob.mockResolvedValue(null);
    const missing = await poll('s2s_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa');

    expect(invalid.status).toBe(400);
    expect(missing.status).toBe(404);
  });

  it.each(['queued', 'running', 'succeeded', 'failed'] as const)(
    'returns a stable, relative and redacted %s representation',
    async (status) => {
      mocks.readJob.mockResolvedValue(job(status));

      const response = await poll();
      const body = await response.json();

      expect(response.status).toBe(200);
      expect(body.jobId).toBe(JOB_ID);
      expect(body.status).toBe(status);
      expect(body.pollPath).toBe(`/api/s2s/classroom-generations/${JOB_ID}`);
      expect(body.done).toBe(status === 'succeeded' || status === 'failed');
      expect(JSON.stringify(body)).not.toContain('http://internal');
      expect(JSON.stringify(body)).not.toContain('provider leaked secret detail');
      expect(body).not.toHaveProperty('message');
      if (status === 'succeeded') {
        expect(body.result).toEqual({
          classroomId: 'classroom-1',
          scenesCount: 12,
        });
      }
      if (status === 'failed') {
        expect(body.errorCode).toBe('GENERATION_FAILED');
      }
    },
  );
});
