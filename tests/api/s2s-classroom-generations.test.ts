import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { NextRequest } from 'next/server';

const mocks = vi.hoisted(() => ({
  after: vi.fn((callback: () => unknown) => callback()),
  createOrGet: vi.fn(),
  run: vi.fn().mockResolvedValue(undefined),
}));

vi.mock('next/server', async (importOriginal) => ({
  ...(await importOriginal<typeof import('next/server')>()),
  after: mocks.after,
}));
vi.mock('@/lib/server/classroom-job-store', () => ({
  createOrGetClassroomGenerationJob: mocks.createOrGet,
}));
vi.mock('@/lib/server/classroom-job-runner', () => ({
  runClassroomGenerationJob: mocks.run,
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
const KEY = 'teacheros:123e4567-e89b-42d3-a456-426614174000';
const BODY = {
  requirement: '八年级语文《桃花源记》第一课时',
  enableWebSearch: false,
  enableImageGeneration: false,
  enableVideoGeneration: false,
  enableTTS: false,
  agentMode: 'default',
};
const JOB = {
  id: 's2s_job',
  status: 'queued',
  step: 'queued',
  message: 'Classroom generation job queued',
};

async function submit(options?: { authorization?: string; body?: unknown; rawBody?: string }) {
  const { POST } = await import('@/app/api/s2s/classroom-generations/route');
  const request = new Request('http://openmaic:3000/api/s2s/classroom-generations', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Idempotency-Key': KEY,
      ...(options?.authorization === undefined
        ? { Authorization: `Bearer ${TOKEN}` }
        : options.authorization
          ? { Authorization: options.authorization }
          : {}),
    },
    body: options?.rawBody ?? JSON.stringify(options?.body ?? BODY),
  });
  return POST(request as unknown as NextRequest);
}

describe('POST /api/s2s/classroom-generations', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN', TOKEN);
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN_PREVIOUS', '');
    mocks.after.mockClear();
    mocks.createOrGet.mockReset();
    mocks.run.mockClear();
  });

  afterEach(() => vi.unstubAllEnvs());

  it('authenticates before parsing or creating a job', async () => {
    const response = await submit({ authorization: '', rawBody: '{broken' });

    expect(response.status).toBe(401);
    expect(mocks.createOrGet).not.toHaveBeenCalled();
  });

  it('rejects invalid JSON and non-allowlisted input without creating a job', async () => {
    const invalidJson = await submit({ rawBody: '{broken' });
    const extraField = await submit({
      body: { ...BODY, pdfContent: { text: 'x', images: [] } },
    });

    expect(invalidJson.status).toBe(400);
    expect(extraField.status).toBe(400);
    expect(mocks.createOrGet).not.toHaveBeenCalled();
  });

  it('returns 202 and schedules a newly created job without exposing input or internal URLs', async () => {
    mocks.createOrGet.mockResolvedValue({ outcome: 'created', job: JOB });

    const response = await submit();
    const body = await response.json();

    expect(response.status).toBe(202);
    expect(body).toEqual({
      success: true,
      jobId: 's2s_job',
      status: 'queued',
      step: 'queued',
      pollPath: '/api/s2s/classroom-generations/s2s_job',
      pollIntervalMs: 5000,
      replayed: false,
    });
    expect(JSON.stringify(body)).not.toContain(BODY.requirement);
    expect(JSON.stringify(body)).not.toContain('openmaic:3000');
    expect(mocks.run).toHaveBeenCalledOnce();
  });

  it('returns 200 and reschedules a queued replay so restart recovery can continue', async () => {
    mocks.createOrGet.mockResolvedValue({ outcome: 'replayed', job: JOB });

    const response = await submit();
    const body = await response.json();

    expect(response.status).toBe(200);
    expect(body.replayed).toBe(true);
    expect(mocks.run).toHaveBeenCalledOnce();
  });

  it('does not reschedule a replay that is already running or terminal', async () => {
    mocks.createOrGet.mockResolvedValue({
      outcome: 'replayed',
      job: { ...JOB, status: 'running', step: 'generating_scenes' },
    });

    const response = await submit();

    expect(response.status).toBe(200);
    expect(mocks.run).not.toHaveBeenCalled();
    expect(mocks.after).not.toHaveBeenCalled();
  });

  it('returns 409 when an idempotency key is reused for another request body', async () => {
    mocks.createOrGet.mockResolvedValue({
      outcome: 'conflict',
      jobId: 's2s_job',
    });

    const response = await submit();
    const body = await response.json();

    expect(response.status).toBe(409);
    expect(body).toMatchObject({
      success: false,
      errorCode: 'IDEMPOTENCY_CONFLICT',
    });
    expect(mocks.run).not.toHaveBeenCalled();
  });
});
