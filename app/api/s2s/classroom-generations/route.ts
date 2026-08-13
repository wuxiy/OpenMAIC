import { after, type NextRequest } from 'next/server';
import { apiError, apiSuccess } from '@/lib/server/api-response';
import { authenticateS2SRequest } from '@/lib/server/s2s-auth';
import { parseS2SClassroomGenerationInput } from '@/lib/server/s2s-classroom-input';
import { runClassroomGenerationJob } from '@/lib/server/classroom-job-runner';
import { createOrGetClassroomGenerationJob } from '@/lib/server/classroom-job-store';
import { createLogger } from '@/lib/logger';

const log = createLogger('S2S ClassroomGeneration API');

export const maxDuration = 30;

export async function POST(request: NextRequest) {
  const authError = authenticateS2SRequest(request);
  if (authError) return authError;

  let rawBody: unknown;
  try {
    rawBody = await request.json();
  } catch {
    return apiError('INVALID_REQUEST', 400, 'Invalid JSON body');
  }

  const parsed = parseS2SClassroomGenerationInput(rawBody, request.headers.get('idempotency-key'));
  if (!parsed.ok) return apiError(parsed.errorCode, parsed.status, parsed.error);

  try {
    const result = await createOrGetClassroomGenerationJob(
      parsed.idempotencyKey,
      parsed.requestChecksum,
      parsed.input,
    );
    if (result.outcome === 'conflict') {
      return apiError(
        'IDEMPOTENCY_CONFLICT',
        409,
        'Idempotency key was already used for another request',
      );
    }

    const { job } = result;
    if (result.outcome === 'created' || job.status === 'queued') {
      // The deterministic job id and runner map make this safe for a queued
      // replay after restart and prevent duplicate work in the same process.
      after(() => runClassroomGenerationJob(job.id, parsed.input, new URL(request.url).origin));
    }

    return apiSuccess(
      {
        jobId: job.id,
        status: job.status,
        step: job.step,
        pollPath: `/api/s2s/classroom-generations/${job.id}`,
        pollIntervalMs: 5000,
        replayed: result.outcome === 'replayed',
      },
      result.outcome === 'created' ? 202 : 200,
    );
  } catch (error) {
    log.error('S2S classroom generation job creation failed', error);
    return apiError('INTERNAL_ERROR', 500, 'Failed to create classroom generation job');
  }
}
