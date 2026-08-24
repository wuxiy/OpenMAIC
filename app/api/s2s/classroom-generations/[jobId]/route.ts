import { type NextRequest } from 'next/server';
import { apiError, apiSuccess } from '@/lib/server/api-response';
import { authenticateS2SRequest } from '@/lib/server/s2s-auth';
import {
  isValidS2SClassroomJobId,
  readClassroomGenerationJob,
} from '@/lib/server/classroom-job-store';
import { createLogger } from '@/lib/logger';

const log = createLogger('S2S ClassroomJob API');

export const dynamic = 'force-dynamic';

export async function GET(request: NextRequest, context: { params: Promise<{ jobId: string }> }) {
  const authError = authenticateS2SRequest(request);
  if (authError) return authError;

  const { jobId } = await context.params;
  if (!isValidS2SClassroomJobId(jobId)) {
    return apiError('INVALID_REQUEST', 400, 'Invalid classroom generation job id');
  }

  try {
    const job = await readClassroomGenerationJob(jobId);
    if (!job) {
      return apiError('INVALID_REQUEST', 404, 'Classroom generation job not found');
    }

    const done = job.status === 'succeeded' || job.status === 'failed';
    return apiSuccess({
      jobId: job.id,
      status: job.status,
      step: job.step,
      progress: job.progress,
      pollPath: `/api/s2s/classroom-generations/${job.id}`,
      pollIntervalMs: 5000,
      scenesGenerated: job.scenesGenerated,
      totalScenes: job.totalScenes,
      result: job.result
        ? {
            classroomId: job.result.classroomId,
            scenesCount: job.result.scenesCount,
          }
        : undefined,
      errorCode: job.status === 'failed' ? 'GENERATION_FAILED' : undefined,
      done,
    });
  } catch (error) {
    log.error(`S2S classroom job retrieval failed [jobId=${jobId}]`, error);
    return apiError('INTERNAL_ERROR', 500, 'Failed to retrieve classroom generation job');
  }
}
