import type { NextRequest } from 'next/server';

import {
  ClassroomPptxExportError,
  classroomPptxExporter,
} from '@/lib/server/classroom-pptx-export';
import { apiError, type ApiErrorCode } from '@/lib/server/api-response';
import { isValidClassroomId, readClassroom } from '@/lib/server/classroom-storage';
import { authenticateS2SRequest } from '@/lib/server/s2s-auth';

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const errorStatus: Record<ClassroomPptxExportError['code'], number> = {
  NO_SLIDES: 422,
  INCONSISTENT_VIEWPORT: 422,
  TOO_MANY_SLIDES: 413,
  TOO_LARGE: 413,
  EXPORT_TIMEOUT: 504,
  EXPORT_UNSUPPORTED: 422,
};

export const dynamic = 'force-dynamic';
export const maxDuration = 60;

export async function GET(
  request: NextRequest,
  context: { params: Promise<{ classroomId: string }> },
) {
  const authError = authenticateS2SRequest(request);
  if (authError) return authError;

  const { classroomId } = await context.params;
  if (!isValidClassroomId(classroomId)) {
    return apiError('INVALID_REQUEST', 400, 'Invalid classroom id');
  }

  const classroom = await readClassroom(classroomId);
  if (!classroom) {
    return apiError('CLASSROOM_NOT_FOUND', 404, 'Classroom not found');
  }

  try {
    const artifact = await classroomPptxExporter.export(classroom);
    const encodedFilename = encodeURIComponent(artifact.filename).replace(
      /['()*]/g,
      (character) => `%${character.charCodeAt(0).toString(16).toUpperCase()}`,
    );
    return new Response(artifact.bytes as BodyInit, {
      status: 200,
      headers: {
        'Content-Type': PPTX_MIME,
        'Content-Length': String(artifact.bytes.byteLength),
        'Content-Disposition': `attachment; filename="openmaic-courseware.pptx"; filename*=UTF-8''${encodedFilename}`,
        'Cache-Control': 'no-store',
        'X-Content-SHA256': artifact.checksum,
        'X-Slide-Count': String(artifact.slideCount),
      },
    });
  } catch (error) {
    if (error instanceof ClassroomPptxExportError) {
      return apiError(
        error.code as ApiErrorCode,
        errorStatus[error.code],
        'Classroom PPTX export failed',
      );
    }
    return apiError('EXPORT_FAILED', 500, 'Classroom PPTX export failed');
  }
}
