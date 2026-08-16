import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { NextRequest } from 'next/server';

const mocks = vi.hoisted(() => ({
  authenticate: vi.fn(),
  readClassroom: vi.fn(),
  exportPptx: vi.fn(),
}));

vi.mock('@/lib/server/s2s-auth', () => ({
  authenticateS2SRequest: mocks.authenticate,
}));

vi.mock('@/lib/server/classroom-storage', async () => {
  const actual = await vi.importActual<typeof import('@/lib/server/classroom-storage')>(
    '@/lib/server/classroom-storage',
  );
  return { ...actual, readClassroom: mocks.readClassroom };
});

vi.mock('@/lib/server/classroom-pptx-export', async () => {
  const actual = await vi.importActual<typeof import('@/lib/server/classroom-pptx-export')>(
    '@/lib/server/classroom-pptx-export',
  );
  return {
    ...actual,
    classroomPptxExporter: { export: mocks.exportPptx },
  };
});

function request() {
  return new Request('http://localhost/api/s2s/classrooms/classroom-1/pptx', {
    headers: { Authorization: 'Bearer test-token' },
  }) as unknown as NextRequest;
}

async function get(classroomId = 'classroom-1') {
  const { GET } = await import('@/app/api/s2s/classrooms/[classroomId]/pptx/route');
  return GET(request(), { params: Promise.resolve({ classroomId }) });
}

describe('GET /api/s2s/classrooms/:classroomId/pptx', () => {
  beforeEach(() => {
    vi.resetModules();
    mocks.authenticate.mockReset().mockReturnValue(null);
    mocks.readClassroom.mockReset().mockResolvedValue({ id: 'classroom-1' });
    mocks.exportPptx.mockReset().mockResolvedValue({
      bytes: Uint8Array.from([0x50, 0x4b, 1, 2, 3]),
      slideCount: 12,
      checksum: 'a'.repeat(64),
      filename: '桃花源记第一课时.pptx',
    });
  });

  it('authenticates before validating or reading classroom state', async () => {
    mocks.authenticate.mockReturnValue(
      Response.json({ success: false, errorCode: 'S2S_FORBIDDEN' }, { status: 403 }),
    );

    const response = await get('../escape');

    expect(response.status).toBe(403);
    expect(mocks.readClassroom).not.toHaveBeenCalled();
    expect(mocks.exportPptx).not.toHaveBeenCalled();
  });

  it('returns 400 for invalid ids and 404 for missing classrooms', async () => {
    expect((await get('../escape')).status).toBe(400);
    mocks.readClassroom.mockResolvedValue(null);
    const missing = await get();
    expect(missing.status).toBe(404);
    await expect(missing.json()).resolves.toMatchObject({
      success: false,
      errorCode: 'CLASSROOM_NOT_FOUND',
    });
  });

  it('returns exact PPTX bytes and strict download headers', async () => {
    const response = await get();

    expect(response.status).toBe(200);
    expect(response.headers.get('content-type')).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    );
    expect(response.headers.get('content-length')).toBe('5');
    expect(response.headers.get('x-content-sha256')).toBe('a'.repeat(64));
    expect(response.headers.get('x-slide-count')).toBe('12');
    expect(response.headers.get('cache-control')).toBe('no-store');
    expect(response.headers.get('content-disposition')).toContain("filename*=UTF-8''");
    expect(new Uint8Array(await response.arrayBuffer())).toEqual(
      Uint8Array.from([0x50, 0x4b, 1, 2, 3]),
    );
  });

  it.each([
    ['NO_SLIDES', 422],
    ['INCONSISTENT_VIEWPORT', 422],
    ['TOO_MANY_SLIDES', 413],
    ['TOO_LARGE', 413],
    ['EXPORT_TIMEOUT', 504],
    ['EXPORT_UNSUPPORTED', 422],
  ] as const)('maps %s to %i without leaking details', async (code, status) => {
    const { ClassroomPptxExportError } = await import('@/lib/server/classroom-pptx-export');
    mocks.exportPptx.mockRejectedValue(new ClassroomPptxExportError(code));

    const response = await get();

    expect(response.status).toBe(status);
    await expect(response.json()).resolves.toEqual({
      success: false,
      errorCode: code,
      error: 'Classroom PPTX export failed',
    });
  });
});
