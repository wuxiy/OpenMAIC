import { createHash } from 'node:crypto';

import { describe, expect, it, vi } from 'vitest';

import type { ServerPptxResult } from '@/lib/export/server-pptx';
import type { PersistedClassroomData } from '@/lib/server/classroom-storage';
import { createClassroomPptxExporter } from '@/lib/server/classroom-pptx-export';

const slide = (id: string, order: number, viewportSize = 960, viewportRatio = 0.5625) => ({
  id: `scene-${id}`,
  stageId: 'stage-taohuayuanji',
  title: id,
  order,
  type: 'slide' as const,
  content: {
    type: 'slide' as const,
    canvas: {
      id: `slide-${id}`,
      viewportSize,
      viewportRatio,
      theme: {
        backgroundColor: '#F7F2E8',
        themeColors: ['#395B64'],
        fontColor: '#2F302D',
        fontName: 'Microsoft YaHei',
      },
      elements: [],
    },
  },
});

const classroom = (scenes = [slide('cover', 0)]): PersistedClassroomData =>
  ({
    id: 'classroom-taohuayuanji',
    stage: {
      id: 'stage-taohuayuanji',
      name: '《桃花源记》第一课时 / 教师版',
      description: '八年级语文固定验收课题',
      createdAt: 1_786_563_200_000,
      updatedAt: 1_786_563_200_000,
    },
    scenes,
    createdAt: '2026-08-13T00:00:00.000Z',
  }) as PersistedClassroomData;

describe('createClassroomPptxExporter', () => {
  it('returns bounded PPTX bytes with stable metadata and a safe Chinese filename', async () => {
    const bytes = Uint8Array.from([0x50, 0x4b, 1, 2, 3]);
    const buildPptx = vi.fn().mockResolvedValue({ bytes, slideCount: 1, skippedElements: [] });
    const exporter = createClassroomPptxExporter({ buildPptx });

    const result = await exporter.export(classroom());

    expect(buildPptx).toHaveBeenCalledOnce();
    expect(result).toEqual({
      bytes,
      slideCount: 1,
      checksum: createHash('sha256').update(bytes).digest('hex'),
      filename: '桃花源记第一课时-教师版.pptx',
    });
  });

  it('rejects no slides, inconsistent viewports, and excessive slide counts before building', async () => {
    const buildPptx = vi.fn();
    const exporter = createClassroomPptxExporter({ buildPptx, maxSlides: 2 });

    await expect(exporter.export(classroom([]))).rejects.toMatchObject({
      code: 'NO_SLIDES',
    });
    await expect(
      exporter.export(classroom([slide('one', 0), slide('two', 1, 1280)])),
    ).rejects.toMatchObject({ code: 'INCONSISTENT_VIEWPORT' });
    await expect(
      exporter.export(classroom([slide('one', 0), slide('two', 1), slide('three', 2)])),
    ).rejects.toMatchObject({ code: 'TOO_MANY_SLIDES' });
    expect(buildPptx).not.toHaveBeenCalled();
  });

  it('rejects oversized output and exporter timeout with stable error codes', async () => {
    const oversized = createClassroomPptxExporter({
      buildPptx: vi.fn().mockResolvedValue({
        bytes: new Uint8Array(6),
        slideCount: 1,
        skippedElements: [],
      }),
      maxBytes: 5,
    });
    await expect(oversized.export(classroom())).rejects.toMatchObject({
      code: 'TOO_LARGE',
    });

    const neverCompletes = vi.fn(
      (): Promise<ServerPptxResult> => new Promise<ServerPptxResult>(() => undefined),
    );
    const timedOut = createClassroomPptxExporter({
      buildPptx: neverCompletes,
      timeoutMs: 5,
    });
    await expect(timedOut.export(classroom())).rejects.toMatchObject({
      code: 'EXPORT_TIMEOUT',
    });
  });

  it('maps unsupported persisted elements to a stable boundary code', async () => {
    const unsupported = new Error('private element details');
    unsupported.name = 'UnsupportedServerPptxElementsError';
    const exporter = createClassroomPptxExporter({
      buildPptx: vi.fn().mockRejectedValue(unsupported),
    });

    await expect(exporter.export(classroom())).rejects.toMatchObject({
      code: 'EXPORT_UNSUPPORTED',
      message: 'EXPORT_UNSUPPORTED',
    });
  });
});
