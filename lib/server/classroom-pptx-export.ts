import { createHash } from 'node:crypto';

import { buildServerPptx, type ServerPptxResult } from '@/lib/export/server-pptx';
import type { PersistedClassroomData } from '@/lib/server/classroom-storage';

export type ClassroomPptxExportErrorCode =
  | 'NO_SLIDES'
  | 'INCONSISTENT_VIEWPORT'
  | 'TOO_MANY_SLIDES'
  | 'TOO_LARGE'
  | 'EXPORT_TIMEOUT'
  | 'EXPORT_UNSUPPORTED';

export class ClassroomPptxExportError extends Error {
  constructor(public readonly code: ClassroomPptxExportErrorCode) {
    super(code);
    this.name = 'ClassroomPptxExportError';
  }
}

export interface ClassroomPptxArtifact {
  bytes: Uint8Array;
  slideCount: number;
  checksum: string;
  filename: string;
}

export interface ClassroomPptxExportLimits {
  maxSlides?: number;
  maxBytes?: number;
  timeoutMs?: number;
}

interface ClassroomPptxExporterDependencies extends ClassroomPptxExportLimits {
  buildPptx?: (input: {
    stage: PersistedClassroomData['stage'];
    scenes: PersistedClassroomData['scenes'];
  }) => Promise<ServerPptxResult>;
}

const DEFAULT_MAX_SLIDES = 60;
const DEFAULT_MAX_BYTES = 25 * 1024 * 1024;
const DEFAULT_TIMEOUT_MS = 60_000;

function safeFilename(title: string): string {
  const base = title
    .normalize('NFKC')
    .replace(/[\r\n《》]/g, '')
    .replace(/[\\/:*?"<>|\s]+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^[.\-]+|[.\-]+$/g, '')
    .slice(0, 80);
  return `${base || 'OpenMAIC课堂'}.pptx`;
}

function slideScenes(classroom: PersistedClassroomData) {
  return classroom.scenes
    .filter((scene) => scene.content.type === 'slide')
    .sort((left, right) => left.order - right.order);
}

export function createClassroomPptxExporter(dependencies: ClassroomPptxExporterDependencies = {}) {
  const buildPptx = dependencies.buildPptx ?? buildServerPptx;
  const maxSlides = dependencies.maxSlides ?? DEFAULT_MAX_SLIDES;
  const maxBytes = dependencies.maxBytes ?? DEFAULT_MAX_BYTES;
  const timeoutMs = dependencies.timeoutMs ?? DEFAULT_TIMEOUT_MS;

  return {
    async export(classroom: PersistedClassroomData): Promise<ClassroomPptxArtifact> {
      const slides = slideScenes(classroom);
      if (slides.length === 0) throw new ClassroomPptxExportError('NO_SLIDES');
      if (slides.length > maxSlides) throw new ClassroomPptxExportError('TOO_MANY_SLIDES');

      const firstCanvas = slides[0].content.type === 'slide' ? slides[0].content.canvas : null;
      const consistent =
        firstCanvas &&
        slides.every(
          (scene) =>
            scene.content.type === 'slide' &&
            scene.content.canvas.viewportSize === firstCanvas.viewportSize &&
            scene.content.canvas.viewportRatio === firstCanvas.viewportRatio,
        );
      if (!consistent) throw new ClassroomPptxExportError('INCONSISTENT_VIEWPORT');

      let timer: ReturnType<typeof setTimeout> | undefined;
      try {
        const result = await Promise.race([
          buildPptx({ stage: classroom.stage, scenes: classroom.scenes }),
          new Promise<never>((_, reject) => {
            timer = setTimeout(
              () => reject(new ClassroomPptxExportError('EXPORT_TIMEOUT')),
              timeoutMs,
            );
          }),
        ]);
        if (result.bytes.byteLength > maxBytes) {
          throw new ClassroomPptxExportError('TOO_LARGE');
        }
        return {
          bytes: result.bytes,
          slideCount: result.slideCount,
          checksum: createHash('sha256').update(result.bytes).digest('hex'),
          filename: safeFilename(classroom.stage.name),
        };
      } catch (error) {
        if (error instanceof ClassroomPptxExportError) throw error;
        if (error instanceof Error && error.name === 'UnsupportedServerPptxElementsError') {
          throw new ClassroomPptxExportError('EXPORT_UNSUPPORTED');
        }
        throw error;
      } finally {
        if (timer) clearTimeout(timer);
      }
    },
  };
}

export const classroomPptxExporter = createClassroomPptxExporter();
