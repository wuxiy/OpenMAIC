import { createHash } from 'crypto';
import type { GenerateClassroomInput } from '@/lib/server/classroom-generation';
import type { ApiErrorCode } from '@/lib/server/api-response';

const INPUT_KEYS = [
  'agentMode',
  'enableImageGeneration',
  'enableTTS',
  'enableVideoGeneration',
  'enableWebSearch',
  'requirement',
] as const;

const IDEMPOTENCY_KEY =
  /^teacheros:[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;

export type S2SClassroomGenerationInput = Required<
  Pick<
    GenerateClassroomInput,
    | 'requirement'
    | 'enableWebSearch'
    | 'enableImageGeneration'
    | 'enableVideoGeneration'
    | 'enableTTS'
    | 'agentMode'
  >
>;

export type ParseS2SClassroomInputResult =
  | {
      ok: true;
      idempotencyKey: string;
      input: S2SClassroomGenerationInput;
      requestChecksum: string;
    }
  | { ok: false; status: 400; errorCode: ApiErrorCode; error: string };

function invalid(error: string): ParseS2SClassroomInputResult {
  return { ok: false, status: 400, errorCode: 'INVALID_REQUEST', error };
}

function normalizeRequirement(value: string): string {
  return value.replace(/\r\n?/g, '\n').trim();
}

/** Parse the exact, media-free Teacher-OS profile and derive its canonical checksum. */
export function parseS2SClassroomGenerationInput(
  body: unknown,
  idempotencyKey: string | null | undefined,
): ParseS2SClassroomInputResult {
  if (!idempotencyKey || !IDEMPOTENCY_KEY.test(idempotencyKey)) {
    return invalid('Invalid Idempotency-Key');
  }
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    return invalid('Invalid JSON object');
  }

  const record = body as Record<string, unknown>;
  if (
    Object.keys(record).length !== INPUT_KEYS.length ||
    Object.keys(record).some((key) => !(INPUT_KEYS as readonly string[]).includes(key))
  ) {
    return invalid('Unexpected or missing request fields');
  }

  if (
    typeof record.requirement !== 'string' ||
    record.enableWebSearch !== false ||
    record.enableImageGeneration !== false ||
    record.enableVideoGeneration !== false ||
    record.enableTTS !== false ||
    record.agentMode !== 'default'
  ) {
    return invalid('Request must use the fixed Teacher-OS generation profile');
  }

  const requirement = normalizeRequirement(record.requirement);
  if (!requirement) return invalid('Requirement must not be blank');

  const input: S2SClassroomGenerationInput = {
    requirement,
    enableWebSearch: false,
    enableImageGeneration: false,
    enableVideoGeneration: false,
    enableTTS: false,
    agentMode: 'default',
  };
  const canonical = JSON.stringify(input);

  return {
    ok: true,
    idempotencyKey,
    input,
    requestChecksum: createHash('sha256').update(canonical, 'utf8').digest('hex'),
  };
}
