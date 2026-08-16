import { describe, expect, it } from 'vitest';
import { parseS2SClassroomGenerationInput } from '@/lib/server/s2s-classroom-input';

const KEY = 'teacheros:123e4567-e89b-42d3-a456-426614174000';
const VALID_INPUT = {
  requirement: '  八年级语文\r\n《桃花源记》第一课时  ',
  enableWebSearch: false,
  enableImageGeneration: false,
  enableVideoGeneration: false,
  enableTTS: false,
  agentMode: 'default',
};

describe('S2S classroom generation input', () => {
  it('accepts only the fixed Teacher-OS profile and returns a stable canonical checksum', () => {
    const result = parseS2SClassroomGenerationInput(VALID_INPUT, KEY);

    expect(result).toEqual({
      ok: true,
      idempotencyKey: KEY,
      input: {
        requirement: '八年级语文\n《桃花源记》第一课时',
        enableWebSearch: false,
        enableImageGeneration: false,
        enableVideoGeneration: false,
        enableTTS: false,
        agentMode: 'default',
      },
      // Independent `printf ... | shasum -a 256` checksum of the canonical JSON literal.
      requestChecksum: 'f1135ac0903e24b16a03c4a7e31742215d5af8f4ba428bfe5301fcce911267e1',
    });
  });

  it.each([undefined, '', 'teacheros:not-a-uuid', 'other:123e4567-e89b-42d3-a456-426614174000'])(
    'rejects an invalid idempotency key: %s',
    (key) => {
      const result = parseS2SClassroomGenerationInput(VALID_INPUT, key);
      expect(result).toMatchObject({
        ok: false,
        status: 400,
        errorCode: 'INVALID_REQUEST',
      });
    },
  );

  it.each([
    [{ ...VALID_INPUT, pdfContent: { text: 'secret', images: [] } }, 'unknown pdfContent'],
    [{ ...VALID_INPUT, providerApiKey: 'secret' }, 'unknown provider key'],
    [{ ...VALID_INPUT, enableWebSearch: true }, 'enabled search'],
    [{ ...VALID_INPUT, enableImageGeneration: true }, 'enabled image generation'],
    [{ ...VALID_INPUT, agentMode: 'generate' }, 'generated agents'],
    [{ ...VALID_INPUT, requirement: '   ' }, 'blank requirement'],
    [null, 'null'],
    [[], 'array'],
  ] as Array<[unknown, string]>)('rejects %s (%s)', (body, _description) => {
    const result = parseS2SClassroomGenerationInput(body, KEY);
    expect(result).toMatchObject({
      ok: false,
      status: 400,
      errorCode: 'INVALID_REQUEST',
    });
  });

  it('treats semantically identical normalized input as the same request', () => {
    const a = parseS2SClassroomGenerationInput(VALID_INPUT, KEY);
    const b = parseS2SClassroomGenerationInput(
      { ...VALID_INPUT, requirement: '八年级语文\n《桃花源记》第一课时' },
      KEY,
    );

    expect(a.ok && b.ok && a.requestChecksum).toBe(b.ok && b.requestChecksum);
  });
});
