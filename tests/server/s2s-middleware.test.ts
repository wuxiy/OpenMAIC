import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { NextRequest } from 'next/server';

async function run(path: string) {
  const { middleware } = await import('@/middleware');
  return middleware(new NextRequest(`http://localhost${path}`));
}

describe('S2S middleware boundary', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubEnv('ACCESS_CODE', 'human-browser-access-code');
  });

  afterEach(() => vi.unstubAllEnvs());

  it('lets the S2S namespace reach its own Bearer authentication', async () => {
    const response = await run('/api/s2s/classroom-generations');

    expect(response.status).toBe(200);
    expect(response.headers.get('x-middleware-next')).toBe('1');
  });

  it('continues to require the signed browser cookie for other API routes', async () => {
    const response = await run('/api/generate-classroom');
    const body = await response.json();

    expect(response.status).toBe(401);
    expect(body).toMatchObject({
      success: false,
      error: 'Access code required',
    });
  });
});
