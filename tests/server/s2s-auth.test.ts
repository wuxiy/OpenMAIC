import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const PRIMARY_TOKEN = 'primary-service-token-that-is-at-least-32-bytes';
const PREVIOUS_TOKEN = 'previous-service-token-that-is-at-least-32-bytes';

async function authenticate(authorization?: string) {
  const { authenticateS2SRequest } = await import('@/lib/server/s2s-auth');
  const headers = authorization ? { Authorization: authorization } : undefined;
  return authenticateS2SRequest(new Request('http://localhost/api/s2s/test', { headers }));
}

describe('S2S Bearer authentication', () => {
  beforeEach(() => {
    vi.resetModules();
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN', PRIMARY_TOKEN);
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN_PREVIOUS', '');
  });

  afterEach(() => vi.unstubAllEnvs());

  it('fails closed with 503 when the primary service token is missing or too short', async () => {
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN', '');
    const missing = await authenticate(`Bearer ${PRIMARY_TOKEN}`);
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN', 'too-short');
    const short = await authenticate(`Bearer ${PRIMARY_TOKEN}`);

    expect(missing?.status).toBe(503);
    expect(short?.status).toBe(503);
    await expect(missing?.json()).resolves.toMatchObject({
      success: false,
      errorCode: 'S2S_NOT_CONFIGURED',
    });
  });

  it('returns 401 when the Authorization header is missing or is not Bearer', async () => {
    const missing = await authenticate();
    const basic = await authenticate('Basic abc');

    expect(missing?.status).toBe(401);
    expect(basic?.status).toBe(401);
    await expect(missing?.json()).resolves.toMatchObject({
      success: false,
      errorCode: 'S2S_AUTH_REQUIRED',
    });
  });

  it('returns 403 for wrong tokens regardless of their original length', async () => {
    const shorter = await authenticate('Bearer wrong');
    const longer = await authenticate(`Bearer ${'wrong'.repeat(30)}`);

    expect(shorter?.status).toBe(403);
    expect(longer?.status).toBe(403);
    await expect(shorter?.json()).resolves.toMatchObject({
      success: false,
      errorCode: 'S2S_FORBIDDEN',
    });
  });

  it('accepts the primary token and an explicitly configured previous token', async () => {
    const primary = await authenticate(`Bearer ${PRIMARY_TOKEN}`);
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN_PREVIOUS', PREVIOUS_TOKEN);
    const previous = await authenticate(`Bearer ${PREVIOUS_TOKEN}`);

    expect(primary).toBeNull();
    expect(previous).toBeNull();
  });

  it('fails closed when a configured previous token is too short', async () => {
    vi.stubEnv('OPENMAIC_SERVICE_TOKEN_PREVIOUS', 'bad-old-token');
    const response = await authenticate(`Bearer ${PRIMARY_TOKEN}`);
    expect(response?.status).toBe(503);
  });
});
