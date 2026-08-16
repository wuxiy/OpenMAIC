import { createHash, timingSafeEqual } from 'crypto';
import { apiError } from '@/lib/server/api-response';

const MIN_TOKEN_LENGTH = 32;

function tokenDigest(value: string): Buffer {
  return createHash('sha256').update(value, 'utf8').digest();
}

function isConfiguredToken(value: string | undefined): value is string {
  return !!value && value.length >= MIN_TOKEN_LENGTH;
}

/** Return null when authorized, otherwise a complete fail-closed API response. */
export function authenticateS2SRequest(request: Request): Response | null {
  const primary = process.env.OPENMAIC_SERVICE_TOKEN;
  const previous = process.env.OPENMAIC_SERVICE_TOKEN_PREVIOUS;
  if (!isConfiguredToken(primary) || (previous && !isConfiguredToken(previous))) {
    return apiError('S2S_NOT_CONFIGURED', 503, 'Service authentication is not configured');
  }

  const authorization = request.headers.get('authorization');
  if (!authorization?.startsWith('Bearer ') || authorization.length === 'Bearer '.length) {
    return apiError('S2S_AUTH_REQUIRED', 401, 'Bearer service token required');
  }

  const suppliedDigest = tokenDigest(authorization.slice('Bearer '.length));
  const primaryMatch = timingSafeEqual(suppliedDigest, tokenDigest(primary));
  // Always perform a second fixed-length comparison, even outside a rotation window.
  const previousMatch = timingSafeEqual(suppliedDigest, tokenDigest(previous || primary));
  if (!(primaryMatch || (previous ? previousMatch : false))) {
    return apiError('S2S_FORBIDDEN', 403, 'Invalid service token');
  }

  return null;
}
