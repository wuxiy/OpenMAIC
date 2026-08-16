import { NextResponse } from 'next/server';

export const API_ERROR_CODES = {
  MISSING_REQUIRED_FIELD: 'MISSING_REQUIRED_FIELD',
  MISSING_API_KEY: 'MISSING_API_KEY',
  INVALID_CREDENTIALS: 'INVALID_CREDENTIALS',
  INVALID_REQUEST: 'INVALID_REQUEST',
  PROVIDER_DISABLED: 'PROVIDER_DISABLED',
  VOXCPM_AUTO_VOICE_REQUIRES_CONTEXT: 'VOXCPM_AUTO_VOICE_REQUIRES_CONTEXT',
  INVALID_URL: 'INVALID_URL',
  REDIRECT_NOT_ALLOWED: 'REDIRECT_NOT_ALLOWED',
  TOO_MANY_REDIRECTS: 'TOO_MANY_REDIRECTS',
  CONTENT_SENSITIVE: 'CONTENT_SENSITIVE',
  UPSTREAM_ERROR: 'UPSTREAM_ERROR',
  RATE_LIMITED: 'RATE_LIMITED',
  GENERATION_FAILED: 'GENERATION_FAILED',
  TRANSCRIPTION_FAILED: 'TRANSCRIPTION_FAILED',
  PARSE_FAILED: 'PARSE_FAILED',
  INTERNAL_ERROR: 'INTERNAL_ERROR',
  S2S_AUTH_REQUIRED: 'S2S_AUTH_REQUIRED',
  S2S_FORBIDDEN: 'S2S_FORBIDDEN',
  S2S_NOT_CONFIGURED: 'S2S_NOT_CONFIGURED',
  IDEMPOTENCY_CONFLICT: 'IDEMPOTENCY_CONFLICT',
  CLASSROOM_NOT_FOUND: 'CLASSROOM_NOT_FOUND',
  NO_SLIDES: 'NO_SLIDES',
  INCONSISTENT_VIEWPORT: 'INCONSISTENT_VIEWPORT',
  TOO_MANY_SLIDES: 'TOO_MANY_SLIDES',
  EXPORT_UNSUPPORTED: 'EXPORT_UNSUPPORTED',
  EXPORT_FAILED: 'EXPORT_FAILED',
  EXPORT_TIMEOUT: 'EXPORT_TIMEOUT',
  TOO_LARGE: 'TOO_LARGE',
} as const;

export type ApiErrorCode = (typeof API_ERROR_CODES)[keyof typeof API_ERROR_CODES];

export interface ApiErrorBody {
  success: false;
  errorCode: ApiErrorCode;
  error: string;
  details?: string;
}

export function apiError(
  code: ApiErrorCode,
  status: number,
  error: string,
  details?: string,
): NextResponse<ApiErrorBody> {
  return NextResponse.json(
    {
      success: false as const,
      errorCode: code,
      error,
      ...(details ? { details } : {}),
    },
    { status },
  );
}

export function apiSuccess<T extends Record<string, unknown>>(data: T, status = 200): NextResponse {
  return NextResponse.json({ success: true, ...data }, { status });
}
