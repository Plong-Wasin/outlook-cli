import { graphRequest, type Context } from '../commands/auth-commands.js';
import { usage } from './errors.js';

export function graphPage(ctx: Context, url: string, headers?: Record<string, string>): Promise<unknown> {
  return /^https?:\/\//i.test(url)
    ? graphRequest(ctx, 'GET', url, { rawUrl: url, ...(headers ? { headers } : {}) })
    : graphRequest(ctx, 'GET', url, headers ? { headers } : undefined);
}

// A nextLink is followed with the bearer token, so it must stay on the configured Graph host.
export function validateContinuation(link: string, ctx: Context): void {
  const baseValue = ctx.env.OUTLOOK_GRAPH_URL || 'https://graph.microsoft.com/v1.0';
  let url: URL;
  try { url = new URL(link, baseValue); }
  catch { usage('Graph returned an invalid continuation URL.'); }
  const base = new URL(baseValue);
  const basePath = base.pathname.replace(/\/$/, '');
  const inBasePath = !basePath || url.pathname === basePath || url.pathname.startsWith(`${basePath}/`);
  if (url.origin !== base.origin || !inBasePath) usage('Graph returned a continuation URL outside the configured Graph URL.');
}
