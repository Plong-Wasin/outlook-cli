import { readFile } from 'node:fs/promises';
import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';
import { validateContinuation } from '../utils/graph-page.js';

export async function runApi(ctx: Context, args: string[]): Promise<void> {
  const method = args[0]?.toUpperCase();
  const path = args[1];
  if (!method || !path) usage("Use 'outlook api <METHOD> <path> [key=value ...]'.");
  if (!['GET', 'POST', 'PUT', 'PATCH', 'DELETE'].includes(method)) usage(`Unsupported HTTP method: ${args[0]}.`);

  const query = new URLSearchParams();
  const headers: Record<string, string> = {};
  let body: string | undefined;
  let all = false;
  for (let index = 2; index < args.length; index += 1) {
    const argument = args[index]!;
    if (argument === '--all') {
      all = true;
      continue;
    }
    if (argument === '--header' || argument.startsWith('--header=')) {
      const header = argument === '--header' ? args[++index] : argument.slice('--header='.length);
      const colon = header?.indexOf(':') ?? -1;
      if (!header || colon <= 0) usage("Option '--header' requires 'Name: value'.");
      headers[header.slice(0, colon).trim()] = header.slice(colon + 1).trim();
      continue;
    }
    if (argument === '--data' || argument.startsWith('--data=')) {
      if (body !== undefined) usage("Option '--data' may only be provided once.");
      const source = argument === '--data' ? args[++index] : argument.slice('--data='.length);
      if (source === undefined) usage("Option '--data' requires JSON, @file, or '-'.");
      const text = source === '-' ? await readStdin() : source.startsWith('@') ? await readFile(source.slice(1), 'utf8') : source;
      try {
        body = JSON.stringify(JSON.parse(text))!;
      } catch {
        usage("Option '--data' must contain valid JSON.");
      }
      continue;
    }
    const separator = argument.indexOf('=');
    if (separator <= 0) usage(`Expected a query pair in key=value form: ${argument}.`);
    query.append(argument.slice(0, separator), argument.slice(separator + 1));
  }
  if (all && method !== 'GET') usage("Option '--all' is only supported with GET requests.");

  let data = await graphRequest(ctx, method, path, { query, body, headers });
  if (all) data = await collectPages(ctx, data);
  if (data !== undefined) process.stdout.write(`${JSON.stringify(data)}\n`);
}

async function collectPages(ctx: Context, initial: unknown): Promise<unknown> {
  if (!isCollection(initial)) return initial;
  const values = [...initial.value];
  const result: Record<string, unknown> = { ...initial, value: values };
  delete result['@odata.nextLink'];
  let nextLink = initial['@odata.nextLink'];
  for (let page = 1; typeof nextLink === 'string' && page < 50; page += 1) {
    validateContinuation(nextLink, ctx);
    const nextUrl = new URL(nextLink);
    const data = await graphRequest(ctx, 'GET', nextUrl.pathname, { rawUrl: nextUrl.href });
    if (!isCollection(data)) break;
    values.push(...data.value);
    nextLink = data['@odata.nextLink'];
  }
  return result;
}

function isCollection(value: unknown): value is Record<string, unknown> & { value: unknown[] } {
  return typeof value === 'object' && value !== null && Array.isArray((value as { value?: unknown }).value);
}

async function readStdin(): Promise<string> {
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString('utf8');
}
