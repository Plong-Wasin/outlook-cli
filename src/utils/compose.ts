import { readFileSync } from 'node:fs';
import { usage } from './errors.js';

export type Recipient = { emailAddress: { address: string } };

export function recipients(values: string[] | undefined): Recipient[] | undefined {
  if (!values) return undefined;
  const addresses = values.flatMap((value) => value.split(',')).map((address) => address.trim()).filter(Boolean);
  if (addresses.length === 0) usage('Recipient options must include at least one email address.');
  return addresses.map((address) => ({ emailAddress: { address } }));
}

// A body is literal text, `@file`, or `-` for stdin.
export async function bodyContent(value: string): Promise<string> {
  if (value === '-') {
    let body = '';
    for await (const chunk of process.stdin) body += chunk.toString();
    return body;
  }
  if (value.startsWith('@')) {
    const path = value.slice(1);
    if (!path) usage("Option '--body @file' requires a file path.");
    try { return readFileSync(path, 'utf8'); }
    catch { usage(`Cannot read body file '${path}'. Check the file path and permissions.`); }
  }
  return value;
}
