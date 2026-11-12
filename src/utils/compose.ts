import { readFileSync } from 'node:fs';
import { basename } from 'node:path';
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

export type FileAttachment = { '@odata.type': '#microsoft.graph.fileAttachment'; name: string; contentBytes: string };

// Graph accepts an inline attachment only when the file is under 3 MB
// (https://learn.microsoft.com/en-us/graph/api/message-post-attachments).
const ATTACHMENT_LIMIT_BYTES = 3_000_000;

export function attachments(paths: string[] | undefined): FileAttachment[] | undefined {
  if (!paths) return undefined;
  return paths.map((path) => {
    let bytes: Buffer;
    try { bytes = readFileSync(path); }
    catch { usage(`Cannot read attachment '${path}'. Check the file path and permissions.`); }
    if (bytes.length >= ATTACHMENT_LIMIT_BYTES) usage(`Attachment '${path}' must be under 3 MB.`);
    return { '@odata.type': '#microsoft.graph.fileAttachment' as const, name: basename(path), contentBytes: bytes.toString('base64') };
  });
}
