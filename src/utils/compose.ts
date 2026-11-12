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

// Graph accepts inline attachment bytes up to 3 MB per file. Files that travel in one
// request body (`send`, new Draft) must fit in 2.5 MB together: base64 adds a third, and
// the whole body has to stay under Graph's ~4 MB request limit.
const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;
const MAX_TOGETHER_BYTES = 2.5 * 1024 * 1024;

export function attachments(paths: string[] | undefined, together = false): FileAttachment[] | undefined {
  if (!paths) return undefined;
  let total = 0;
  const files = paths.map((path) => {
    let bytes: Buffer;
    try { bytes = readFileSync(path); }
    catch { usage(`Cannot read attachment '${path}'. Check the file path and permissions.`); }
    if (bytes.length > MAX_ATTACHMENT_BYTES) usage(`Attachment '${path}' is larger than 3 MB.`);
    total += bytes.length;
    return { '@odata.type': '#microsoft.graph.fileAttachment' as const, name: basename(path), contentBytes: bytes.toString('base64') };
  });
  if (together && total > MAX_TOGETHER_BYTES) usage('Attachments sent in one request must total 2.5 MB or less. Save a Draft with fewer files, then add the rest with \'draft --update\'.');
  return files;
}
