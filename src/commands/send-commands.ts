import { readFileSync } from 'node:fs';
import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';

interface SendOptions {
  draft?: string;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  html: boolean;
  yes: boolean;
}

function parseArgs(args: string[]): SendOptions {
  const options: SendOptions = { html: false, yes: false };
  const values = new Map<string, string[]>();
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '--html') { options.html = true; continue; }
    if (arg === '--yes') { options.yes = true; continue; }
    if (!['--draft', '--to', '--cc', '--bcc', '--subject', '--body'].includes(arg)) {
      usage(`Unknown send option: ${arg}.`);
    }
    const value = args[index + 1];
    if (value === undefined || (value.startsWith('--') && value !== '-')) usage(`Option '${arg}' requires a value.`);
    const items = values.get(arg) ?? [];
    items.push(value);
    values.set(arg, items);
    index += 1;
  }
  const one = (name: string): string | undefined => {
    const items = values.get(name);
    if (!items) return undefined;
    if (items.length > 1) usage(`Option '${name}' may be used only once.`);
    return items[0];
  };
  options.draft = one('--draft');
  options.to = values.get('--to');
  options.cc = values.get('--cc');
  options.bcc = values.get('--bcc');
  options.subject = one('--subject');
  options.body = one('--body');
  if (options.draft !== undefined && !options.draft) usage("Option '--draft' requires a Draft id.");
  if (options.draft !== undefined && (options.to || options.cc || options.bcc || options.subject !== undefined || options.body !== undefined || options.html)) {
    usage("Option '--draft' cannot be combined with compose options.");
  }
  if (options.body === undefined && options.html) usage("Option '--html' requires '--body'.");
  if (options.draft === undefined && !options.to && !options.cc && !options.bcc) usage("Provide at least one recipient with '--to', '--cc', or '--bcc'.");
  return options;
}

type Recipient = { emailAddress: { address: string } };
function recipients(values: string[] | undefined): Recipient[] | undefined {
  if (!values) return undefined;
  const addresses = values.flatMap((value) => value.split(',')).map((address) => address.trim()).filter(Boolean);
  if (addresses.length === 0) usage('Recipient options must include at least one email address.');
  return addresses.map((address) => ({ emailAddress: { address } }));
}

async function bodyContent(value: string): Promise<string> {
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

interface MessageSummary {
  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  bccRecipients?: Recipient[];
  subject?: string;
}

function preview(message: MessageSummary): void {
  const addresses = [message.toRecipients, message.ccRecipients, message.bccRecipients]
    .flatMap((group) => group ?? [])
    .map((recipient) => recipient.emailAddress?.address)
    .filter((address): address is string => Boolean(address));
  process.stdout.write(`Recipients: ${addresses.length ? addresses.join(', ') : '(none)'}\nSubject: ${message.subject ?? ''}\n`);
}

export async function runSendCommand(ctx: Context, args: string[]): Promise<void> {
  const options = parseArgs(args);
  if (options.draft !== undefined) {
    const id = options.draft;
    if (!options.yes) {
      const summary = await graphRequest(ctx, 'GET', `/me/messages/${encodeURIComponent(id)}`, {
        query: new URLSearchParams({ '$select': 'toRecipients,ccRecipients,bccRecipients,subject' }),
      }) as MessageSummary;
      preview(summary ?? {});
      usage("Sending a Draft requires '--yes'. Run the command again with '--yes' to send it.");
    }
    await graphRequest(ctx, 'POST', `/me/messages/${encodeURIComponent(id)}/send`);
    process.stdout.write(`Draft sent: ${id}\n`);
    return;
  }

  const message: Record<string, unknown> = {
    ...(options.to ? { toRecipients: recipients(options.to) } : {}),
    ...(options.cc ? { ccRecipients: recipients(options.cc) } : {}),
    ...(options.bcc ? { bccRecipients: recipients(options.bcc) } : {}),
    ...(options.subject === undefined ? {} : { subject: options.subject }),
    ...(options.body === undefined ? {} : { body: {
      contentType: options.html ? 'HTML' : 'Text',
      content: await bodyContent(options.body),
    } }),
  };
  const summary = message as MessageSummary;
  if (!options.yes) {
    preview(summary);
    usage("Sending a Message requires '--yes'. Run the command again with '--yes' to send it.");
  }
  await graphRequest(ctx, 'POST', '/me/sendMail', { body: JSON.stringify({ message, saveToSentItems: true }) });
  process.stdout.write('Message sent.\n');
}
