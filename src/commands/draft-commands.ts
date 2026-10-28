import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';
import { bodyContent, recipients } from '../utils/compose.js';
import { messageId } from '../utils/message-id.js';

interface DraftOptions {
  update?: string;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  html: boolean;
}

function parseArgs(args: string[]): DraftOptions {
  const options: DraftOptions = { html: false };
  const values = new Map<string, string[]>();
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '--html') { options.html = true; continue; }
    if (!['--update', '--to', '--cc', '--bcc', '--subject', '--body'].includes(arg)) {
      usage(`Unknown draft option: ${arg}.`);
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
  options.update = one('--update');
  options.to = values.get('--to');
  options.cc = values.get('--cc');
  options.bcc = values.get('--bcc');
  options.subject = one('--subject');
  options.body = one('--body');
  if (options.update !== undefined && !options.update) usage("Option '--update' requires a Draft id.");
  if (options.body === undefined && options.html) usage("Option '--html' requires '--body'.");
  return options;
}

export async function runDraftCommand(ctx: Context, args: string[]): Promise<void> {
  const options = parseArgs(args);
  const body = options.body === undefined ? undefined : {
    contentType: options.html ? 'HTML' : 'Text',
    content: await bodyContent(options.body),
  };
  const patch: Record<string, unknown> = {
    ...(options.to ? { toRecipients: recipients(options.to) } : {}),
    ...(options.cc ? { ccRecipients: recipients(options.cc) } : {}),
    ...(options.bcc ? { bccRecipients: recipients(options.bcc) } : {}),
    ...(options.subject === undefined ? {} : { subject: options.subject }),
    ...(body === undefined ? {} : { body }),
  };
  const method = options.update === undefined ? 'POST' : 'PATCH';
  const path = options.update === undefined ? '/me/messages' : `/me/messages/${encodeURIComponent(messageId(options.update))}`;
  const response = await graphRequest(ctx, method, path, { body: JSON.stringify(patch) });
  if (options.update === undefined) {
    const id = (response as { id?: string } | undefined)?.id;
    process.stdout.write(id ? `Draft created: ${id}\n` : 'Draft created.\n');
  } else process.stdout.write(`Draft updated: ${options.update}\n`);
}
