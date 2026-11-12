import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { attachments, bodyContent, recipients } from '../utils/compose.js';
import { messageId } from '../utils/message-id.js';

interface DraftOptions {
  update?: string;
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  body?: string;
  attach?: string[];
  html: boolean;
}

function parseArgs(args: string[]): DraftOptions {
  const options: DraftOptions = { html: false };
  const values = new Map<string, string[]>();
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '--html') { options.html = true; continue; }
    if (!['--update', '--to', '--cc', '--bcc', '--subject', '--body', '--attach'].includes(arg)) {
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
  options.attach = values.get('--attach');
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
  const files = attachments(options.attach, options.update === undefined);
  const patch: Record<string, unknown> = {
    ...(files && options.update === undefined ? { attachments: files } : {}),
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
  } else {
    // Graph cannot patch attachments onto an existing Draft; add each one separately.
    const updateId = messageId(options.update);
    for (const file of files ?? []) {
      try {
        await graphRequest(ctx, 'POST', `/me/messages/${encodeURIComponent(updateId)}/attachments`, { body: JSON.stringify(file) });
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error);
        throw new CliError(`Draft ${options.update} was updated but '${file.name}' was not attached (files listed as 'Attached' stay on the Draft; do not repeat them): ${detail}`, error instanceof CliError ? error.exitCode : 1);
      }
      process.stdout.write(`Attached: ${file.name}\n`);
    }
    process.stdout.write(`Draft updated: ${options.update}\n`);
  }
}
