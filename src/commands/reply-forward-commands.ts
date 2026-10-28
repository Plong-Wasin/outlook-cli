import { readFileSync } from 'node:fs';
import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';
import { messageId } from '../utils/message-id.js';

type ReplyCommand = 'reply' | 'reply-all' | 'forward';

interface ReplyOptions {
  id?: string;
  body?: string;
  to?: string;
  draft: boolean;
  yes: boolean;
}

function parseArgs(command: ReplyCommand, args: string[]): ReplyOptions {
  const options: ReplyOptions = { draft: false, yes: false };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '--draft') { options.draft = true; continue; }
    if (arg === '--yes') { options.yes = true; continue; }
    if (arg === '--body' || arg === '--to') {
      const value = args[index + 1];
      if (value === undefined || (value.startsWith('--') && value !== '-')) usage(`Option '${arg}' requires a value.`);
      const property = arg === '--body' ? 'body' : 'to';
      if (options[property] !== undefined) usage(`Option '${arg}' may be used only once.`);
      options[property] = value;
      index += 1;
      continue;
    }
    if (arg.startsWith('-')) usage(`Unknown ${command} option: ${arg}.`);
    if (options.id !== undefined) usage(`Usage: outlook ${command} <id> --body <text> [--draft|--yes]`);
    options.id = arg;
  }
  if (!options.id) usage(`Usage: outlook ${command} <id> --body <text> [--draft|--yes]`);
  if (options.body === undefined) usage("Option '--body' is required.");
  if (options.to !== undefined && command !== 'forward') usage("Option '--to' is only supported by 'forward'.");
  if (command === 'forward' && !options.to) usage("Option '--to' is required for 'forward'.");
  if (options.draft && options.yes) usage("Choose either '--draft' or '--yes', not both.");
  return options;
}

export async function runReplyForwardCommand(ctx: Context, command: ReplyCommand, args: string[]): Promise<void> {
  const options = parseArgs(command, args);
  if (!options.draft && !options.yes) usage(`Refusing to ${command} without '--yes' or '--draft'.`);
  const id = messageId(options.id!);
  const input = options.body!;
  let comment: string;
  if (input === '-') {
    comment = '';
    for await (const chunk of process.stdin) comment += chunk.toString();
  } else if (input.startsWith('@')) {
    const path = input.slice(1);
    if (!path) usage("Option '--body @file' requires a file path.");
    try { comment = readFileSync(path, 'utf8'); }
    catch { usage(`Cannot read body file '${path}'. Check the file path and permissions.`); }
  } else comment = input;
  const operation = command === 'reply-all' ? 'replyAll' : command;
  const action = options.draft ? `create${operation[0]!.toUpperCase()}${operation.slice(1)}` : operation;
  const recipients = options.to?.split(',').map((address) => address.trim()).filter(Boolean).map((address) => ({ emailAddress: { address } }));
  if (options.to !== undefined && !recipients?.length) usage("Option '--to' must include at least one email address.");
  const body = command === 'forward'
    ? options.draft ? { message: { toRecipients: recipients }, comment } : { toRecipients: recipients, comment }
    : { comment };
  const response = await graphRequest(ctx, 'POST', `/me/messages/${encodeURIComponent(id)}/${action}`, {
    body: JSON.stringify(body),
  });
  if (options.draft) {
    const draftId = (response as { id?: string } | undefined)?.id;
    process.stdout.write(draftId ? `Draft created: ${draftId}\n` : 'Draft created.\n');
  } else process.stdout.write(`${command === 'reply-all' ? 'Reply-all' : command[0]!.toUpperCase() + command.slice(1)} sent.\n`);
}
