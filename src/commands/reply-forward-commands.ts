import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { messageId } from '../utils/message-id.js';
import { attachments, bodyContent } from '../utils/compose.js';

type ReplyCommand = 'reply' | 'reply-all' | 'forward';

function escapeHtml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// Plain text keeps its line breaks as <br>; with --html the body is already HTML.
function commentHtml(comment: string, isHtml: boolean): string {
  return isHtml ? comment : escapeHtml(comment).replace(/\r?\n/g, '<br>\n');
}

// Put the new text directly after <body>, so the quoted original stays below it.
async function writeHtmlBody(ctx: Context, draftId: string, html: string): Promise<void> {
  const path = `/me/messages/${encodeURIComponent(draftId)}`;
  const draft = await graphRequest(ctx, 'GET', path, { query: new URLSearchParams({ $select: 'body' }), headers: { Prefer: 'outlook.body-content-type="html"' } }) as { body?: { content?: string } } | undefined;
  const existing = draft?.body?.content ?? '';
  const opening = /<body[^>]*>/i.exec(existing);
  const content = opening
    ? existing.slice(0, opening.index + opening[0].length) + html + existing.slice(opening.index + opening[0].length)
    : html + existing;
  await graphRequest(ctx, 'PATCH', path, { body: JSON.stringify({ body: { contentType: 'html', content } }) });
}

interface ReplyOptions {
  id?: string;
  body?: string;
  to?: string;
  attach: string[];
  html: boolean;
  draft: boolean;
  yes: boolean;
}

function parseArgs(command: ReplyCommand, args: string[]): ReplyOptions {
  const options: ReplyOptions = { attach: [], html: false, draft: false, yes: false };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '--draft') { options.draft = true; continue; }
    if (arg === '--yes') { options.yes = true; continue; }
    if (arg === '--html') { options.html = true; continue; }
    if (arg === '--attach') {
      const value = args[index + 1];
      if (value === undefined || (value.startsWith('--') && value !== '-')) usage(`Option '${arg}' requires a value.`);
      options.attach.push(value);
      index += 1;
      continue;
    }
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
    if (options.id !== undefined) usage(`Usage: outlook ${command} <id> --body <text> [--html] [--attach <file>]... [--draft|--yes]`);
    options.id = arg;
  }
  if (!options.id) usage(`Usage: outlook ${command} <id> --body <text> [--html] [--attach <file>]... [--draft|--yes]`);
  if (options.body === undefined) usage("Option '--body' is required.");
  if (options.to !== undefined && command !== 'forward') usage("Option '--to' is only supported by 'forward'.");
  if (command === 'forward' && !options.to) usage("Option '--to' is required for 'forward'.");
  if (options.draft && options.yes) usage("Choose either '--draft' or '--yes', not both.");
  return options;
}

export async function runReplyForwardCommand(ctx: Context, command: ReplyCommand, args: string[]): Promise<void> {
  const options = parseArgs(command, args);
  const id = messageId(options.id!);
  if (!options.draft && !options.yes) {
    process.stdout.write(`Would ${command} Message #${id}${options.to ? ` to ${options.to}` : ''}\n`);
    usage(`Refusing to ${command} without '--yes' or '--draft'.`);
  }
  const input = options.body!;
  // A trailing newline (`@file`, stdin) is not a line break the author meant to keep.
  const raw = await bodyContent(input);
  const comment = options.html ? raw : raw.replace(/(\r?\n)+$/, '');
  const operation = command === 'reply-all' ? 'replyAll' : command;
  const recipients = options.to?.split(',').map((address) => address.trim()).filter(Boolean).map((address) => ({ emailAddress: { address } }));
  if (options.to !== undefined && !recipients?.length) usage("Option '--to' must include at least one email address.");
  const files = attachments(options.attach.length ? options.attach : undefined);
  // Graph's `comment` is plain text that the Draft keeps on one line, so a body with line
  // breaks (or raw HTML) is written into the Draft's HTML body instead, above the quote.
  const htmlBody = options.html || /[\r\n]/.test(comment) ? commentHtml(comment, options.html) : undefined;
  // Attachments and HTML bodies go onto a Draft (documented `/attachments` and PATCH
  // endpoints), so with either of them even `--yes` creates the Draft first and sends it afterwards.
  const viaDraft = options.draft || Boolean(files) || htmlBody !== undefined;
  const action = viaDraft ? `create${operation[0]!.toUpperCase()}${operation.slice(1)}` : operation;
  const requestComment = htmlBody === undefined ? { comment } : {};
  const body = command === 'forward'
    ? viaDraft ? { message: { toRecipients: recipients }, ...requestComment } : { toRecipients: recipients, ...requestComment }
    : requestComment;
  const response = await graphRequest(ctx, 'POST', `/me/messages/${encodeURIComponent(id)}/${action}`, {
    body: JSON.stringify(body),
  });
  const draftId = (response as { id?: string } | undefined)?.id;
  if (files || htmlBody !== undefined) {
    if (!draftId) throw new CliError('Graph did not return a Draft id, so the body and attachments were not added.');
    let sending = false;
    try {
      if (htmlBody !== undefined) await writeHtmlBody(ctx, draftId, htmlBody);
      for (const file of files ?? []) {
        await graphRequest(ctx, 'POST', `/me/messages/${encodeURIComponent(draftId)}/attachments`, { body: JSON.stringify(file) });
      }
      if (!options.draft) {
        sending = true;
        await graphRequest(ctx, 'POST', `/me/messages/${encodeURIComponent(draftId)}/send`);
      }
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      const code = error instanceof CliError ? error.exitCode : 1;
      throw new CliError(`Draft ${draftId} was created but ${sending ? 'sending was not confirmed (check Sent Items before retrying)' : (htmlBody === undefined ? 'not all files were attached and nothing was sent' : 'not all of the body and files were added and nothing was sent')}: ${detail}`, code);
    }
  }
  if (options.draft) process.stdout.write(draftId ? `Draft created: ${draftId}\n` : 'Draft created.\n');
  else process.stdout.write(`${command === 'reply-all' ? 'Reply-all' : command[0]!.toUpperCase() + command.slice(1)} sent.\n`);
}
