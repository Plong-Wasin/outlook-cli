import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';
import { messageId } from '../utils/message-id.js';
import { graphPage, validateContinuation } from '../utils/graph-page.js';
import { resolveMailFolderId } from '../utils/mail-folders.js';

interface Message {
  id?: string;
  subject?: string;
  receivedDateTime?: string;
  from?: { emailAddress?: { name?: string; address?: string } };
  isRead?: boolean;
}

interface Folder {
  id?: string;
  displayName?: string;
  totalItemCount?: number;
  unreadItemCount?: number;
}

interface Collection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
  [key: string]: unknown;
}

interface ReadCollection<T> {
  response: Collection<T>;
  items: T[];
  nextLink?: string;
}

interface DirectoryPerson {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  emailAddresses?: Array<{ address?: string }>;
}

const ALL_PAGE_CAP = 100;

function table(headers: string[], rows: string[][]): string {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values: string[]) => values.map((value, index) => value.padEnd(widths[index]!)).join('  ').trimEnd();
  return `${line(headers)}\n${rows.map(line).join('\n')}${rows.length ? '\n' : ''}`;
}

export async function runReadCommand(ctx: Context, command: string, args: string[]): Promise<void> {
  if (command === 'message') return readMessage(ctx, args);
  if (['contacts', 'users', 'groups', 'members'].includes(command)) return readDirectory(ctx, command, args);
  let response: unknown;
  let rows: unknown[] | undefined;
  let nextLink: string | undefined;
  if (command === 'me') response = await graphRequest(ctx, 'GET', '/me');
  else if (command === 'folders') {
    const result = await readCollection<Folder>(ctx, '/me/mailFolders', args);
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else if (command === 'messages') {
    const folder = option(args, '--folder');
    const folderId = folder ? await resolveMailFolderId(ctx, folder) : 'inbox';
    const result = await readCollection<Message>(ctx, `/me/mailFolders/${encodeURIComponent(folderId)}/messages`, args, messageQuery(args));
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else usage(`Unknown read command: ${command}`);

  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
    printContinuation(nextLink);
    return;
  }
  if (command === 'me') {
    const user = response as { displayName?: string; userPrincipalName?: string; mail?: string };
    process.stdout.write(table(['NAME', 'EMAIL'], [[user.displayName ?? 'Current user', user.mail ?? user.userPrincipalName ?? '']]));
    return;
  }
  if (command === 'folders') {
    const folders = (rows as Folder[] | undefined) ?? (response as Collection<Folder>).value ?? [];
    process.stdout.write(table(['ID', 'NAME', 'TOTAL', 'UNREAD'], folders.map((folder) => [folder.id ?? '', folder.displayName ?? '', String(folder.totalItemCount ?? ''), String(folder.unreadItemCount ?? '')])));
    printContinuation(nextLink);
    return;
  }
  const messages = (rows as Message[] | undefined) ?? (response as Collection<Message>).value ?? [];
  process.stdout.write(table(['ID', 'RECEIVED', 'FROM', 'SUBJECT', 'READ'], messages.map((message) => [
    message.id ?? '',
    message.receivedDateTime ?? '',
    message.from?.emailAddress?.name || message.from?.emailAddress?.address || '',
    message.subject ?? '',
    message.isRead === undefined ? '' : message.isRead ? 'yes' : 'no',
  ])));
  printContinuation(nextLink);
}

async function readCollection<T>(ctx: Context, path: string, args: string[], query?: URLSearchParams, headers?: Record<string, string>): Promise<ReadCollection<T>> {
  const token = option(args, '--cursor');
  let url = token ? decodeCursor(token, ctx) : withQuery(path, query);
  const all = args.includes('--all');
  const pages: Collection<T>[] = [];
  let nextLink: string | undefined;
  for (let page = 0; page < (all ? ALL_PAGE_CAP : 1); page += 1) {
    const result = await graphPage(ctx, url, headers) as Collection<T>;
    pages.push(result);
    nextLink = result['@odata.nextLink'];
    if (!all || !nextLink) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  const first = pages[0] ?? { value: [] };
  const items = pages.flatMap((page) => page.value ?? []);
  const response = all ? { ...first, value: items, ...(nextLink ? { '@odata.nextLink': nextLink } : {}) } : first;
  if (all && !nextLink) delete response['@odata.nextLink'];
  return { response, items, nextLink };
}

function messageQuery(args: string[]): URLSearchParams {
  const query = new URLSearchParams();
  const top = option(args, '--top') ?? '25';
  if (!/^\d+$/.test(top) || Number(top) < 1 || Number(top) > 1000) usage("Option '--top' must be an integer from 1 to 1000.");
  query.set('$top', top);
  const search = option(args, '--search');
  const unread = args.includes('--unread');
  const from = option(args, '--from');
  const since = option(args, '--since');
  if (search) {
    // Graph rejects $search combined with $filter or $orderby; results come back ranked by relevance.
    if (unread || from || since) usage("Option '--search' cannot be combined with '--unread', '--from', or '--since'.");
    query.set('$search', `"${search.replaceAll('"', '\\"')}"`);
    return query;
  }
  query.set('$orderby', 'receivedDateTime desc');
  let sinceDate = new Date(0);
  if (since) {
    sinceDate = new Date(since);
    if (Number.isNaN(sinceDate.getTime())) usage("Option '--since' must be a valid date.");
  }
  // Graph requires the $orderby property to lead the $filter, or it fails with InefficientFilter.
  const filters = [`receivedDateTime ge ${sinceDate.toISOString()}`];
  if (unread) filters.push('isRead eq false');
  if (from) filters.push(`from/emailAddress/address eq '${from.replaceAll("'", "''")}'`);
  query.set('$filter', filters.join(' and '));
  return query;
}

function withQuery(path: string, query?: URLSearchParams): string {
  return query && [...query].length ? `${path}?${query.toString()}` : path;
}

function decodeCursor(token: string, ctx: Context): string {
  let link: string;
  try { link = Buffer.from(token, 'base64url').toString('utf8'); }
  catch { usage('Invalid cursor. Copy the cursor from the previous command output.'); }
  validateContinuation(link, ctx);
  return link;
}

function printContinuation(nextLink: string | undefined): void {
  if (nextLink) process.stderr.write(`more available: --cursor ${Buffer.from(nextLink).toString('base64url')} (or --all)\n`);
}

function option(args: string[], name: string): string | undefined {
  const index = args.findIndex((arg) => arg === name || arg.startsWith(`${name}=`));
  if (index < 0) return undefined;
  const arg = args[index]!;
  const value = arg === name ? args[index + 1] : arg.slice(name.length + 1);
  if (!value || value.startsWith('-')) usage(`Option '${name}' requires a value.`);
  return value;
}

async function readMessage(ctx: Context, args: string[]): Promise<void> {
  if (args.some((arg) => arg.startsWith('--') && arg !== '--html' && arg !== '--json')) {
    usage("Use 'outlook message <id|url> [--html] [--json]'.");
  }
  const positional = args.filter((arg) => !arg.startsWith('--'));
  if (positional.length !== 1) usage("Use 'outlook message <id|url> [--html] [--json]'.");
  const id = messageId(positional[0]!);
  const html = args.includes('--html');
  const response = await graphRequest(ctx, 'GET', `/me/messages/${encodeURIComponent(id)}`, {
    ...(html ? {} : { headers: { Prefer: 'outlook.body-content-type="text"' } }),
  });
  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}\n`);
    return;
  }
  const message = response as { body?: { content?: string } };
  process.stdout.write(`${message.body?.content ?? ''}\n`);
}

async function readDirectory(ctx: Context, command: string, args: string[]): Promise<void> {
  let path: string;
  if (command === 'contacts') path = '/me/contacts';
  else if (command === 'users') path = '/users';
  else if (command === 'groups') path = '/groups';
  else {
    const group = args[0];
    if (!group || group.startsWith('-')) usage("A group id is required. Usage: 'outlook members <group>'.");
    path = `/groups/${encodeURIComponent(group)}/members`;
  }
  const top = option(args, '--top') ?? '25';
  if (!/^\d+$/.test(top) || Number(top) < 1 || Number(top) > 999) usage("Option '--top' must be an integer from 1 to 999.");
  const query = new URLSearchParams({ $top: top });
  const search = option(args, '--search');
  if (search) {
    if (command !== 'users') usage("Option '--search' is only supported by 'users'.");
    query.set('$search', `"displayName:${search.replaceAll('"', '\\"')}"`);
  }
  // Directory $search needs advanced-query consistency.
  const result = await readCollection<DirectoryPerson>(ctx, path, args, query, search ? { ConsistencyLevel: 'eventual' } : undefined);
  if (args.includes('--json')) {
    process.stdout.write(`${JSON.stringify(result.response, null, 2)}\n`);
    printContinuation(result.nextLink);
    return;
  }
  process.stdout.write(table(['ID', 'NAME', 'EMAIL'], result.items.map((item) => [
    item.id ?? '',
    item.displayName ?? '',
    item.mail ?? item.userPrincipalName ?? item.emailAddresses?.[0]?.address ?? '',
  ])));
  printContinuation(result.nextLink);
}
