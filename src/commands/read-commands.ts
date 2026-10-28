import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';

function table(headers: string[], rows: string[][]): string {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values: string[]) => values.map((value, index) => value.padEnd(widths[index]!)).join('  ').trimEnd();
  return `${line(headers)}\n${rows.map(line).join('\n')}${rows.length ? '\n' : ''}`;
}
export async function runReadCommand(ctx: Context, command: string, args: string[]): Promise<void> {
  if (command === 'message') return readMessage(ctx, args);
  if (command !== 'me') usage(`Unknown read command: ${command}`);
  const response = await graphRequest(ctx, 'GET', '/me');
  if (args.includes('--json')) { process.stdout.write(`${JSON.stringify(response, null, 2)}\n`); return; }
  const user = response as { displayName?: string; userPrincipalName?: string; mail?: string };
  process.stdout.write(table(['NAME', 'EMAIL'], [[user.displayName ?? 'Current user', user.mail ?? user.userPrincipalName ?? '']]));
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

function messageId(input: string): string {
  let id = input;
  if (/^https?:\/\//i.test(input)) {
    let url: URL;
    try { url = new URL(input); }
    catch { usage('Invalid Outlook on the web URL. Copy the Message URL and try again.'); }
    if (!/(^|\.)outlook\.(office\.com|office365\.com|live\.com)$/i.test(url.hostname)) {
      usage('Expected an Outlook on the web URL. Copy the Message URL and try again.');
    }
    const pathParts = url.pathname.split('/').filter(Boolean);
    const idIndex = pathParts.findIndex((part) => part.toLowerCase() === 'id');
    const deeplinkIndex = pathParts.findIndex((part, index) => part.toLowerCase() === 'read' && pathParts[index - 1]?.toLowerCase() === 'deeplink');
    id = (idIndex >= 0 ? pathParts[idIndex + 1] : undefined)
      ?? (deeplinkIndex >= 0 ? pathParts[deeplinkIndex + 1] : undefined)
      ?? url.searchParams.get('ItemID')
      ?? url.searchParams.get('itemid')
      ?? '';
    try { id = decodeURIComponent(id); }
    catch { usage('Invalid Message id in Outlook URL.'); }
  }
  if (!/^[A-Za-z0-9_+\/=\-]+$/.test(id)) usage('Invalid Message id. Use a Graph id or Outlook on the web URL.');
  return id;
}
