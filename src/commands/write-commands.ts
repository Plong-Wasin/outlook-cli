import { graphRequest, type Context } from './auth-commands.js';
import { CliError, usage } from '../utils/errors.js';
import { messageId } from '../utils/message-id.js';

interface WriteOptions {
  ids: string[];
  yes: boolean;
  isRead?: boolean;
}

export async function runWriteCommand(ctx: Context, command: 'mark' | 'delete', args: string[]): Promise<void> {
  const options = parseOptions(command, args);
  const ids = options.ids.map(messageId);
  if (!ids.length) usage(`Usage: outlook ${command} <id...>${command === 'mark' ? ' --read|--unread' : ' --yes'}`);
  if (command === 'mark' && options.isRead === undefined) usage("Choose either '--read' or '--unread'.");
  if (command === 'delete' && !options.yes) {
    process.stdout.write(`Would delete Messages:\n${ids.map((id) => `#${id}`).join('\n')}\n`);
    throw new CliError("Refusing to delete Messages without '--yes'.", 2);
  }

  const failures: string[] = [];
  for (const id of ids) {
    try {
      if (command === 'delete') await graphRequest(ctx, 'DELETE', `/me/messages/${encodeURIComponent(id)}`);
      else await graphRequest(ctx, 'PATCH', `/me/messages/${encodeURIComponent(id)}`, { body: JSON.stringify({ isRead: options.isRead }) });
      process.stdout.write(`${command === 'delete' ? 'Deleted' : `Marked ${options.isRead ? 'read' : 'unread'}`} #${id}\n`);
    } catch (error) {
      failures.push(id);
      process.stderr.write(`#${id}: ${error instanceof Error ? error.message : String(error)}\n`);
    }
  }
  if (failures.length) {
    process.stderr.write(`${failures.length} Message${failures.length === 1 ? '' : 's'} failed.\n`);
    process.exitCode = 1;
  }
}

function parseOptions(command: 'mark' | 'delete', args: string[]): WriteOptions {
  const options: WriteOptions = { ids: [], yes: false };
  for (const arg of args) {
    if (arg === '--yes') { options.yes = true; continue; }
    if (arg === '--json') usage("Option '--json' is only supported by read commands.");
    if (command === 'mark' && (arg === '--read' || arg === '--unread')) {
      if (options.isRead !== undefined) usage("Choose either '--read' or '--unread', not both.");
      options.isRead = arg === '--read';
      continue;
    }
    if (arg.startsWith('-')) usage(`Unknown ${command} option: ${arg}.`);
    options.ids.push(arg);
  }
  return options;
}
