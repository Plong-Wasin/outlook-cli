import { graphRequest, type Context } from './auth-commands.js';
import { usage } from '../utils/errors.js';

function table(headers: string[], rows: string[][]): string {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values: string[]) => values.map((value, index) => value.padEnd(widths[index]!)).join('  ').trimEnd();
  return `${line(headers)}\n${rows.map(line).join('\n')}${rows.length ? '\n' : ''}`;
}
export async function runReadCommand(ctx: Context, command: string, args: string[]): Promise<void> {
  if (command !== 'me') usage(`Unknown read command: ${command}`);
  const response = await graphRequest(ctx, 'GET', '/me');
  if (args.includes('--json')) { process.stdout.write(`${JSON.stringify(response, null, 2)}\n`); return; }
  const user = response as { displayName?: string; userPrincipalName?: string; mail?: string };
  process.stdout.write(table(['NAME', 'EMAIL'], [[user.displayName ?? 'Current user', user.mail ?? user.userPrincipalName ?? '']]));
}
