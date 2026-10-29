import pkg from '../package.json' with { type: 'json' };
import { HELP } from './help.js';
import { runAuth, selectedContext, type GlobalOptions } from './commands/auth-commands.js';
import { runConfig } from './commands/config-commands.js';
import { runReadCommand } from './commands/read-commands.js';
import { runApi } from './commands/api-commands.js';
import { runDraftCommand } from './commands/draft-commands.js';
import { runSendCommand } from './commands/send-commands.js';
import { runReplyForwardCommand } from './commands/reply-forward-commands.js';
import { runWriteCommand } from './commands/write-commands.js';
import { CliError, usage } from './utils/errors.js';

const READ_COMMANDS = ['me', 'folders', 'messages', 'message', 'contacts', 'users', 'groups', 'members'];

function globals(args: string[]): { options: GlobalOptions; rest: string[] } {
  const options: GlobalOptions = {}; const rest: string[] = [];
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]!;
    if (arg === '-p' || arg === '--profile' || arg === '--client-id' || arg === '--scopes') {
      const value = args[index + 1];
      if (!value || value.startsWith('-')) usage(`Option '${arg}' requires a value.`);
      if (arg === '--client-id') options.clientId = value; else if (arg === '--scopes') options.scopes = value; else options.profile = value;
      index += 1;
    } else if (arg.startsWith('--profile=')) options.profile = arg.slice('--profile='.length);
    else if (arg.startsWith('--client-id=')) options.clientId = arg.slice('--client-id='.length);
    else if (arg.startsWith('--scopes=')) options.scopes = arg.slice('--scopes='.length);
    else rest.push(arg);
  }
  return { options, rest };
}
async function main(args: string[]): Promise<void> {
  if (args.length === 1 && args[0] === '--version') {
    process.stdout.write(`outlook ${pkg.version}\n`); return;
  }
  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    process.stdout.write(`${HELP}\n`); return;
  }
  const { options, rest } = globals(args);
  const [command, ...commandArgs] = rest;
  const ctx = selectedContext(process.env, options);
  if (command === 'auth') return runAuth(ctx, commandArgs[0]);
  if (command === 'config') return runConfig(ctx, commandArgs);
  if (command === 'api') return runApi(ctx, commandArgs);
  if (command === 'draft') return runDraftCommand(ctx, commandArgs);
  if (command === 'send') return runSendCommand(ctx, commandArgs);
  if (command === 'reply' || command === 'reply-all' || command === 'forward') return runReplyForwardCommand(ctx, command, commandArgs);
  if (command === 'mark' || command === 'delete') return runWriteCommand(ctx, command, commandArgs);
  if (command && READ_COMMANDS.includes(command)) return runReadCommand(ctx, command, commandArgs);
  if (!command) usage("Run 'outlook --help' for usage.");
  throw new CliError(`Unknown command: ${command}. Run 'outlook --help' for usage.`, 2);
}
main(process.argv.slice(2)).catch((error: unknown) => {
  const code = error instanceof CliError ? error.exitCode : 1;
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`outlook: ${message}\n`);
  process.exitCode = code;
});
