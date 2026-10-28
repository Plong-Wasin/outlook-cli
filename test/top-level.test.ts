import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';
let servers: FakeServers; let cli: Cli;
before(async () => { servers = await startFakeServers(); cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url, OUTLOOK_LOGIN_URL: servers.login.url }); });
after(async () => { await servers.close(); cli.cleanup(); });
it('prints help without credentials and reports unknown commands as usage errors', async () => {
  const help = await cli.run(['--help']); assert.equal(help.code, 0); assert.match(help.stdout, /auth login \| check \| logout/);
  for (const item of ['config show', 'me', '--profile', '--client-id', 'OUTLOOK_TOKEN']) assert.ok(help.stdout.includes(item));
  const unknown = await cli.run(['nope']); assert.equal(unknown.code, 2); assert.match(unknown.stderr, /^outlook: /);
});
