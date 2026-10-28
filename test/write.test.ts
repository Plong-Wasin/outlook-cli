import assert from 'node:assert/strict';
import { after, before, beforeEach, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;
before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url, OUTLOOK_TOKEN: 'write-token' });
});
beforeEach(() => servers.graph.reset());
after(async () => { await servers.close(); cli.cleanup(); });

it('previews deletion without --yes and sends no request', async () => {
  const result = await cli.run(['delete', 'message-one', 'message-two']);
  assert.equal(result.code, 2);
  assert.match(result.stdout, /message-one/);
  assert.match(result.stdout, /message-two/);
  assert.equal(servers.graph.requests.length, 0);
});

it('marks Messages read or unread with one PATCH per id', async () => {
  servers.graph.setReply((_request, response) => { response.writeHead(204).end(); });
  const result = await cli.run(['mark', 'message-one', 'message-two', '--unread']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.map(({ method, path, body }) => [method, path, JSON.parse(body)]), [
    ['PATCH', '/me/messages/message-one', { isRead: false }],
    ['PATCH', '/me/messages/message-two', { isRead: false }],
  ]);

  servers.graph.reset();
  const read = await cli.run(['mark', 'message-three', '--read']);
  assert.equal(read.code, 0, read.stderr);
  assert.deepEqual(JSON.parse(servers.graph.requests[0]!.body), { isRead: true });
});

it('continues bulk deletion after a failure and reports the failed id and count', async () => {
  servers.graph.setReply((request, response) => {
    if (request.path.endsWith('message-two')) {
      response.writeHead(404, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify({ error: { message: 'Message not found' } }));
    } else response.writeHead(204).end();
  });
  const result = await cli.run(['delete', 'message-one', 'message-two', 'message-three', '--yes']);
  assert.equal(result.code, 1);
  assert.match(result.stdout, /Deleted #message-one/);
  assert.match(result.stdout, /Deleted #message-three/);
  assert.match(result.stderr, /#message-two: .*Message not found/);
  assert.match(result.stderr, /1 Message failed/);
  assert.equal(servers.graph.requests.length, 3);
});

it('rejects --json for every write command without sending a request', async () => {
  for (const args of [
    ['mark', 'message-one', '--read', '--json'],
    ['delete', 'message-one', '--yes', '--json'],
  ]) {
    const result = await cli.run(args);
    assert.equal(result.code, 2, args.join(' '));
    assert.match(result.stderr, /--json.*read commands/);
  }
  assert.equal(servers.graph.requests.length, 0);
});

it('marks without --yes because the change is reversible', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => { response.writeHead(200, { 'Content-Type': 'application/json' }); response.end('{}'); });
  const result = await cli.run(['mark', 'message-one', '--read']);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(servers.graph.requests.length, 1);
  assert.equal(servers.graph.requests[0]!.method, 'PATCH');
});

it('rejects --yes for mark and prints the intended action when reply is refused', async () => {
  servers.graph.reset();
  const mark = await cli.run(['mark', 'message-one', '--read', '--yes']);
  assert.equal(mark.code, 2);
  const reply = await cli.run(['reply', 'message-one', '--body', 'hi']);
  assert.equal(reply.code, 2);
  assert.match(reply.stdout, /Would reply Message #message-one/);
  assert.equal(servers.graph.requests.length, 0);
});

it('rejects --json for send, draft, reply, reply-all, and forward without sending', async () => {
  servers.graph.reset();
  for (const args of [['send', '--to', 'a@example.test', '--yes', '--json'], ['draft', '--to', 'a@example.test', '--json'], ['reply', 'message-one', '--body', 'x', '--yes', '--json'], ['reply-all', 'message-one', '--body', 'x', '--yes', '--json'], ['forward', 'message-one', '--to', 'a@example.test', '--body', 'x', '--yes', '--json']]) {
    const result = await cli.run(args);
    assert.equal(result.code, 2, args[0]);
  }
  assert.equal(servers.graph.requests.length, 0);
});
