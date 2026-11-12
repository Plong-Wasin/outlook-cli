import assert from 'node:assert/strict';
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;
before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url, OUTLOOK_TOKEN: 'test-token' });
});
after(async () => { await servers.close(); cli.cleanup(); });

it('refuses replies and forwards without --yes or --draft and sends no request', async () => {
  for (const command of ['reply', 'reply-all', 'forward']) {
    servers.graph.reset();
    const result = await cli.run([command, 'message-123', ...(command === 'forward' ? ['--to', 'person@example.test'] : []), '--body', 'Hello']);
    assert.equal(result.code, 2, command);
    assert.match(result.stderr, /--yes|--draft/);
    assert.equal(servers.graph.requests.length, 0, command);
  }
});

it('creates the matching Graph Draft for replies and forwards', async () => {
  servers.graph.setReply((_request, response) => {
    response.writeHead(201, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'draft-123' }));
  });
  for (const [command, endpoint] of [
    ['reply', 'createReply'],
    ['reply-all', 'createReplyAll'],
    ['forward', 'createForward'],
  ] as const) {
    servers.graph.reset();
    const result = await cli.run([command, 'message-123', ...(command === 'forward' ? ['--to', 'person@example.test'] : []), '--body', 'Hello', '--draft']);
    assert.equal(result.code, 0, result.stderr);
    assert.match(result.stdout, /Draft created: draft-123/);
    assert.deepEqual(servers.graph.requests.map(({ method, path, body }) => [method, path, JSON.parse(body)]), [
      ['POST', `/me/messages/message-123/${endpoint!}`, command === 'forward'
        ? { message: { toRecipients: [{ emailAddress: { address: 'person@example.test' } }] }, comment: 'Hello' }
        : { comment: 'Hello' }],
    ]);
  }
});

it('sends the matching reply or forward request when --yes is supplied', async () => {
  servers.graph.setReply((_request, response) => { response.writeHead(202).end(); });
  for (const [command, endpoint] of [
    ['reply', 'reply'],
    ['reply-all', 'replyAll'],
    ['forward', 'forward'],
  ] as const) {
    servers.graph.reset();
    const result = await cli.run([command, 'message-123', ...(command === 'forward' ? ['--to', 'person@example.test'] : []), '--body', 'Hello', '--yes']);
    assert.equal(result.code, 0, result.stderr);
    assert.deepEqual(servers.graph.requests.map(({ method, path, body }) => [method, path, JSON.parse(body)]), [
      ['POST', `/me/messages/message-123/${endpoint!}`, command === 'forward'
        ? { toRecipients: [{ emailAddress: { address: 'person@example.test' } }], comment: 'Hello' }
        : { comment: 'Hello' }],
    ]);
  }
});

it('sends attachments in the message part of a reply', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-attach-'));
  const file = join(dir, 'a.txt');
  writeFileSync(file, 'abc');
  try {
    servers.graph.reset();
    const result = await cli.run(['reply', 'message-123', '--body', 'Hello', '--attach', file, '--draft']);
    assert.equal(result.code, 0, result.stderr);
    assert.equal(servers.graph.requests[0]!.path, '/me/messages/message-123/createReply');
    assert.equal(JSON.parse(servers.graph.requests[0]!.body).message.attachments[0].name, 'a.txt');
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
