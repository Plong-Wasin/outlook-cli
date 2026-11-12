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

it('adds attachments to the Draft and sends it for reply, reply-all and forward with --yes', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-attach-'));
  const file = join(dir, 'a.txt');
  writeFileSync(file, 'abc');
  servers.graph.setReply((_request, response) => {
    response.writeHead(201, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'draft-9' }));
  });
  try {
    for (const [command, endpoint] of [['reply', 'createReply'], ['reply-all', 'createReplyAll'], ['forward', 'createForward']] as const) {
      servers.graph.reset();
      const result = await cli.run([command, 'message-123', ...(command === 'forward' ? ['--to', 'person@example.test'] : []), '--body', 'Hello', '--attach', file, '--yes']);
      assert.equal(result.code, 0, result.stderr);
      assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [
        ['POST', `/me/messages/message-123/${endpoint}`],
        ['POST', '/me/messages/draft-9/attachments'],
        ['POST', '/me/messages/draft-9/send'],
      ], command);
      assert.equal(JSON.parse(servers.graph.requests[1]!.body).name, 'a.txt');
    }
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

it('leaves the Draft unsent when --draft is used with --attach', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-attach-'));
  const file = join(dir, 'a.txt');
  writeFileSync(file, 'abc');
  servers.graph.setReply((_request, response) => {
    response.writeHead(201, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'draft-9' }));
  });
  try {
    servers.graph.reset();
    const result = await cli.run(['reply', 'message-123', '--body', 'Hello', '--attach', file, '--draft']);
    assert.equal(result.code, 0, result.stderr);
    assert.match(result.stdout, /Draft created: draft-9/);
    assert.deepEqual(servers.graph.requests.map(({ path }) => path), ['/me/messages/message-123/createReply', '/me/messages/draft-9/attachments']);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

it('names the created Draft when adding an attachment fails', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-attach-'));
  const file = join(dir, 'a.txt');
  writeFileSync(file, 'abc');
  servers.graph.setReply((request, response) => {
    if (request.path.endsWith('/attachments')) { response.writeHead(500, { 'Content-Type': 'application/json' }); response.end('{}'); return; }
    response.writeHead(201, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'draft-9' }));
  });
  try {
    servers.graph.reset();
    const result = await cli.run(['reply', 'message-123', '--body', 'Hello', '--attach', file, '--yes']);
    assert.equal(result.code, 1);
    assert.match(result.stderr, /Draft draft-9 was created but was not sent/);
    assert.ok(!servers.graph.requests.some(({ path }) => path.endsWith('/send')));
  } finally { rmSync(dir, { recursive: true, force: true }); }
});
