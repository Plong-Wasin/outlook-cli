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
    assert.match(result.stderr, /Draft draft-9 was created but not all files were attached and nothing was sent/);
    assert.ok(!servers.graph.requests.some(({ path }) => path.endsWith('/send')));
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

it('says sending was not confirmed when /send fails after the files are attached', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'outlook-attach-'));
  const file = join(dir, 'a.txt');
  writeFileSync(file, 'abc');
  servers.graph.setReply((request, response) => {
    if (request.path.endsWith('/send')) { response.writeHead(500, { 'Content-Type': 'application/json' }); response.end('{}'); return; }
    response.writeHead(201, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'draft-9' }));
  });
  try {
    servers.graph.reset();
    const result = await cli.run(['reply', 'message-123', '--body', 'Hello', '--attach', file, '--yes']);
    assert.equal(result.code, 1);
    assert.match(result.stderr, /Draft draft-9 was created but sending was not confirmed/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

const QUOTED_DRAFT = '<html><head></head><body dir="ltr"><hr><div id="divRplyFwdMsg">quoted original</div></body></html>';

function multilineDraftServer(): void {
  servers.graph.setReply((request, response) => {
    response.writeHead(request.method === 'POST' && request.path.endsWith('/send') ? 202 : 200, { 'Content-Type': 'application/json' });
    if (request.method === 'GET') response.end(JSON.stringify({ id: 'draft-5', body: { contentType: 'html', content: QUOTED_DRAFT } }));
    else if (request.method === 'POST' && /\/create(Reply|ReplyAll|Forward)$/.test(request.path)) response.end(JSON.stringify({ id: 'draft-5' }));
    else response.end('{}');
  });
}

it('keeps line breaks of a multi-line --body by writing HTML into the Draft above the quoted message', async () => {
  for (const [command, endpoint] of [['reply', 'createReply'], ['reply-all', 'createReplyAll'], ['forward', 'createForward']] as const) {
    multilineDraftServer();
    servers.graph.reset();
    const result = await cli.run([command, 'message-123', ...(command === 'forward' ? ['--to', 'person@example.test'] : []), '--body', 'Line 1\nLine 2 & <b>\n\nLine 4', '--draft']);
    assert.equal(result.code, 0, result.stderr);
    assert.match(result.stdout, /Draft created: draft-5/);
    const requests = servers.graph.requests;
    assert.deepEqual(requests.map(({ method, path }) => [method, path.split('?')[0]]), [
      ['POST', `/me/messages/message-123/${endpoint}`],
      ['GET', '/me/messages/draft-5'],
      ['PATCH', '/me/messages/draft-5'],
    ], command);
    assert.deepEqual(JSON.parse(requests[0]!.body), command === 'forward' ? { message: { toRecipients: [{ emailAddress: { address: 'person@example.test' } }] } } : {}, command);
    const patched = JSON.parse(requests[2]!.body).body;
    assert.equal(patched.contentType, 'html');
    assert.match(patched.content, /<body dir="ltr">.*Line 1<br>\s*Line 2 &amp; &lt;b&gt;<br>\s*<br>\s*Line 4.*<hr>/s, command);
    assert.match(patched.content, /quoted original/, command);
  }
});

it('sends the patched Draft when a multi-line --body is used with --yes', async () => {
  multilineDraftServer();
  servers.graph.reset();
  const result = await cli.run(['reply', 'message-123', '--body', 'A\nB', '--yes']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /Reply sent\./);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path.split('?')[0]]), [
    ['POST', '/me/messages/message-123/createReply'],
    ['GET', '/me/messages/draft-5'],
    ['PATCH', '/me/messages/draft-5'],
    ['POST', '/me/messages/draft-5/send'],
  ]);
});

it('uses --body as raw HTML with --html', async () => {
  multilineDraftServer();
  servers.graph.reset();
  const result = await cli.run(['reply-all', 'message-123', '--html', '--body', '<p>One</p><p>Two & three</p>', '--draft']);
  assert.equal(result.code, 0, result.stderr);
  const patched = JSON.parse(servers.graph.requests.find(({ method }) => method === 'PATCH')!.body).body;
  assert.match(patched.content, /<body dir="ltr"><p>One<\/p><p>Two & three<\/p><hr>/);
});

it('still requires --body when only --html is given', async () => {
  servers.graph.reset();
  const result = await cli.run(['reply', 'message-123', '--html', '--draft']);
  assert.equal(result.code, 2);
  assert.equal(servers.graph.requests.length, 0);
});

it('names the Draft when patching the multi-line body fails and sends nothing', async () => {
  servers.graph.setReply((request, response) => {
    if (request.method === 'PATCH') { response.writeHead(500, { 'Content-Type': 'application/json' }); response.end('{}'); return; }
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'draft-5', body: { contentType: 'html', content: QUOTED_DRAFT } }));
  });
  servers.graph.reset();
  const result = await cli.run(['reply', 'message-123', '--body', 'A\nB', '--yes']);
  assert.equal(result.code, 1);
  assert.match(result.stderr, /Draft draft-5 was created but .*nothing was sent/);
  assert.ok(!servers.graph.requests.some(({ path }) => path.endsWith('/send')));
});

it('keeps the one-request path for a single-line body that ends with a newline', async () => {
  servers.graph.setReply((_request, response) => { response.writeHead(202).end(); });
  servers.graph.reset();
  const dir = mkdtempSync(join(tmpdir(), 'outlook-body-'));
  const file = join(dir, 'body.txt');
  writeFileSync(file, 'Hello\n');
  try {
    const result = await cli.run(['reply', 'message-123', '--body', `@${file}`, '--yes']);
    assert.equal(result.code, 0, result.stderr);
    assert.deepEqual(servers.graph.requests.map(({ method, path, body }) => [method, path, JSON.parse(body)]), [
      ['POST', '/me/messages/message-123/reply', { comment: 'Hello' }],
    ]);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

it('asks Graph for the Draft body as HTML', async () => {
  multilineDraftServer();
  servers.graph.reset();
  const result = await cli.run(['reply', 'message-123', '--body', 'A\nB', '--draft']);
  assert.equal(result.code, 0, result.stderr);
  const get = servers.graph.requests.find(({ method }) => method === 'GET')!;
  assert.equal(get.headers.prefer, 'outlook.body-content-type="html"');
});
