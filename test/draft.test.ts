import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;
before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url, OUTLOOK_TOKEN: 'test-token' });
  servers.graph.setReply((_request, response) => {
    response.writeHead(201, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'draft-123' }));
  });
});
after(async () => { await servers.close(); cli.cleanup(); });

it('creates a Draft with recipients and an HTML body read from a file', async () => {
  const bodyPath = join(cli.home, 'body.html');
  writeFileSync(bodyPath, '<p>Hello</p>');
  const result = await cli.run(['draft', '--to', 'one@example.test,two@example.test', '--cc', 'copy@example.test', '--bcc', 'blind@example.test', '--subject', 'Hello', '--body', `@${bodyPath}`, '--html']);
  assert.equal(result.code, 0, result.stderr);
  const request = servers.graph.requests.at(-1)!;
  assert.equal(request.method, 'POST');
  assert.equal(request.path, '/me/messages');
  assert.deepEqual(JSON.parse(request.body), {
    toRecipients: [
      { emailAddress: { address: 'one@example.test' } },
      { emailAddress: { address: 'two@example.test' } },
    ],
    ccRecipients: [{ emailAddress: { address: 'copy@example.test' } }],
    bccRecipients: [{ emailAddress: { address: 'blind@example.test' } }],
    subject: 'Hello',
    body: { contentType: 'HTML', content: '<p>Hello</p>' },
  });
});

it('creates a text Draft from stdin and PATCHes only supplied fields on update', async () => {
  servers.graph.reset();
  const created = await cli.run(['draft', '--to', 'person@example.test', '--body', '-'], {}, 'From stdin');
  assert.equal(created.code, 0, created.stderr);
  assert.equal(servers.graph.requests[0]!.method, 'POST');
  assert.deepEqual(JSON.parse(servers.graph.requests[0]!.body), {
    toRecipients: [{ emailAddress: { address: 'person@example.test' } }],
    body: { contentType: 'Text', content: 'From stdin' },
  });

  const updated = await cli.run(['draft', '--update', 'draft-123', '--subject', 'Revised', '--body', 'New body']);
  assert.equal(updated.code, 0, updated.stderr);
  assert.equal(servers.graph.requests[1]!.method, 'PATCH');
  assert.equal(servers.graph.requests[1]!.path, '/me/messages/draft-123');
  assert.deepEqual(JSON.parse(servers.graph.requests[1]!.body), {
    subject: 'Revised',
    body: { contentType: 'Text', content: 'New body' },
  });
});

it('creates a Draft with inline attachments', async () => {
  const file = join(cli.home, 'letter.docx');
  writeFileSync(file, 'docx-bytes');
  const result = await cli.run(['draft', '--to', 'one@example.test', '--attach', file]);
  assert.equal(result.code, 0, result.stderr);
  const request = servers.graph.requests.at(-1)!;
  assert.equal(request.path, '/me/messages');
  assert.deepEqual(JSON.parse(request.body).attachments, [{
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: 'letter.docx',
    contentBytes: Buffer.from('docx-bytes').toString('base64'),
  }]);
});

it('PATCHes a Draft then POSTs each attachment on update', async () => {
  const file = join(cli.home, 'page.html');
  writeFileSync(file, '<b>x</b>');
  servers.graph.reset();
  const result = await cli.run(['draft', '--update', 'draft-123', '--attach', file]);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [
    ['PATCH', '/me/messages/draft-123'],
    ['POST', '/me/messages/draft-123/attachments'],
  ]);
  assert.equal(JSON.parse(servers.graph.requests[1]!.body).name, 'page.html');
});
