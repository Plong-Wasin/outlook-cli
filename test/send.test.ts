import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;
before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url, OUTLOOK_TOKEN: 'test-token' });
  servers.graph.setReply((_request, response) => {
    response.writeHead(202);
    response.end();
  });
});
after(async () => { await servers.close(); cli.cleanup(); });

it('prints recipients and subject and refuses to send without --yes', async () => {
  servers.graph.reset();
  const result = await cli.run(['send', '--to', 'one@example.test,two@example.test', '--subject', 'Hello', '--body', 'Message body']);
  assert.equal(result.code, 2);
  assert.match(result.stdout, /one@example\.test/);
  assert.match(result.stdout, /two@example\.test/);
  assert.match(result.stdout, /Subject: Hello/);
  assert.match(result.stderr, /--yes/);
  assert.equal(servers.graph.requests.length, 0);
});

it('sends a new Message with the expected Graph request when --yes is supplied', async () => {
  servers.graph.reset();
  const result = await cli.run(['send', '--to', 'person@example.test', '--cc', 'copy@example.test', '--subject', 'Hello', '--body', 'Message body', '--yes']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [['POST', '/me/sendMail']]);
  assert.deepEqual(JSON.parse(servers.graph.requests[0]!.body), {
    message: {
      toRecipients: [{ emailAddress: { address: 'person@example.test' } }],
      ccRecipients: [{ emailAddress: { address: 'copy@example.test' } }],
      subject: 'Hello',
      body: { contentType: 'Text', content: 'Message body' },
    },
    saveToSentItems: true,
  });
});

it('sends a saved Draft by id when --yes is supplied', async () => {
  servers.graph.reset();
  const result = await cli.run(['send', '--draft', 'draft-123', '--yes']);
  assert.equal(result.code, 0, result.stderr);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [['POST', '/me/messages/draft-123/send']]);
});

it('shows saved Draft recipients and subject without sending when --yes is omitted', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({
      toRecipients: [{ emailAddress: { address: 'person@example.test' } }],
      subject: 'Saved subject',
    }));
  });
  const result = await cli.run(['send', '--draft', 'draft-123']);
  assert.equal(result.code, 2);
  assert.match(result.stdout, /person@example\.test/);
  assert.match(result.stdout, /Subject: Saved subject/);
  assert.match(result.stderr, /--yes/);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [['GET', '/me/messages/draft-123?%24select=toRecipients%2CccRecipients%2CbccRecipients%2Csubject']]);
});
