import assert from 'node:assert/strict';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url, OUTLOOK_LOGIN_URL: servers.login.url });
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ id: 'abc123', body: { contentType: 'text', content: 'Hello' } }));
  });
});

after(async () => { await servers.close(); cli.cleanup(); });

it('reads a Message as raw JSON and requests its plain text body', async () => {
  servers.graph.reset();
  const result = await cli.run(['message', 'abc123', '--json'], { OUTLOOK_TOKEN: 'test-token' });
  assert.equal(result.code, 0);
  assert.deepEqual(JSON.parse(result.stdout), { id: 'abc123', body: { contentType: 'text', content: 'Hello' } });
  assert.equal(servers.graph.requests[0]?.method, 'GET');
  assert.equal(servers.graph.requests[0]?.path, '/me/messages/abc123');
  assert.equal(servers.graph.requests[0]?.headers.prefer, 'outlook.body-content-type="text"');
});

it('resolves an Outlook URL to the same Message request and returns original HTML when requested', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ body: { contentType: 'html', content: '<p>Hello</p>' } }));
  });
  const result = await cli.run(['message', 'https://outlook.office.com/mail/inbox/id/abc123', '--html'], { OUTLOOK_TOKEN: 'test-token' });
  assert.equal(result.code, 0);
  assert.equal(result.stdout, '<p>Hello</p>\n');
  assert.equal(servers.graph.requests[0]?.path, '/me/messages/abc123');
  assert.equal(servers.graph.requests[0]?.headers.prefer, undefined);
  const byId = await cli.run(['message', 'abc123', '--html'], { OUTLOOK_TOKEN: 'test-token' });
  assert.equal(byId.code, 0);
  assert.deepEqual(
    [servers.graph.requests[0]?.method, servers.graph.requests[0]?.path, servers.graph.requests[0]?.headers.prefer],
    [servers.graph.requests[1]?.method, servers.graph.requests[1]?.path, servers.graph.requests[1]?.headers.prefer],
  );
});

it('prints the plain text body by default', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ body: { contentType: 'text', content: 'Plain message' } }));
  });
  const result = await cli.run(['message', 'abc123'], { OUTLOOK_TOKEN: 'test-token' });
  assert.equal(result.code, 0);
  assert.equal(result.stdout, 'Plain message\n');
  assert.equal(servers.graph.requests[0]?.headers.prefer, 'outlook.body-content-type="text"');
});

it('rejects malformed Message ids as usage errors without making a request', async () => {
  servers.graph.reset();
  const result = await cli.run(['message', 'abc!'], { OUTLOOK_TOKEN: 'test-token' });
  assert.equal(result.code, 2);
  assert.match(result.stderr, /Invalid Message id/);
  assert.equal(servers.graph.requests.length, 0);
});
