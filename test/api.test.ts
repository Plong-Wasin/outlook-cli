import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
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

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('passes raw API query pairs, body, and headers through to Graph', async () => {
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end('{"ok":true}');
  });
  const result = await cli.run([
    'api', 'POST', '/me/messages', '$select=id,subject', 'filter=subject eq \'Hello\'',
    '--data', '{ "subject" : "Hello" }', '--header', 'Prefer: outlook.body-content-type="text"',
    '--header', 'Content-Type: application/merge-patch+json',
  ]);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, '{"ok":true}\n');
  assert.equal(servers.graph.requests[0]?.method, 'POST');
  assert.equal(servers.graph.requests[0]?.path, '/me/messages?%24select=id%2Csubject&filter=subject+eq+%27Hello%27');
  assert.equal(servers.graph.requests[0]?.body, '{"subject":"Hello"}');
  assert.equal(servers.graph.requests[0]?.headers.prefer, 'outlook.body-content-type="text"');
  assert.equal(servers.graph.requests[0]?.headers['content-type'], 'application/merge-patch+json');
  assert.equal(servers.graph.requests[0]?.headers.authorization, 'Bearer test-token');
});

it('reads API JSON bodies from a file and stdin', async () => {
  servers.graph.reset();
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(request.body);
  });
  const filePath = join(cli.home, 'payload.json');
  writeFileSync(filePath, '{"subject":"From file"}');
  const fromFile = await cli.run(['api', 'POST', '/me/messages', '--data', `@${filePath}`]);
  assert.equal(fromFile.code, 0, fromFile.stderr);
  assert.equal(servers.graph.requests[0]?.body, '{"subject":"From file"}');

  servers.graph.reset();
  const fromStdin = await cli.run(['api', 'POST', '/me/messages', '--data', '-'], {}, '{"subject":"From stdin"}');
  assert.equal(fromStdin.code, 0, fromStdin.stderr);
  assert.equal(servers.graph.requests[0]?.body, '{"subject":"From stdin"}');
});

it('follows continuation links with --all and combines collection values', async () => {
  servers.graph.reset();
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(request.path === '/me/messages?%24top=1'
      ? { value: [{ id: 'first' }], '@odata.nextLink': `${servers.graph.url}/me/messages?$skiptoken=next` }
      : { value: [{ id: 'second' }] }));
  });
  const result = await cli.run(['api', 'GET', '/me/messages', '$top=1', '--all']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, '{"value":[{"id":"first"},{"id":"second"}]}\n');
  assert.equal(servers.graph.requests.length, 2);
  assert.equal(servers.graph.requests[1]?.path, '/me/messages?$skiptoken=next');
});

it('stops --all at the 50-page cap', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({
      value: [{ id: String(servers.graph.requests.length) }],
      '@odata.nextLink': `${servers.graph.url}/me/messages?$skiptoken=next`,
    }));
  });
  const result = await cli.run(['api', 'GET', '/me/messages', '--all']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(servers.graph.requests.length, 50);
  assert.equal(JSON.parse(result.stdout).value.length, 50);
});

it('treats an empty 204 response as success', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => { response.writeHead(204).end(); });
  const result = await cli.run(['api', 'DELETE', '/me/messages/message-1']);

  assert.equal(result.code, 0, result.stderr);
  assert.equal(result.stdout, '');
  assert.equal(result.stderr, '');
});
