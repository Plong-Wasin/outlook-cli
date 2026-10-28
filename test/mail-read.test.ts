import assert from 'node:assert/strict';
import { after, before, beforeEach, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;
before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url, OUTLOOK_LOGIN_URL: servers.login.url, OUTLOOK_TOKEN: 'read-token' });
});
beforeEach(() => servers.graph.reset());
after(async () => { await servers.close(); cli.cleanup(); });

function reply(body: unknown): void {
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(body));
  });
}

function query(path: string): URLSearchParams {
  return new URL(`http://graph.test${path}`).searchParams;
}

it('lists folders with full ids and raw Graph JSON', async () => {
  const folders = { value: [{ id: 'folder-long-id-123456789', displayName: 'Archive' }] };
  reply(folders);
  const table = await cli.run(['folders']);
  assert.equal(table.code, 0, table.stderr);
  assert.match(table.stdout, /folder-long-id-123456789\s+Archive/);
  assert.equal(servers.graph.requests[0]?.path, '/me/mailFolders');

  const json = await cli.run(['folders', '--json']);
  assert.deepEqual(JSON.parse(json.stdout), folders);
});

it('lists newest Inbox Messages by default with 25 rows and full ids', async () => {
  reply({ value: [{ id: 'AAMk-full-message-id', subject: 'Newest', receivedDateTime: '2026-10-29T10:00:00Z' }] });
  const result = await cli.run(['messages']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /AAMk-full-message-id/);
  assert.match(result.stdout, /Newest/);
  const request = servers.graph.requests[0]?.path ?? '';
  assert.ok(request.startsWith('/me/mailFolders/inbox/messages?'));
  assert.equal(query(request).get('$top'), '25');
  assert.equal(query(request).get('$orderby'), 'receivedDateTime desc');
});

it('resolves a Mail folder name and sends each message filter to Graph', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/me/mailFolders') response.end(JSON.stringify({ value: [{ id: 'folder-archive-id', displayName: 'Archive' }] }));
    else response.end(JSON.stringify({ value: [] }));
  });
  const result = await cli.run(['messages', '--folder', 'Archive', '--unread', '--from', 'ada@example.test', '--since', '2026-10-01', '--search', 'budget', '--top', '7']);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(servers.graph.requests[0]?.path, '/me/mailFolders');
  const request = servers.graph.requests[1]?.path ?? '';
  assert.ok(request.startsWith('/me/mailFolders/folder-archive-id/messages?'));
  const params = query(request);
  assert.equal(params.get('$top'), '7');
  assert.equal(params.get('$orderby'), 'receivedDateTime desc');
  assert.equal(params.get('$filter'), "isRead eq false and from/emailAddress/address eq 'ada@example.test' and receivedDateTime ge 2026-10-01T00:00:00.000Z");
  assert.equal(params.get('$search'), '"budget"');
});

it('rejects unknown and ambiguous folder names without listing Messages', async () => {
  servers.graph.setReply((_request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [
      { id: 'folder-a', displayName: 'Archive' },
      { id: 'folder-b', displayName: 'Archive' },
    ] }));
  });
  const unknown = await cli.run(['messages', '--folder', 'Missing']);
  assert.equal(unknown.code, 2);
  assert.match(unknown.stderr, /Unknown Mail folder 'Missing'/);
  const ambiguous = await cli.run(['messages', '--folder', 'Archive']);
  assert.equal(ambiguous.code, 2);
  assert.match(ambiguous.stderr, /Ambiguous Mail folder 'Archive'.*folder-a.*folder-b/);
  assert.ok(servers.graph.requests.every(({ path }) => path === '/me/mailFolders'));
});

it('prints a continuation cursor to stderr and accepts it on the next request', async () => {
  const next = `${servers.graph.url}/me/mailFolders/inbox/messages?$skiptoken=next-page`;
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(request.path.includes('skiptoken')
      ? { value: [{ id: 'message-page-2', subject: 'Second' }] }
      : { value: [{ id: 'message-page-1', subject: 'First' }], '@odata.nextLink': next }));
  });
  const first = await cli.run(['messages']);
  assert.equal(first.code, 0, first.stderr);
  assert.match(first.stderr, /more available: --cursor/);
  assert.doesNotMatch(first.stdout, /cursor/);
  const cursor = first.stderr.match(/--cursor ([^ ]+)/)?.[1];
  assert.ok(cursor);
  servers.graph.reset();
  const second = await cli.run(['messages', '--cursor', cursor]);
  assert.equal(second.code, 0, second.stderr);
  assert.match(second.stdout, /message-page-2/);
  assert.match(servers.graph.requests[0]?.path ?? '', /\$skiptoken=next-page/);
});

it('stops --all at the 100 page cap and leaves the next cursor on stderr', async () => {
  let page = 0;
  servers.graph.setReply((_request, response) => {
    page += 1;
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ value: [{ id: `message-${page}`, subject: `Message ${page}` }], '@odata.nextLink': `${servers.graph.url}/me/mailFolders/inbox/messages?page=${page + 1}` }));
  });
  const result = await cli.run(['messages', '--all']);
  assert.equal(result.code, 0, result.stderr);
  assert.equal(servers.graph.requests.length, 100);
  assert.match(result.stdout, /message-100/);
  assert.match(result.stderr, /more available: --cursor/);
});
