import assert from 'node:assert/strict';
import { writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;

before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_GRAPH_URL: servers.graph.url });
  writeFileSync(`${cli.home}/config.json`, JSON.stringify({ profiles: { default: { accessToken: 'read-token' } } }));
});

after(async () => {
  await servers.close();
  cli.cleanup();
});

it('lists Contacts, Users, Groups, and group Members from their Graph endpoints', async () => {
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    const fixtures: Record<string, unknown> = {
      '/me/contacts': { value: [{ id: 'contact-1', displayName: 'Ada Lovelace', emailAddresses: [{ address: 'ada@example.test' }] }] },
      '/users': { value: [{ id: 'user-1', displayName: 'Grace Hopper', mail: 'grace@example.test' }] },
      '/groups': { value: [{ id: 'group-1', displayName: 'Engineering' }] },
      '/groups/group-1/members': { value: [{ id: 'user-2', displayName: 'Katherine Johnson', mail: 'kat@example.test' }] },
    };
    response.end(JSON.stringify(fixtures[request.path.split('?')[0]!]));
  });
  servers.graph.reset();

  const contacts = await cli.run(['contacts']);
  assert.equal(contacts.code, 0, contacts.stderr);
  assert.match(contacts.stdout, /contact-1.*Ada Lovelace.*ada@example\.test/);
  const users = await cli.run(['users']);
  assert.equal(users.code, 0, users.stderr);
  assert.match(users.stdout, /user-1.*Grace Hopper.*grace@example\.test/);
  const groups = await cli.run(['groups']);
  assert.equal(groups.code, 0, groups.stderr);
  assert.match(groups.stdout, /group-1.*Engineering/);
  const members = await cli.run(['members', 'group-1']);
  assert.equal(members.code, 0, members.stderr);
  assert.match(members.stdout, /user-2.*Katherine Johnson.*kat@example\.test/);
  assert.deepEqual(servers.graph.requests.map(({ method, path }) => [method, path]), [
    ['GET', '/me/contacts?%24top=25'], ['GET', '/users?%24top=25'], ['GET', '/groups?%24top=25'], ['GET', '/groups/group-1/members?%24top=25'],
  ]);
});

it('prints raw directory JSON and follows continuation links with --cursor and --all', async () => {
  const nextPage = `${servers.graph.url}/users?$skiptoken=next-page`;
  servers.graph.setReply((request, response) => {
    response.writeHead(200, { 'Content-Type': 'application/json' });
    if (request.path === '/users?%24top=25') response.end(JSON.stringify({ value: [{ id: 'user-1', displayName: 'Ada' }], '@odata.nextLink': nextPage }));
    else if (request.path === '/users?$skiptoken=next-page') response.end(JSON.stringify({ value: [{ id: 'user-2', displayName: 'Grace' }] }));
    else response.end(JSON.stringify({ value: [] }));
  });
  servers.graph.reset();

  const first = await cli.run(['users', '--json']);
  assert.equal(first.code, 0, first.stderr);
  assert.deepEqual(JSON.parse(first.stdout), { value: [{ id: 'user-1', displayName: 'Ada' }], '@odata.nextLink': nextPage });
  const cursor = first.stderr.match(/--cursor ([^ ]+)/)?.[1];
  assert.ok(cursor);

  servers.graph.reset();
  const next = await cli.run(['users', '--cursor', cursor, '--json']);
  assert.equal(next.code, 0, next.stderr);
  assert.deepEqual(JSON.parse(next.stdout), { value: [{ id: 'user-2', displayName: 'Grace' }] });
  assert.equal(servers.graph.requests[0]?.path, '/users?$skiptoken=next-page');

  servers.graph.reset();
  const all = await cli.run(['users', '--all', '--json']);
  assert.equal(all.code, 0, all.stderr);
  assert.deepEqual(JSON.parse(all.stdout), { value: [{ id: 'user-1', displayName: 'Ada' }, { id: 'user-2', displayName: 'Grace' }] });
  assert.deepEqual(servers.graph.requests.map(({ path }) => path), ['/users?%24top=25', '/users?$skiptoken=next-page']);
});

it('searches Users by display name with advanced-query consistency', async () => {
  servers.graph.reset();
  servers.graph.setReply((_request, response) => { response.writeHead(200, { 'Content-Type': 'application/json' }); response.end(JSON.stringify({ value: [] })); });
  const result = await cli.run(['users', '--search', 'grace']);
  assert.equal(result.code, 0, result.stderr);
  const request = servers.graph.requests[0]!;
  assert.equal(new URL(`http://x${request.path}`).searchParams.get('$search'), '"displayName:grace"');
  assert.equal(request.headers.consistencylevel, 'eventual');
  const refused = await cli.run(['groups', '--search', 'x']);
  assert.equal(refused.code, 2);
});
