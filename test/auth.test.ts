import assert from 'node:assert/strict';
import { readFileSync, statSync, writeFileSync } from 'node:fs';
import { after, before, it } from 'node:test';
import { createCli, type Cli } from './helpers/cli.ts';
import { startFakeServers, type FakeServers } from './helpers/fake-server.ts';

let servers: FakeServers;
let cli: Cli;
before(async () => {
  servers = await startFakeServers();
  cli = createCli({ OUTLOOK_LOGIN_URL: servers.login.url, OUTLOOK_GRAPH_URL: servers.graph.url });
  let polls = 0;
  servers.login.setReply((request, response) => {
    if (request.path.endsWith('/devicecode')) {
      response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify({ device_code: 'device-secret', user_code: 'ABCD-EFGH', verification_uri: 'https://login.example.test/device', interval: 1, expires_in: 30 }));
    } else if (request.path.endsWith('/token')) {
      polls += 1; response.writeHead(200, { 'Content-Type': 'application/json' });
      response.end(JSON.stringify(polls === 1 ? { error: 'authorization_pending' } : { access_token: 'access-secret', refresh_token: 'refresh-secret', expires_in: 3600 }));
    } else response.writeHead(404).end();
  });
});
after(async () => { await servers.close(); cli.cleanup(); });
function setConfig(config: unknown): void { writeFileSync(`${cli.home}/config.json`, JSON.stringify(config), { mode: 0o600 }); }
function graphReply(status = 200): void {
  servers.graph.setReply((request, response) => {
    if (status !== 200) {
      response.writeHead(status, { 'Content-Type': 'application/json', ...(status === 429 ? { 'Retry-After': '0' } : {}) });
      response.end(JSON.stringify({ error: 'fake failure' })); return;
    }
    assert.equal(request.method, 'GET'); assert.equal(request.path, '/me');
    response.writeHead(200, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify({ displayName: 'Outlook User', userPrincipalName: 'user@example.test', mail: 'user@example.test' }));
  });
  servers.graph.reset();
}

it('logs in with device code and saves refresh credentials in a private config', async () => {
  servers.login.reset();
  const result = await cli.run(['auth', 'login', '--client-id', 'test-client']);
  assert.equal(result.code, 0, result.stderr);
  assert.match(result.stdout, /https:\/\/login\.example\.test\/device/);
  assert.match(result.stdout, /ABCD-EFGH/);
  assert.match(result.stdout, /Signed in/);
  assert.equal(servers.login.requests.length, 3);
  const device = servers.login.requests[0]!;
  assert.equal(device.method, 'POST');
  assert.equal(device.path, '/common/oauth2/v2.0/devicecode');
  const deviceBody = new URLSearchParams(device.body);
  assert.equal(deviceBody.get('client_id'), 'test-client');
  assert.match(deviceBody.get('scope') ?? '', /offline_access/);
  const tokenBody = new URLSearchParams(servers.login.requests[2]!.body);
  assert.equal(tokenBody.get('grant_type'), 'urn:ietf:params:oauth:grant-type:device_code');
  const config = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8')) as { profiles: Record<string, { refreshToken: string }> };
  assert.equal(config.profiles.default?.refreshToken, 'refresh-secret');
  assert.equal(statSync(`${cli.home}/config.json`).mode & 0o777, 0o600);
});

it('uses an environment token without a Profile and reports specific Graph failures', async () => {
  graphReply();
  const ok = await cli.run(['auth', 'check'], { OUTLOOK_TOKEN: 'headless-token' });
  assert.equal(ok.code, 0, ok.stderr);
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer headless-token');
  for (const [status, phrase] of [[401, 'Authentication failed'], [403, 'Access denied'], [404, 'Not found'], [429, 'Rate limited']] as const) {
    graphReply(status);
    const result = await cli.run(['auth', 'check'], { OUTLOOK_TOKEN: 'headless-token' });
    assert.equal(result.code, 1);
    assert.match(result.stderr, new RegExp(`${phrase}.*GET \/me`));
  }
});

it('selects Profiles, checks identity, masks credentials, and logs out only that Profile', async () => {
  setConfig({ defaultProfile: 'primary', profiles: { primary: { accessToken: 'primary-token', refreshToken: 'primary-refresh' }, backup: { accessToken: 'backup-token', refreshToken: 'backup-refresh' } } });
  graphReply();
  const checked = await cli.run(['auth', 'check', '-p', 'backup'], { OUTLOOK_PROFILE: 'primary' });
  assert.equal(checked.code, 0, checked.stderr);
  assert.match(checked.stdout, /Outlook User/);
  assert.equal(servers.graph.requests[0]!.headers.authorization, 'Bearer backup-token');
  graphReply();
  const me = await cli.run(['me', '-p', 'backup']);
  assert.equal(me.code, 0, me.stderr);
  assert.match(me.stdout, /user@example\.test/);
  const shown = await cli.run(['config', 'show']);
  assert.equal(shown.code, 0, shown.stderr);
  assert.doesNotMatch(shown.stdout, /primary-token|primary-refresh|backup-token|backup-refresh/);
  assert.equal((JSON.parse(shown.stdout) as { profiles: Record<string, { refreshToken: string }> }).profiles.backup?.refreshToken, '********');
  const loggedOut = await cli.run(['auth', 'logout', '-p', 'backup']);
  assert.equal(loggedOut.code, 0, loggedOut.stderr);
  const saved = JSON.parse(readFileSync(`${cli.home}/config.json`, 'utf8')) as { profiles: Record<string, Record<string, unknown>> };
  assert.equal(saved.profiles.backup?.accessToken, undefined);
  assert.equal(saved.profiles.primary?.accessToken, 'primary-token');
});
