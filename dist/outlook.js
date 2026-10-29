#!/usr/bin/env node

// package.json
var package_default = {
  name: "outlook-cli",
  version: "0.1.0",
  description: "Command-line client for Microsoft Outlook",
  type: "module",
  bin: { outlook: "./dist/outlook.js" },
  scripts: {
    typecheck: "tsc --noEmit && tsc --noEmit -p tsconfig.test.json",
    build: 'pnpm typecheck && esbuild src/index.ts --bundle --platform=node --target=node20 --format=esm --banner:js="#!/usr/bin/env node" --outfile=dist/outlook.js && chmod +x dist/outlook.js',
    test: 'pnpm build && node --test "test/*.test.ts"',
    "plugin:sync": "node scripts/sync-plugin-manifests.mjs",
    "plugin:check": "node scripts/sync-plugin-manifests.mjs --check"
  },
  engines: { node: ">=20" },
  devDependencies: { "@types/node": "^20.19.0", esbuild: "^0.28.2", typescript: "^7.0.2" },
  private: true,
  license: "MIT",
  packageManager: "pnpm@12.3.4"
};

// src/help.ts
var HELP = `Usage: outlook <command> [options]
       outlook --help
       outlook --version

Commands:
  auth       Sign in and manage Profiles
  config     Show configuration with secrets masked
  me         Show the signed-in user
  api        Call any Microsoft Graph endpoint
  draft      Create or update a Draft
  send       Send a new Message or saved Draft (requires --yes)
  reply      Reply to a Message (--yes to send, --draft to save)
  reply-all  Reply to all recipients (--yes to send, --draft to save)
  forward    Forward a Message (--yes to send, --draft to save)
  mark       Mark Messages read or unread
  delete     Delete Messages (requires --yes)
  folders    List Mail folders
  messages   List Messages in Inbox or another Mail folder
  message    Read one Message by id or Outlook URL
  contacts   List your Contacts
  users      List or search directory Users
  groups     List Groups
  members    List a Group's members

Commands and options:
  auth login | check | logout
  config show
  me [--json]
  api <METHOD> <path> [key=value ...] [--data JSON|@file|-] [--header 'Name: value'] [--all]
  draft [--to <addresses>] [--cc <addresses>] [--bcc <addresses>] [--subject <text>] [--body <text|@file|->] [--html]
  draft --update <id> [fields]
  send [--to <addresses>] [--cc <addresses>] [--bcc <addresses>] [--subject <text>] [--body <text|@file|->] [--html] --yes
  send --draft <id> --yes
  reply <id> --body <text> (--yes|--draft)
  reply-all <id> --body <text> (--yes|--draft)
  forward <id> --to <addresses> --body <text> (--yes|--draft)
  mark <id...> --read|--unread
  delete <id...> --yes
  folders [--cursor <cursor> | --all] [--json]
  messages [--folder <name|id>] [--unread] [--from <address>] [--since <date>]
           [--search <text>] [--top <count>] [--cursor <cursor> | --all] [--json]
  message <id|url> [--html] [--json]
  contacts | groups [--top N] [--cursor CURSOR] [--all] [--json]
  users [--search <text>] [--top N] [--cursor CURSOR] [--all] [--json]
  members <group> [--top N] [--cursor CURSOR] [--all] [--json]

Global options:
  -p, --profile <name>  Select a Profile (or set OUTLOOK_PROFILE)
  --client-id <id>      Override the OAuth client id (or set OUTLOOK_CLIENT_ID)
  --scopes <list>       OAuth scopes for sign-in, space or comma separated (or set OUTLOOK_SCOPES); saved in the Profile on login
  -h, --help            Show help
      --version         Show the installed version

Environment:
  OUTLOOK_PROFILE, OUTLOOK_TOKEN, OUTLOOK_CLIENT_ID, OUTLOOK_SCOPES, OUTLOOK_CONFIG
  OUTLOOK_GRAPH_URL and OUTLOOK_LOGIN_URL (for alternate services and tests)

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;

// src/utils/errors.ts
var CliError = class extends Error {
  exitCode;
  status;
  constructor(message, exitCode = 1, status) {
    super(message);
    this.exitCode = exitCode;
    this.status = status;
  }
};
function usage(message) {
  throw new CliError(message, 2);
}

// src/utils/config.ts
import { chmodSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname } from "node:path";
function configPath(env) {
  return env.OUTLOOK_CONFIG || `${homedir()}/.config/outlook/config.json`;
}
function loadConfig(path) {
  try {
    const value = JSON.parse(readFileSync(path, "utf8"));
    if (!value.profiles || typeof value.profiles !== "object") return { profiles: {} };
    return { defaultProfile: value.defaultProfile, profiles: value.profiles };
  } catch (error) {
    if (error.code === "ENOENT") return { profiles: {} };
    throw new CliError(`Cannot read config file ${path}: ${error.message}`);
  }
}
function saveConfig(path, config) {
  mkdirSync(dirname(path), { recursive: true, mode: 448 });
  writeFileSync(path, `${JSON.stringify(config, null, 2)}
`, { mode: 384 });
  chmodSync(path, 384);
}
function maskSecrets(config) {
  return {
    ...config,
    profiles: Object.fromEntries(Object.entries(config.profiles).map(([name, profile]) => [name, {
      ...profile,
      ...profile.accessToken ? { accessToken: "********" } : {},
      ...profile.refreshToken ? { refreshToken: "********" } : {}
    }]))
  };
}

// src/commands/auth-commands.ts
var DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
var DEFAULT_SCOPES = "Mail.ReadWrite Mail.Send Contacts.Read User.ReadBasic.All Group.Read.All offline_access";
function context(env, options) {
  const path = configPath(env);
  const config = loadConfig(path);
  const profileName = options.profile || env.OUTLOOK_PROFILE || config.defaultProfile || "default";
  config.profiles[profileName] ??= {};
  return { env, options, config, path, profileName, profile: config.profiles[profileName] };
}
function persist(ctx) {
  ctx.config.profiles[ctx.profileName] = ctx.profile;
  ctx.config.defaultProfile ??= ctx.profileName;
  saveConfig(ctx.path, ctx.config);
}
function explicitScopes(ctx) {
  const value = ctx.options.scopes || ctx.env.OUTLOOK_SCOPES;
  return value ? value.split(/[\s,]+/).filter(Boolean).join(" ") : void 0;
}
function scopesFor(ctx) {
  return explicitScopes(ctx) || ctx.profile.scopes || DEFAULT_SCOPES;
}
function clientIdFor(ctx) {
  return ctx.options.clientId || ctx.env.OUTLOOK_CLIENT_ID || ctx.profile.clientId || DEFAULT_CLIENT_ID;
}
function loginBase(ctx) {
  return (ctx.env.OUTLOOK_LOGIN_URL || "https://login.microsoftonline.com").replace(/\/$/, "");
}
function tokenEndpoint(ctx) {
  return `${loginBase(ctx)}/common/oauth2/v2.0/token`;
}
async function getAccessToken(ctx) {
  if (ctx.env.OUTLOOK_TOKEN) return ctx.env.OUTLOOK_TOKEN;
  if (!ctx.profile.accessToken) throw new CliError(`Profile '${ctx.profileName}' is not signed in. Run 'outlook auth login'.`);
  if (ctx.profile.expiresAt && ctx.profile.expiresAt <= Date.now() + 6e4) await refresh(ctx);
  return ctx.profile.accessToken;
}
async function refresh(ctx) {
  if (!ctx.profile.refreshToken) throw new CliError("Sign-in expired. Run 'outlook auth login' to sign in again.");
  let response;
  try {
    response = await fetch(tokenEndpoint(ctx), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: clientIdFor(ctx), grant_type: "refresh_token", refresh_token: ctx.profile.refreshToken, scope: scopesFor(ctx) })
    });
  } catch {
    throw new CliError("Could not refresh sign-in. Run 'outlook auth login' to sign in again.");
  }
  const data = await response.json();
  if (!response.ok || !data.access_token) throw new CliError("Sign-in expired. Run 'outlook auth login' to sign in again.");
  ctx.profile.accessToken = data.access_token;
  ctx.profile.refreshToken = data.refresh_token || ctx.profile.refreshToken;
  ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1e3;
  persist(ctx);
}
var MAX_THROTTLE_RETRIES = 2;
async function graph(ctx, method, path, options = {}, retried = false, throttled = 0) {
  const base = (ctx.env.OUTLOOK_GRAPH_URL || "https://graph.microsoft.com/v1.0").replace(/\/$/, "");
  const url = new URL(options.rawUrl || `${base}${path.startsWith("/") ? path : `/${path}`}`);
  options.query?.forEach((value, key) => url.searchParams.append(key, value));
  const token = await getAccessToken(ctx);
  let response;
  try {
    response = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        ...options.body === void 0 ? {} : { "Content-Type": "application/json" },
        ...options.headers
      },
      ...options.body === void 0 ? {} : { body: options.body }
    });
  } catch {
    throw new CliError(`Cannot reach ${url} for ${method} ${path}. Check the network and Graph URL.`);
  }
  if (response.status === 401 && !retried && !ctx.env.OUTLOOK_TOKEN && ctx.profile.refreshToken) {
    await refresh(ctx);
    return graph(ctx, method, path, options, true, throttled);
  }
  if (response.status === 429 && throttled < MAX_THROTTLE_RETRIES) {
    const header = response.headers.get("retry-after");
    const wait = header === null ? Number.NaN : Number(header);
    const seconds = Number.isFinite(wait) && wait >= 0 ? Math.min(wait, 30) : 2 ** throttled;
    await new Promise((resolve) => setTimeout(resolve, seconds * 1e3));
    return graph(ctx, method, path, options, retried, throttled + 1);
  }
  if (!response.ok) {
    const detail = await graphErrorDetail(response);
    const messages = {
      401: `Authentication failed for ${method} ${path}. Run 'outlook auth login'.`,
      403: `Access denied for ${method} ${path}. Check that your account has the required Outlook permissions.`,
      404: `Not found: ${method} ${path}. Check the id and path.`,
      429: `Rate limited for ${method} ${path}. Wait, then try again.`
    };
    throw new CliError(`${messages[response.status] || `HTTP ${response.status} for ${method} ${path}.`}${detail ? ` Graph said: ${detail}` : ""}`, 1, response.status);
  }
  const text = await response.text();
  return text ? JSON.parse(text) : void 0;
}
async function graphErrorDetail(response) {
  try {
    const body = JSON.parse(await response.text());
    if (typeof body.error === "string") return body.error;
    return [body.error?.code, body.error?.message].filter(Boolean).join(": ");
  } catch {
    return "";
  }
}
async function runAuth(ctx, subcommand) {
  if (subcommand === "login") return login(ctx);
  if (subcommand === "check") {
    const me = await graph(ctx, "GET", "/me");
    process.stdout.write(`Signed in as ${me.displayName || me.userPrincipalName || me.mail || "current user"}
`);
    return;
  }
  if (subcommand === "logout") {
    delete ctx.profile.accessToken;
    delete ctx.profile.refreshToken;
    delete ctx.profile.expiresAt;
    persist(ctx);
    process.stdout.write(`Signed out of Profile '${ctx.profileName}'.
`);
    return;
  }
  usage("Use 'outlook auth login', 'outlook auth check', or 'outlook auth logout'.");
}
async function login(ctx) {
  const clientId = clientIdFor(ctx);
  const deviceUrl = `${loginBase(ctx)}/common/oauth2/v2.0/devicecode`;
  let deviceResponse;
  try {
    deviceResponse = await fetch(deviceUrl, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: clientId, scope: scopesFor(ctx) })
    });
  } catch {
    throw new CliError(`Cannot reach ${deviceUrl}. Check the login URL and network.`);
  }
  const device = await deviceResponse.json();
  if (!deviceResponse.ok || !device.device_code) throw new CliError("Could not start device-code sign-in. Check the client id and try again.");
  process.stdout.write(`To sign in, visit ${device.verification_uri} and enter code ${device.user_code}.
`);
  const interval = Math.max(1, device.interval || 5) * 1e3;
  const expiresAt = Date.now() + (device.expires_in || 900) * 1e3;
  while (Date.now() < expiresAt) {
    const response = await fetch(tokenEndpoint(ctx), {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: clientId, grant_type: "urn:ietf:params:oauth:grant-type:device_code", device_code: device.device_code })
    });
    const data = await response.json();
    if (response.ok && data.access_token) {
      ctx.profile.accessToken = data.access_token;
      ctx.profile.refreshToken = data.refresh_token;
      ctx.profile.scopes = explicitScopes(ctx) ?? ctx.profile.scopes;
      ctx.profile.clientId = clientId;
      ctx.profile.expiresAt = Date.now() + (data.expires_in ?? 3600) * 1e3;
      persist(ctx);
      process.stdout.write(`Signed in to Profile '${ctx.profileName}'.
`);
      return;
    }
    if (data.error === "authorization_pending") {
      await new Promise((resolve) => setTimeout(resolve, interval));
      continue;
    }
    if (data.error === "slow_down") {
      await new Promise((resolve) => setTimeout(resolve, interval + 5e3));
      continue;
    }
    if (data.error === "authorization_declined" || data.error === "expired_token") throw new CliError(`Device-code sign-in ${data.error === "expired_token" ? "expired" : "was declined"}. Run 'outlook auth login' to try again.`);
    throw new CliError("Device-code sign-in failed. Run 'outlook auth login' to try again.");
  }
  throw new CliError("Device code expired. Run 'outlook auth login' to try again.");
}
async function graphRequest(ctx, method, path, options) {
  return graph(ctx, method, path, options);
}
function selectedContext(env, options) {
  return context(env, options);
}

// src/commands/config-commands.ts
async function runConfig(ctx, args) {
  if (args.length === 1 && args[0] === "show") {
    process.stdout.write(`${JSON.stringify(maskSecrets(ctx.config), null, 2)}
`);
    return;
  }
  usage("Use 'outlook config show'.");
}

// src/utils/message-id.ts
function messageId(input) {
  let id = input;
  if (/^https?:\/\//i.test(input)) {
    let url;
    try {
      url = new URL(input);
    } catch {
      usage("Invalid Outlook on the web URL. Copy the Message URL and try again.");
    }
    if (!/(^|\.)outlook\.(office\.com|office365\.com|live\.com)$/i.test(url.hostname)) {
      usage("Expected an Outlook on the web URL. Copy the Message URL and try again.");
    }
    const pathParts = url.pathname.split("/").filter(Boolean);
    const idIndex = pathParts.findIndex((part) => part.toLowerCase() === "id");
    const deeplinkIndex = pathParts.findIndex((part, index) => part.toLowerCase() === "read" && pathParts[index - 1]?.toLowerCase() === "deeplink");
    id = (idIndex >= 0 ? pathParts[idIndex + 1] : void 0) ?? (deeplinkIndex >= 0 ? pathParts[deeplinkIndex + 1] : void 0) ?? url.searchParams.get("ItemID") ?? url.searchParams.get("itemid") ?? "";
    try {
      id = decodeURIComponent(id);
    } catch {
      usage("Invalid Message id in Outlook URL.");
    }
  }
  if (!/^[A-Za-z0-9_+\/=\-]+$/.test(id)) usage("Invalid Message id. Use a Graph id or Outlook on the web URL.");
  return id;
}

// src/utils/graph-page.ts
function graphPage(ctx, url, headers) {
  return /^https?:\/\//i.test(url) ? graphRequest(ctx, "GET", url, { rawUrl: url, ...headers ? { headers } : {} }) : graphRequest(ctx, "GET", url, headers ? { headers } : void 0);
}
function validateContinuation(link, ctx) {
  const baseValue = ctx.env.OUTLOOK_GRAPH_URL || "https://graph.microsoft.com/v1.0";
  let url;
  try {
    url = new URL(link, baseValue);
  } catch {
    usage("Graph returned an invalid continuation URL.");
  }
  const base = new URL(baseValue);
  const basePath = base.pathname.replace(/\/$/, "");
  const inBasePath = !basePath || url.pathname === basePath || url.pathname.startsWith(`${basePath}/`);
  if (url.origin !== base.origin || !inBasePath) usage("Graph returned a continuation URL outside the configured Graph URL.");
}

// src/utils/mail-folders.ts
var PAGE_CAP = 100;
var WELL_KNOWN = /* @__PURE__ */ new Set(["inbox", "drafts", "sentitems", "deleteditems", "archive", "junkemail", "outbox"]);
async function resolveMailFolderId(ctx, value) {
  if (WELL_KNOWN.has(value.trim().toLowerCase())) return value.trim().toLowerCase();
  const folders = await listMailFolders(ctx);
  const matches = folders.filter((folder) => folder.id === value || equal(folder.displayName, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  fail(value, matches.length > 1, (matches.length ? matches : folders).map((folder) => `${folder.displayName ?? "Untitled"} (${folder.id ?? "unknown id"})`));
}
async function listMailFolders(ctx) {
  const folders = [];
  let next = "/me/mailFolders";
  for (let page = 0; page < PAGE_CAP; page += 1) {
    const response = await graphPage(ctx, next);
    folders.push(...response.value ?? []);
    if (!response["@odata.nextLink"]) return folders;
    next = response["@odata.nextLink"];
    validateContinuation(next, ctx);
  }
  return folders;
}
function equal(left, right) {
  return (left ?? "").trim().toLocaleLowerCase() === right.trim().toLocaleLowerCase();
}
function fail(value, ambiguous, candidates) {
  usage(`${ambiguous ? "Ambiguous" : "Unknown"} Mail folder '${value}'. Candidates: ${candidates.join(", ") || "none"}.`);
}

// src/commands/read-commands.ts
var ALL_PAGE_CAP = 100;
function table(headers, rows) {
  const widths = headers.map((header, index) => Math.max(header.length, ...rows.map((row) => row[index]?.length ?? 0)));
  const line = (values) => values.map((value, index) => value.padEnd(widths[index])).join("  ").trimEnd();
  return `${line(headers)}
${rows.map(line).join("\n")}${rows.length ? "\n" : ""}`;
}
async function runReadCommand(ctx, command, args) {
  if (command === "message") return readMessage(ctx, args);
  if (["contacts", "users", "groups", "members"].includes(command)) return readDirectory(ctx, command, args);
  let response;
  let rows;
  let nextLink;
  if (command === "me") response = await graphRequest(ctx, "GET", "/me");
  else if (command === "folders") {
    const result = await readCollection(ctx, "/me/mailFolders", args);
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else if (command === "messages") {
    const folder = option(args, "--folder");
    const folderId = folder ? await resolveMailFolderId(ctx, folder) : "inbox";
    const result = await readCollection(ctx, `/me/mailFolders/${encodeURIComponent(folderId)}/messages`, args, messageQuery(args));
    response = result.response;
    rows = result.items;
    nextLink = result.nextLink;
  } else usage(`Unknown read command: ${command}`);
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}
`);
    printContinuation(nextLink);
    return;
  }
  if (command === "me") {
    const user = response;
    process.stdout.write(table(["NAME", "EMAIL"], [[user.displayName ?? "Current user", user.mail ?? user.userPrincipalName ?? ""]]));
    return;
  }
  if (command === "folders") {
    const folders = rows ?? response.value ?? [];
    process.stdout.write(table(["ID", "NAME", "TOTAL", "UNREAD"], folders.map((folder) => [folder.id ?? "", folder.displayName ?? "", String(folder.totalItemCount ?? ""), String(folder.unreadItemCount ?? "")])));
    printContinuation(nextLink);
    return;
  }
  const messages = rows ?? response.value ?? [];
  process.stdout.write(table(["ID", "RECEIVED", "FROM", "SUBJECT", "READ"], messages.map((message) => [
    message.id ?? "",
    message.receivedDateTime ?? "",
    message.from?.emailAddress?.name || message.from?.emailAddress?.address || "",
    message.subject ?? "",
    message.isRead === void 0 ? "" : message.isRead ? "yes" : "no"
  ])));
  printContinuation(nextLink);
}
async function readCollection(ctx, path, args, query, headers) {
  const token = option(args, "--cursor");
  let url = token ? decodeCursor(token, ctx) : withQuery(path, query);
  const all = args.includes("--all");
  const pages = [];
  let nextLink;
  for (let page = 0; page < (all ? ALL_PAGE_CAP : 1); page += 1) {
    const result = await graphPage(ctx, url, headers);
    pages.push(result);
    nextLink = result["@odata.nextLink"];
    if (!all || !nextLink) break;
    validateContinuation(nextLink, ctx);
    url = nextLink;
  }
  const first = pages[0] ?? { value: [] };
  const items = pages.flatMap((page) => page.value ?? []);
  const response = all ? { ...first, value: items, ...nextLink ? { "@odata.nextLink": nextLink } : {} } : first;
  if (all && !nextLink) delete response["@odata.nextLink"];
  return { response, items, nextLink };
}
function messageQuery(args) {
  const query = new URLSearchParams();
  const top = option(args, "--top") ?? "25";
  if (!/^\d+$/.test(top) || Number(top) < 1 || Number(top) > 1e3) usage("Option '--top' must be an integer from 1 to 1000.");
  query.set("$top", top);
  const search = option(args, "--search");
  const unread = args.includes("--unread");
  const from = option(args, "--from");
  const since = option(args, "--since");
  if (search) {
    if (unread || from || since) usage("Option '--search' cannot be combined with '--unread', '--from', or '--since'.");
    query.set("$search", `"${search.replaceAll('"', '\\"')}"`);
    return query;
  }
  query.set("$orderby", "receivedDateTime desc");
  let sinceDate = /* @__PURE__ */ new Date(0);
  if (since) {
    sinceDate = new Date(since);
    if (Number.isNaN(sinceDate.getTime())) usage("Option '--since' must be a valid date.");
  }
  const filters = [`receivedDateTime ge ${sinceDate.toISOString()}`];
  if (unread) filters.push("isRead eq false");
  if (from) filters.push(`from/emailAddress/address eq '${from.replaceAll("'", "''")}'`);
  query.set("$filter", filters.join(" and "));
  return query;
}
function withQuery(path, query) {
  return query && [...query].length ? `${path}?${query.toString()}` : path;
}
function decodeCursor(token, ctx) {
  let link;
  try {
    link = Buffer.from(token, "base64url").toString("utf8");
  } catch {
    usage("Invalid cursor. Copy the cursor from the previous command output.");
  }
  validateContinuation(link, ctx);
  return link;
}
function printContinuation(nextLink) {
  if (nextLink) process.stderr.write(`more available: --cursor ${Buffer.from(nextLink).toString("base64url")} (or --all)
`);
}
function option(args, name) {
  const index = args.findIndex((arg2) => arg2 === name || arg2.startsWith(`${name}=`));
  if (index < 0) return void 0;
  const arg = args[index];
  const value = arg === name ? args[index + 1] : arg.slice(name.length + 1);
  if (!value || value.startsWith("-")) usage(`Option '${name}' requires a value.`);
  return value;
}
async function readMessage(ctx, args) {
  if (args.some((arg) => arg.startsWith("--") && arg !== "--html" && arg !== "--json")) {
    usage("Use 'outlook message <id|url> [--html] [--json]'.");
  }
  const positional = args.filter((arg) => !arg.startsWith("--"));
  if (positional.length !== 1) usage("Use 'outlook message <id|url> [--html] [--json]'.");
  const id = messageId(positional[0]);
  const html = args.includes("--html");
  const response = await graphRequest(ctx, "GET", `/me/messages/${encodeURIComponent(id)}`, {
    ...html ? {} : { headers: { Prefer: 'outlook.body-content-type="text"' } }
  });
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(response, null, 2)}
`);
    return;
  }
  const message = response;
  process.stdout.write(`${message.body?.content ?? ""}
`);
}
async function readDirectory(ctx, command, args) {
  let path;
  if (command === "contacts") path = "/me/contacts";
  else if (command === "users") path = "/users";
  else if (command === "groups") path = "/groups";
  else {
    const group = args[0];
    if (!group || group.startsWith("-")) usage("A group id is required. Usage: 'outlook members <group>'.");
    path = `/groups/${encodeURIComponent(group)}/members`;
  }
  const top = option(args, "--top") ?? "25";
  if (!/^\d+$/.test(top) || Number(top) < 1 || Number(top) > 999) usage("Option '--top' must be an integer from 1 to 999.");
  const query = new URLSearchParams({ $top: top });
  const search = option(args, "--search");
  if (search) {
    if (command !== "users") usage("Option '--search' is only supported by 'users'.");
    query.set("$search", `"displayName:${search.replaceAll('"', '\\"')}"`);
  }
  const result = await readCollection(ctx, path, args, query, search ? { ConsistencyLevel: "eventual" } : void 0);
  if (args.includes("--json")) {
    process.stdout.write(`${JSON.stringify(result.response, null, 2)}
`);
    printContinuation(result.nextLink);
    return;
  }
  process.stdout.write(table(["ID", "NAME", "EMAIL"], result.items.map((item) => [
    item.id ?? "",
    item.displayName ?? "",
    item.mail ?? item.userPrincipalName ?? item.emailAddresses?.[0]?.address ?? ""
  ])));
  printContinuation(result.nextLink);
}

// src/commands/api-commands.ts
import { readFile } from "node:fs/promises";
async function runApi(ctx, args) {
  const method = args[0]?.toUpperCase();
  const path = args[1];
  if (!method || !path) usage("Use 'outlook api <METHOD> <path> [key=value ...]'.");
  if (!["GET", "POST", "PUT", "PATCH", "DELETE"].includes(method)) usage(`Unsupported HTTP method: ${args[0]}.`);
  const query = new URLSearchParams();
  const headers = {};
  let body;
  let all = false;
  for (let index = 2; index < args.length; index += 1) {
    const argument = args[index];
    if (argument === "--all") {
      all = true;
      continue;
    }
    if (argument === "--header" || argument.startsWith("--header=")) {
      const header = argument === "--header" ? args[++index] : argument.slice("--header=".length);
      const colon = header?.indexOf(":") ?? -1;
      if (!header || colon <= 0) usage("Option '--header' requires 'Name: value'.");
      headers[header.slice(0, colon).trim()] = header.slice(colon + 1).trim();
      continue;
    }
    if (argument === "--data" || argument.startsWith("--data=")) {
      if (body !== void 0) usage("Option '--data' may only be provided once.");
      const source = argument === "--data" ? args[++index] : argument.slice("--data=".length);
      if (source === void 0) usage("Option '--data' requires JSON, @file, or '-'.");
      const text = source === "-" ? await readStdin() : source.startsWith("@") ? await readFile(source.slice(1), "utf8") : source;
      try {
        body = JSON.stringify(JSON.parse(text));
      } catch {
        usage("Option '--data' must contain valid JSON.");
      }
      continue;
    }
    const separator = argument.indexOf("=");
    if (separator <= 0) usage(`Expected a query pair in key=value form: ${argument}.`);
    query.append(argument.slice(0, separator), argument.slice(separator + 1));
  }
  if (all && method !== "GET") usage("Option '--all' is only supported with GET requests.");
  let data = await graphRequest(ctx, method, path, { query, body, headers });
  if (all) data = await collectPages(ctx, data);
  if (data !== void 0) process.stdout.write(`${JSON.stringify(data)}
`);
}
async function collectPages(ctx, initial) {
  if (!isCollection(initial)) return initial;
  const values = [...initial.value];
  const result = { ...initial, value: values };
  delete result["@odata.nextLink"];
  let nextLink = initial["@odata.nextLink"];
  for (let page = 1; typeof nextLink === "string" && page < 50; page += 1) {
    validateContinuation(nextLink, ctx);
    const nextUrl = new URL(nextLink);
    const data = await graphRequest(ctx, "GET", nextUrl.pathname, { rawUrl: nextUrl.href });
    if (!isCollection(data)) break;
    values.push(...data.value);
    nextLink = data["@odata.nextLink"];
  }
  return result;
}
function isCollection(value) {
  return typeof value === "object" && value !== null && Array.isArray(value.value);
}
async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks).toString("utf8");
}

// src/utils/compose.ts
import { readFileSync as readFileSync2 } from "node:fs";
function recipients(values) {
  if (!values) return void 0;
  const addresses = values.flatMap((value) => value.split(",")).map((address) => address.trim()).filter(Boolean);
  if (addresses.length === 0) usage("Recipient options must include at least one email address.");
  return addresses.map((address) => ({ emailAddress: { address } }));
}
async function bodyContent(value) {
  if (value === "-") {
    let body = "";
    for await (const chunk of process.stdin) body += chunk.toString();
    return body;
  }
  if (value.startsWith("@")) {
    const path = value.slice(1);
    if (!path) usage("Option '--body @file' requires a file path.");
    try {
      return readFileSync2(path, "utf8");
    } catch {
      usage(`Cannot read body file '${path}'. Check the file path and permissions.`);
    }
  }
  return value;
}

// src/commands/draft-commands.ts
function parseArgs(args) {
  const options = { html: false };
  const values = /* @__PURE__ */ new Map();
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "--html") {
      options.html = true;
      continue;
    }
    if (!["--update", "--to", "--cc", "--bcc", "--subject", "--body"].includes(arg)) {
      usage(`Unknown draft option: ${arg}.`);
    }
    const value = args[index + 1];
    if (value === void 0 || value.startsWith("--") && value !== "-") usage(`Option '${arg}' requires a value.`);
    const items = values.get(arg) ?? [];
    items.push(value);
    values.set(arg, items);
    index += 1;
  }
  const one = (name) => {
    const items = values.get(name);
    if (!items) return void 0;
    if (items.length > 1) usage(`Option '${name}' may be used only once.`);
    return items[0];
  };
  options.update = one("--update");
  options.to = values.get("--to");
  options.cc = values.get("--cc");
  options.bcc = values.get("--bcc");
  options.subject = one("--subject");
  options.body = one("--body");
  if (options.update !== void 0 && !options.update) usage("Option '--update' requires a Draft id.");
  if (options.body === void 0 && options.html) usage("Option '--html' requires '--body'.");
  return options;
}
async function runDraftCommand(ctx, args) {
  const options = parseArgs(args);
  const body = options.body === void 0 ? void 0 : {
    contentType: options.html ? "HTML" : "Text",
    content: await bodyContent(options.body)
  };
  const patch = {
    ...options.to ? { toRecipients: recipients(options.to) } : {},
    ...options.cc ? { ccRecipients: recipients(options.cc) } : {},
    ...options.bcc ? { bccRecipients: recipients(options.bcc) } : {},
    ...options.subject === void 0 ? {} : { subject: options.subject },
    ...body === void 0 ? {} : { body }
  };
  const method = options.update === void 0 ? "POST" : "PATCH";
  const path = options.update === void 0 ? "/me/messages" : `/me/messages/${encodeURIComponent(messageId(options.update))}`;
  const response = await graphRequest(ctx, method, path, { body: JSON.stringify(patch) });
  if (options.update === void 0) {
    const id = response?.id;
    process.stdout.write(id ? `Draft created: ${id}
` : "Draft created.\n");
  } else process.stdout.write(`Draft updated: ${options.update}
`);
}

// src/commands/send-commands.ts
function parseArgs2(args) {
  const options = { html: false, yes: false };
  const values = /* @__PURE__ */ new Map();
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "--html") {
      options.html = true;
      continue;
    }
    if (arg === "--yes") {
      options.yes = true;
      continue;
    }
    if (!["--draft", "--to", "--cc", "--bcc", "--subject", "--body"].includes(arg)) {
      usage(`Unknown send option: ${arg}.`);
    }
    const value = args[index + 1];
    if (value === void 0 || value.startsWith("--") && value !== "-") usage(`Option '${arg}' requires a value.`);
    const items = values.get(arg) ?? [];
    items.push(value);
    values.set(arg, items);
    index += 1;
  }
  const one = (name) => {
    const items = values.get(name);
    if (!items) return void 0;
    if (items.length > 1) usage(`Option '${name}' may be used only once.`);
    return items[0];
  };
  options.draft = one("--draft");
  options.to = values.get("--to");
  options.cc = values.get("--cc");
  options.bcc = values.get("--bcc");
  options.subject = one("--subject");
  options.body = one("--body");
  if (options.draft !== void 0 && !options.draft) usage("Option '--draft' requires a Draft id.");
  if (options.draft !== void 0 && (options.to || options.cc || options.bcc || options.subject !== void 0 || options.body !== void 0 || options.html)) {
    usage("Option '--draft' cannot be combined with compose options.");
  }
  if (options.body === void 0 && options.html) usage("Option '--html' requires '--body'.");
  if (options.draft === void 0 && !options.to && !options.cc && !options.bcc) usage("Provide at least one recipient with '--to', '--cc', or '--bcc'.");
  return options;
}
function preview(message) {
  const addresses = [message.toRecipients, message.ccRecipients, message.bccRecipients].flatMap((group) => group ?? []).map((recipient) => recipient.emailAddress?.address).filter((address) => Boolean(address));
  process.stdout.write(`Recipients: ${addresses.length ? addresses.join(", ") : "(none)"}
Subject: ${message.subject ?? ""}
`);
}
async function runSendCommand(ctx, args) {
  const options = parseArgs2(args);
  if (options.draft !== void 0) {
    const id = messageId(options.draft);
    if (!options.yes) {
      const summary2 = await graphRequest(ctx, "GET", `/me/messages/${encodeURIComponent(id)}`, {
        query: new URLSearchParams({ "$select": "toRecipients,ccRecipients,bccRecipients,subject" })
      });
      preview(summary2 ?? {});
      usage("Sending a Draft requires '--yes'. Run the command again with '--yes' to send it.");
    }
    await graphRequest(ctx, "POST", `/me/messages/${encodeURIComponent(id)}/send`);
    process.stdout.write(`Draft sent: ${id}
`);
    return;
  }
  const message = {
    ...options.to ? { toRecipients: recipients(options.to) } : {},
    ...options.cc ? { ccRecipients: recipients(options.cc) } : {},
    ...options.bcc ? { bccRecipients: recipients(options.bcc) } : {},
    ...options.subject === void 0 ? {} : { subject: options.subject },
    ...options.body === void 0 ? {} : { body: {
      contentType: options.html ? "HTML" : "Text",
      content: await bodyContent(options.body)
    } }
  };
  const summary = message;
  if (!options.yes) {
    preview(summary);
    usage("Sending a Message requires '--yes'. Run the command again with '--yes' to send it.");
  }
  await graphRequest(ctx, "POST", "/me/sendMail", { body: JSON.stringify({ message, saveToSentItems: true }) });
  process.stdout.write("Message sent.\n");
}

// src/commands/reply-forward-commands.ts
function parseArgs3(command, args) {
  const options = { draft: false, yes: false };
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "--draft") {
      options.draft = true;
      continue;
    }
    if (arg === "--yes") {
      options.yes = true;
      continue;
    }
    if (arg === "--body" || arg === "--to") {
      const value = args[index + 1];
      if (value === void 0 || value.startsWith("--") && value !== "-") usage(`Option '${arg}' requires a value.`);
      const property = arg === "--body" ? "body" : "to";
      if (options[property] !== void 0) usage(`Option '${arg}' may be used only once.`);
      options[property] = value;
      index += 1;
      continue;
    }
    if (arg.startsWith("-")) usage(`Unknown ${command} option: ${arg}.`);
    if (options.id !== void 0) usage(`Usage: outlook ${command} <id> --body <text> [--draft|--yes]`);
    options.id = arg;
  }
  if (!options.id) usage(`Usage: outlook ${command} <id> --body <text> [--draft|--yes]`);
  if (options.body === void 0) usage("Option '--body' is required.");
  if (options.to !== void 0 && command !== "forward") usage("Option '--to' is only supported by 'forward'.");
  if (command === "forward" && !options.to) usage("Option '--to' is required for 'forward'.");
  if (options.draft && options.yes) usage("Choose either '--draft' or '--yes', not both.");
  return options;
}
async function runReplyForwardCommand(ctx, command, args) {
  const options = parseArgs3(command, args);
  const id = messageId(options.id);
  if (!options.draft && !options.yes) {
    process.stdout.write(`Would ${command} Message #${id}${options.to ? ` to ${options.to}` : ""}
`);
    usage(`Refusing to ${command} without '--yes' or '--draft'.`);
  }
  const input = options.body;
  const comment = await bodyContent(input);
  const operation = command === "reply-all" ? "replyAll" : command;
  const action = options.draft ? `create${operation[0].toUpperCase()}${operation.slice(1)}` : operation;
  const recipients2 = options.to?.split(",").map((address) => address.trim()).filter(Boolean).map((address) => ({ emailAddress: { address } }));
  if (options.to !== void 0 && !recipients2?.length) usage("Option '--to' must include at least one email address.");
  const body = command === "forward" ? options.draft ? { message: { toRecipients: recipients2 }, comment } : { toRecipients: recipients2, comment } : { comment };
  const response = await graphRequest(ctx, "POST", `/me/messages/${encodeURIComponent(id)}/${action}`, {
    body: JSON.stringify(body)
  });
  if (options.draft) {
    const draftId = response?.id;
    process.stdout.write(draftId ? `Draft created: ${draftId}
` : "Draft created.\n");
  } else process.stdout.write(`${command === "reply-all" ? "Reply-all" : command[0].toUpperCase() + command.slice(1)} sent.
`);
}

// src/commands/write-commands.ts
async function runWriteCommand(ctx, command, args) {
  const options = parseOptions(command, args);
  const ids = options.ids.map(messageId);
  if (!ids.length) usage(`Usage: outlook ${command} <id...>${command === "mark" ? " --read|--unread" : " --yes"}`);
  if (command === "mark" && options.isRead === void 0) usage("Choose either '--read' or '--unread'.");
  if (command === "delete" && !options.yes) {
    process.stdout.write(`Would delete Messages:
${ids.map((id) => `#${id}`).join("\n")}
`);
    throw new CliError("Refusing to delete Messages without '--yes'.", 2);
  }
  const failures = [];
  for (const id of ids) {
    try {
      if (command === "delete") await graphRequest(ctx, "DELETE", `/me/messages/${encodeURIComponent(id)}`);
      else await graphRequest(ctx, "PATCH", `/me/messages/${encodeURIComponent(id)}`, { body: JSON.stringify({ isRead: options.isRead }) });
      process.stdout.write(`${command === "delete" ? "Deleted" : `Marked ${options.isRead ? "read" : "unread"}`} #${id}
`);
    } catch (error) {
      failures.push(id);
      process.stderr.write(`#${id}: ${error instanceof Error ? error.message : String(error)}
`);
    }
  }
  if (failures.length) {
    process.stderr.write(`${failures.length} Message${failures.length === 1 ? "" : "s"} failed.
`);
    process.exitCode = 1;
  }
}
function parseOptions(command, args) {
  const options = { ids: [], yes: false };
  for (const arg of args) {
    if (arg === "--yes") {
      if (command === "mark") usage("Option '--yes' is not accepted by 'mark'; marking is reversible.");
      options.yes = true;
      continue;
    }
    if (arg === "--json") usage("Option '--json' is only supported by read commands.");
    if (command === "mark" && (arg === "--read" || arg === "--unread")) {
      if (options.isRead !== void 0) usage("Choose either '--read' or '--unread', not both.");
      options.isRead = arg === "--read";
      continue;
    }
    if (arg.startsWith("-")) usage(`Unknown ${command} option: ${arg}.`);
    options.ids.push(arg);
  }
  return options;
}

// src/index.ts
var READ_COMMANDS = ["me", "folders", "messages", "message", "contacts", "users", "groups", "members"];
function globals(args) {
  const options = {};
  const rest = [];
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === "-p" || arg === "--profile" || arg === "--client-id" || arg === "--scopes") {
      const value = args[index + 1];
      if (!value || value.startsWith("-")) usage(`Option '${arg}' requires a value.`);
      if (arg === "--client-id") options.clientId = value;
      else if (arg === "--scopes") options.scopes = value;
      else options.profile = value;
      index += 1;
    } else if (arg.startsWith("--profile=")) options.profile = arg.slice("--profile=".length);
    else if (arg.startsWith("--client-id=")) options.clientId = arg.slice("--client-id=".length);
    else if (arg.startsWith("--scopes=")) options.scopes = arg.slice("--scopes=".length);
    else rest.push(arg);
  }
  return { options, rest };
}
async function main(args) {
  if (args.length === 1 && args[0] === "--version") {
    process.stdout.write(`outlook ${package_default.version}
`);
    return;
  }
  if (args.length === 0 || args.includes("--help") || args.includes("-h")) {
    process.stdout.write(`${HELP}
`);
    return;
  }
  const { options, rest } = globals(args);
  const [command, ...commandArgs] = rest;
  const ctx = selectedContext(process.env, options);
  if (command === "auth") return runAuth(ctx, commandArgs[0]);
  if (command === "config") return runConfig(ctx, commandArgs);
  if (command === "api") return runApi(ctx, commandArgs);
  if (command === "draft") return runDraftCommand(ctx, commandArgs);
  if (command === "send") return runSendCommand(ctx, commandArgs);
  if (command === "reply" || command === "reply-all" || command === "forward") return runReplyForwardCommand(ctx, command, commandArgs);
  if (command === "mark" || command === "delete") return runWriteCommand(ctx, command, commandArgs);
  if (command && READ_COMMANDS.includes(command)) return runReadCommand(ctx, command, commandArgs);
  if (!command) usage("Run 'outlook --help' for usage.");
  throw new CliError(`Unknown command: ${command}. Run 'outlook --help' for usage.`, 2);
}
main(process.argv.slice(2)).catch((error) => {
  const code = error instanceof CliError ? error.exitCode : 1;
  const message = error instanceof Error ? error.message : String(error);
  process.stderr.write(`outlook: ${message}
`);
  process.exitCode = code;
});
