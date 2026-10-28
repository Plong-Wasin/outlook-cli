# Outlook CLI

`outlook` is a command-line client for Outlook mail. It calls Microsoft Graph directly, stores sign-in credentials in local Profiles, and prints compact tables by default. It requires Node.js 20 or later and is bundled as a single script.

## Install

```sh
npm install -g --install-links git+https://github.com/Plong-Wasin/outlook-cli.git#release
```

The workflow on `main` typechecks, tests, and builds the CLI before force-publishing the source plus `dist/outlook.js` to `release`. Re-running the install command updates the CLI.

## Sign in

```sh
outlook auth login
outlook auth check
outlook me
```

`auth login` prints a device code and verification URL. Complete sign-in in a browser; credentials are stored in the selected Profile. `auth check` verifies that Profile against Graph. Use `-p <name>` to maintain separate Profiles; the first Profile becomes the default.

```sh
outlook -p personal auth login
outlook -p work auth login
outlook -p personal messages
```

`OUTLOOK_PROFILE` selects a Profile. `OUTLOOK_TOKEN` provides a bearer token for headless use without a Profile. `--client-id` or `OUTLOOK_CLIENT_ID` overrides the OAuth client id. `OUTLOOK_CONFIG` selects another config file. `OUTLOOK_GRAPH_URL` and `OUTLOOK_LOGIN_URL` can point to alternate services, such as test servers. Flags take precedence over environment values, which take precedence over Profile defaults.

## Commands

Run `outlook --help` for the concise reference and `outlook --version` to check the installed version. Commands marked **writes** change mailbox data; use them only when the requested change is explicit. `--yes` is required for Send and delete operations. `mark` changes read state without `--yes`. `api` can write when its method is `POST`, `PUT`, `PATCH`, or `DELETE`.

| Command | Purpose and options |
|---|---|
| `auth login` | Sign in to the selected Profile using a device code. **Writes local credentials.** |
| `auth check` | Verify the selected Profile. |
| `auth logout` | Remove the selected Profile's saved credentials. **Writes local config.** |
| `config show` | Show Profiles and defaults with secrets masked. |
| `me [--json]` | Show the signed-in user. |
| `folders [--cursor <cursor> \| --all] [--json]` | List Mail folders. |
| `messages [--folder <name\|id>] [--unread] [--from <address>] [--since <date>] [--search <text>] [--top <count>] [--cursor <cursor> \| --all] [--json]` | List newest Messages, Inbox by default. `--search` cannot be combined with `--unread`, `--from`, or `--since`; `--folder` also accepts Graph well-known names such as `archive` and `sentitems`. |
| `message <id\|url> [--html] [--json]` | Read one Message as plain text, or HTML with `--html`. |
| `contacts [--top N] [--cursor CURSOR] [--all] [--json]` | List Contacts. |
| `users [--top N] [--cursor CURSOR] [--all] [--json]` | Search directory Users. |
| `groups [--top N] [--cursor CURSOR] [--all] [--json]` | List Groups. |
| `members <group> [--top N] [--cursor CURSOR] [--all] [--json]` | List members of a Group. |
| `mark <id...> --read\|--unread` | Mark Messages read or unread. **Writes mailbox data.** |
| `delete <id...> --yes` | Delete Messages. **Writes mailbox data; requires `--yes`.** |
| `draft [--to <addresses>] [--cc <addresses>] [--bcc <addresses>] [--subject <text>] [--body <text\|@file\|->] [--html]` | Save a Draft. **Writes mailbox data.** |
| `draft --update <id> [fields]` | Update a Draft with the compose fields above. **Writes mailbox data.** |
| `send [compose fields] --yes` | Send a new Message. **Writes mailbox data; requires `--yes`.** |
| `send --draft <id> --yes` | Send a saved Draft. **Writes mailbox data; requires `--yes`.** |
| `reply <id> --body <text> (--yes\|--draft)` | Send a reply or save it as a Draft. **Writes mailbox data.** |
| `reply-all <id> --body <text> (--yes\|--draft)` | Reply to all recipients or save a Draft. **Writes mailbox data.** |
| `forward <id> --to <addresses> --body <text> (--yes\|--draft)` | Forward a Message or save a Draft. **Writes mailbox data.** |
| `api <METHOD> <path> [key=value ...] [--data JSON\|@file\|-] [--header 'Name: value'] [--all]` | Call any Microsoft Graph endpoint. Non-GET methods **write mailbox data** and require explicit user instruction. |

Compose `--body` accepts literal text, `@file`, or `-` for stdin. Use `--html` to send an HTML body. Bulk commands continue after individual failures and report a failure count. `--json` is for read commands only. List commands print a continuation cursor on stderr; pass it with `--cursor` or use `--all` up to the safety cap. IDs may be pasted from Outlook on the web URLs.

## Agent plugin

The Claude Code and Codex plugin manifests are generated from `plugins/outlook-cli/plugin.json`, and both plugins load the same `outlook` skill.

Claude Code:

```text
/plugin marketplace add Plong-Wasin/outlook-cli
/plugin install outlook-cli@outlook-cli
```

Codex:

```sh
codex plugin marketplace add https://github.com/Plong-Wasin/outlook-cli.git
codex plugin add outlook-cli
```

After installing the plugin, install the CLI using the command above. The skill checks the installation and active Profile before use.

## Exit codes and output

- `0`: command succeeded.
- `1`: server, API, or authentication error.
- `2`: usage error or a refused Send/delete without `--yes`.

Data goes to stdout; notes and errors go to stderr. Errors start with `outlook: <message>`. Graph 401, 403, 404, and 429 errors identify the HTTP method and path.

## Development

```sh
pnpm install --frozen-lockfile
pnpm typecheck
pnpm test
pnpm build
pnpm plugin:sync
```

Tests run the built CLI as a child process with a temporary home directory and fake Graph and login servers. They do not contact Microsoft services or read a real config file. `pnpm plugin:check` verifies the generated plugin manifests are current.
