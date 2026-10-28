export const HELP = `Usage: outlook <command> [options]
       outlook --help

Commands:
  auth       Sign in and manage Profiles
  config     Show configuration with secrets masked
  me         Show the signed-in user
  api        Call any Microsoft Graph endpoint
  draft      Create or update a Draft
  folders    List Mail folders
  messages   List Messages in Inbox or another Mail folder
  message    Read one Message by id or Outlook URL
  contacts   List your Contacts
  users      List directory Users
  groups     List Groups
  members    List a Group's members

Commands and options:
  auth login | check | logout
  config show
  me [--json]
  api <METHOD> <path> [key=value ...] [--data JSON|@file|-] [--header 'Name: value'] [--all]
  draft [--to <addresses>] [--cc <addresses>] [--bcc <addresses>] [--subject <text>] [--body <text|@file|->] [--html]
  draft --update <id> [fields]
  folders [--cursor <cursor> | --all] [--json]
  messages [--folder <name|id>] [--unread] [--from <address>] [--since <date>]
           [--search <text>] [--top <count>] [--cursor <cursor> | --all] [--json]
  message <id|url> [--html] [--json]
  contacts | users | groups [--top N] [--cursor CURSOR] [--all] [--json]
  members <group> [--top N] [--cursor CURSOR] [--all] [--json]

Global options:
  -p, --profile <name>  Select a Profile (or set OUTLOOK_PROFILE)
  --client-id <id>      Override the OAuth client id (or set OUTLOOK_CLIENT_ID)
  -h, --help            Show help

Environment:
  OUTLOOK_PROFILE, OUTLOOK_TOKEN, OUTLOOK_CLIENT_ID, OUTLOOK_CONFIG
  OUTLOOK_GRAPH_URL and OUTLOOK_LOGIN_URL (for alternate services and tests)

Exit codes:
  0  Success
  1  Server or authentication error
  2  Usage error`;
