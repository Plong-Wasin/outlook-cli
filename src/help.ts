export const HELP = `Usage: outlook <command> [options]
       outlook --help

Commands:
  auth       Sign in and manage Profiles
  config     Show configuration with secrets masked
  me         Show the signed-in user
  message    Read one Message by id or Outlook URL

Commands and options:
  auth login | check | logout
  config show
  me [--json]
  message <id|url> [--html] [--json]

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
