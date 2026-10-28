# Outlook CLI

A command-line client for Outlook mail. It talks to the Microsoft Graph API so people and agents can read, draft, and send mail from a shell, with compact output. It replaces the Outlook MCP server.

## Language

**Message**:
An email in the signed-in user's mailbox, in one Mail folder.
_Avoid_: Mail, email item

**Mail folder**:
A named container of Messages in the signed-in user's mailbox, such as Inbox or Sent Items.
_Avoid_: Directory, label

**Draft**:
A Message that is saved but not sent.
_Avoid_: Unsent mail

**Send**:
Delivering a Message to its recipients. It cannot be undone, so it needs explicit confirmation.
_Avoid_: Dispatch, post

**Profile**:
A named set of sign-in credentials selected with `-p`.
_Avoid_: Account, environment, context

**Mailbox**:
The signed-in user's own mail store. Shared mailboxes and other users' mailboxes are out of scope.
_Avoid_: Inbox (that is one Mail folder)
