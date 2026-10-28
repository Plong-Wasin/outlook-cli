import { graphRequest, type Context } from '../commands/auth-commands.js';
import { usage } from './errors.js';
import { graphPage, validateContinuation } from './graph-page.js';

interface MailFolder {
  id?: string;
  displayName?: string;
}

interface FolderCollection {
  value?: MailFolder[];
  '@odata.nextLink'?: string;
}

const PAGE_CAP = 100;
// Graph well-known folder names work in any mailbox language and for nested folders.
const WELL_KNOWN = new Set(['inbox', 'drafts', 'sentitems', 'deleteditems', 'archive', 'junkemail', 'outbox']);

export async function resolveMailFolderId(ctx: Context, value: string): Promise<string> {
  if (WELL_KNOWN.has(value.trim().toLowerCase())) return value.trim().toLowerCase();
  const folders = await listMailFolders(ctx);
  const matches = folders.filter((folder) => folder.id === value || equal(folder.displayName, value));
  if (matches.length === 1 && matches[0]?.id) return matches[0].id;
  fail(value, matches.length > 1, (matches.length ? matches : folders).map((folder) => `${folder.displayName ?? 'Untitled'} (${folder.id ?? 'unknown id'})`));
}

export async function listMailFolders(ctx: Context): Promise<MailFolder[]> {
  const folders: MailFolder[] = [];
  let next = '/me/mailFolders';
  for (let page = 0; page < PAGE_CAP; page += 1) {
    const response = await graphPage(ctx, next) as FolderCollection;
    folders.push(...(response.value ?? []));
    if (!response['@odata.nextLink']) return folders;
    next = response['@odata.nextLink'];
    validateContinuation(next, ctx);
  }
  return folders;
}

function equal(left: string | undefined, right: string): boolean {
  return (left ?? '').trim().toLocaleLowerCase() === right.trim().toLocaleLowerCase();
}

function fail(value: string, ambiguous: boolean, candidates: string[]): never {
  usage(`${ambiguous ? 'Ambiguous' : 'Unknown'} Mail folder '${value}'. Candidates: ${candidates.join(', ') || 'none'}.`);
}
