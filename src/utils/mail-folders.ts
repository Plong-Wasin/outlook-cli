import { graphRequest, type Context } from '../commands/auth-commands.js';
import { usage } from './errors.js';

interface MailFolder {
  id?: string;
  displayName?: string;
}

interface FolderCollection {
  value?: MailFolder[];
  '@odata.nextLink'?: string;
}

const PAGE_CAP = 100;

export async function resolveMailFolderId(ctx: Context, value: string): Promise<string> {
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

function graphPage(ctx: Context, url: string): Promise<unknown> {
  return /^https?:\/\//i.test(url)
    ? graphRequest(ctx, 'GET', url, { rawUrl: url })
    : graphRequest(ctx, 'GET', url);
}

function validateContinuation(link: string, ctx: Context): void {
  const baseValue = ctx.env.OUTLOOK_GRAPH_URL || 'https://graph.microsoft.com/v1.0';
  let url: URL;
  try { url = new URL(link, baseValue); }
  catch { usage('Graph returned an invalid continuation URL while resolving a Mail folder.'); }
  const base = new URL(baseValue);
  const basePath = base.pathname.replace(/\/$/, '');
  const inBasePath = !basePath || url.pathname === basePath || url.pathname.startsWith(`${basePath}/`);
  if (url.origin !== base.origin || !inBasePath) usage('Graph returned a continuation URL outside the configured Graph URL.');
}

function equal(left: string | undefined, right: string): boolean {
  return (left ?? '').trim().toLocaleLowerCase() === right.trim().toLocaleLowerCase();
}

function fail(value: string, ambiguous: boolean, candidates: string[]): never {
  usage(`${ambiguous ? 'Ambiguous' : 'Unknown'} Mail folder '${value}'. Candidates: ${candidates.join(', ') || 'none'}.`);
}
