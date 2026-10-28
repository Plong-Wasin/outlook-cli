import { usage } from './errors.js';

export function messageId(input: string): string {
  let id = input;
  if (/^https?:\/\//i.test(input)) {
    let url: URL;
    try { url = new URL(input); }
    catch { usage('Invalid Outlook on the web URL. Copy the Message URL and try again.'); }
    if (!/(^|\.)outlook\.(office\.com|office365\.com|live\.com)$/i.test(url.hostname)) {
      usage('Expected an Outlook on the web URL. Copy the Message URL and try again.');
    }
    const pathParts = url.pathname.split('/').filter(Boolean);
    const idIndex = pathParts.findIndex((part) => part.toLowerCase() === 'id');
    const deeplinkIndex = pathParts.findIndex((part, index) => part.toLowerCase() === 'read' && pathParts[index - 1]?.toLowerCase() === 'deeplink');
    id = (idIndex >= 0 ? pathParts[idIndex + 1] : undefined)
      ?? (deeplinkIndex >= 0 ? pathParts[deeplinkIndex + 1] : undefined)
      ?? url.searchParams.get('ItemID')
      ?? url.searchParams.get('itemid')
      ?? '';
    try { id = decodeURIComponent(id); }
    catch { usage('Invalid Message id in Outlook URL.'); }
  }
  if (!/^[A-Za-z0-9_+\/=\-]+$/.test(id)) usage('Invalid Message id. Use a Graph id or Outlook on the web URL.');
  return id;
}
