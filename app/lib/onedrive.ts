import type { AccountInfo, PublicClientApplication } from '@azure/msal-browser';

export type CloudDocument = {
  id: string;
  name: string;
  eTag?: string;
  lastModifiedDateTime?: string;
  size?: number;
};

const scopes = ['Files.ReadWrite.AppFolder', 'User.Read'];
const graphRoot = 'https://graph.microsoft.com/v1.0';
let client: PublicClientApplication | null = null;

function clientId() {
  return process.env.NEXT_PUBLIC_MICROSOFT_CLIENT_ID?.trim() ?? '';
}

export function isOneDriveConfigured() {
  return Boolean(clientId());
}

async function getClient() {
  if (client) return client;
  if (!clientId()) throw new Error('尚未配置 Microsoft Application Client ID');
  const { BrowserCacheLocation, PublicClientApplication } = await import('@azure/msal-browser');
  client = new PublicClientApplication({
    auth: {
      clientId: clientId(),
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: `${window.location.origin}/auth/callback`,
    },
    cache: {
      cacheLocation: BrowserCacheLocation.SessionStorage,
      storeAuthStateInCookie: true,
    },
  });
  await client.initialize();
  const existing = client.getAllAccounts()[0];
  if (existing && !client.getActiveAccount()) client.setActiveAccount(existing);
  return client;
}

export async function restoreOneDriveSession(): Promise<AccountInfo | null> {
  if (!isOneDriveConfigured()) return null;
  const instance = await getClient();
  return instance.getActiveAccount() ?? instance.getAllAccounts()[0] ?? null;
}

export async function connectOneDrive(): Promise<AccountInfo> {
  const instance = await getClient();
  const result = await instance.loginPopup({ scopes, prompt: 'select_account' });
  instance.setActiveAccount(result.account);
  return result.account;
}

export async function disconnectOneDrive() {
  const instance = await getClient();
  const account = instance.getActiveAccount();
  if (account) await instance.logoutPopup({ account, postLogoutRedirectUri: window.location.origin });
}

async function accessToken() {
  const instance = await getClient();
  const account = instance.getActiveAccount() ?? instance.getAllAccounts()[0];
  if (!account) throw new Error('请先连接 OneDrive');
  try {
    return (await instance.acquireTokenSilent({ account, scopes })).accessToken;
  } catch {
    return (await instance.acquireTokenPopup({ account, scopes })).accessToken;
  }
}

async function graph(path: string, init: RequestInit = {}) {
  const token = await accessToken();
  const response = await fetch(`${graphRoot}${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, ...init.headers },
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => null) as { error?: { message?: string } } | null;
    throw new Error(payload?.error?.message || `OneDrive 请求失败 (${response.status})`);
  }
  return response;
}

export async function listDocuments(): Promise<CloudDocument[]> {
  const response = await graph('/me/drive/special/approot/children?$select=id,name,eTag,lastModifiedDateTime,size,file');
  const payload = await response.json() as { value: Array<CloudDocument & { file?: unknown }> };
  return payload.value
    .filter((item) => item.file && item.name.toLowerCase().endsWith('.md'))
    .sort((a, b) => (b.lastModifiedDateTime ?? '').localeCompare(a.lastModifiedDateTime ?? ''));
}

export async function readDocument(id: string) {
  const response = await graph(`/me/drive/items/${encodeURIComponent(id)}/content`);
  return response.text();
}

export async function saveDocument(name: string, content: string, eTag?: string): Promise<CloudDocument> {
  const safeName = name.toLowerCase().endsWith('.md') ? name : `${name}.md`;
  const headers: Record<string, string> = { 'Content-Type': 'text/markdown; charset=utf-8' };
  if (eTag) headers['If-Match'] = eTag;
  const response = await graph(
    `/me/drive/special/approot:/${encodeURIComponent(safeName)}:/content`,
    { method: 'PUT', headers, body: content },
  );
  return response.json() as Promise<CloudDocument>;
}
