import { existsSync, readFileSync, writeFileSync } from 'fs';
import { PublicClientApplication, type TokenCacheContext } from '@azure/msal-node';

const TOKEN_CACHE_FILE = '.token-cache.json';
const SCOPES = ['Mail.Read', 'Mail.ReadWrite'];
const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ── Payee rules ───────────────────────────────────────────────────────────────
// Agrega entradas aquí para mapear establecimientos a payees de Actual Budget.

const PAYEE_RULES: Array<{ pattern: RegExp; payee: string }> = [
  { pattern: /UBER/i, payee: 'Uber' },
  { pattern: /COMISARIATO/i, payee: 'Comisariato' },
];

// ── Types ─────────────────────────────────────────────────────────────────────

interface RawEmail {
  id: string;
  subject: string;
  receivedDateTime: string;
  categories: string[];
  body: { content: string };
}

export interface ParsedEmail {
  id: string;
  categories: string[];
  date: string;
  amountUSD: number;
  notes: string;
  payee?: string;
}

export interface FetchOptions {
  subjectRegex: string;
  processedLabel: string;
  fromDate?: string | null;
  limit: number;
}

// ── Auth ──────────────────────────────────────────────────────────────────────

export async function getToken(clientId: string): Promise<string> {
  const pca = new PublicClientApplication({
    auth: { clientId, authority: 'https://login.microsoftonline.com/consumers' },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (ctx: TokenCacheContext) => {
          if (existsSync(TOKEN_CACHE_FILE))
            ctx.tokenCache.deserialize(readFileSync(TOKEN_CACHE_FILE, 'utf8'));
        },
        afterCacheAccess: async (ctx: TokenCacheContext) => {
          if (ctx.cacheHasChanged)
            writeFileSync(TOKEN_CACHE_FILE, ctx.tokenCache.serialize());
        },
      },
    },
  });

  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length > 0) {
    try {
      const result = await pca.acquireTokenSilent({ account: accounts[0], scopes: SCOPES });
      if (result) return result.accessToken;
    } catch {}
  }

  const result = await pca.acquireTokenByDeviceCode({
    scopes: SCOPES,
    deviceCodeCallback: r => {
      console.log(`\nAbre: ${r.verificationUri}\nCódigo: ${r.userCode}\n`);
    },
  });
  if (!result) throw new Error('No se pudo obtener el token de Outlook.');
  return result.accessToken;
}

// ── Graph helpers ─────────────────────────────────────────────────────────────

async function graphGet<T>(
  token: string,
  path: string,
  params: Record<string, string>,
  headers: Record<string, string> = {},
): Promise<T> {
  const url = new URL(`${GRAPH_BASE}${path}`);
  for (const [k, v] of Object.entries(params)) url.searchParams.set(k, v);
  const res = await fetch(url.toString(), {
    headers: { Authorization: `Bearer ${token}`, ...headers },
  });
  if (!res.ok) throw new Error(`Graph GET ${res.status}: ${await res.text()}`);
  return res.json() as Promise<T>;
}

async function graphPatch(token: string, path: string, body: unknown): Promise<void> {
  const res = await fetch(`${GRAPH_BASE}${path}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`Graph PATCH ${res.status}: ${await res.text()}`);
}

// ── Parsing ───────────────────────────────────────────────────────────────────

function parseAmount(subject: string, subjectRegex: string): number | null {
  const match = subject.match(new RegExp(subjectRegex, 'i'));
  if (!match) return null;
  const last = match.slice(1).filter(g => g && /[\d.]+/.test(g)).at(-1);
  return last ? parseFloat(last) : null;
}

function parseEstablishment(body: string): string | null {
  const text = body
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(?:p|div|tr)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/[ \t]+/g, ' ');
  const match = text.match(/Establecimiento:\s*([^\n\r]+)/i);
  return match ? match[1].trim() : null;
}

function matchPayee(establishment: string): string | undefined {
  return PAYEE_RULES.find(r => r.pattern.test(establishment))?.payee;
}

// ── Public API ────────────────────────────────────────────────────────────────

export async function fetchUnprocessedEmails(
  token: string,
  opts: FetchOptions,
): Promise<ParsedEmail[]> {
  const params: Record<string, string> = {
    $orderby: 'receivedDateTime desc',
    $select: 'id,subject,receivedDateTime,categories,body',
    ...(opts.fromDate
      ? { $filter: `receivedDateTime ge ${opts.fromDate}T00:00:00Z`, $top: '999' }
      : { $top: String(opts.limit) }),
  };

  const { value: raw } = await graphGet<{ value: RawEmail[] }>(
    token,
    '/me/mailFolders/inbox/messages',
    params,
    { Prefer: 'outlook.body-content-type="text"' },
  );

  const subjectRegex = new RegExp(opts.subjectRegex, 'i');

  return raw
    .filter(e => subjectRegex.test(e.subject) && !e.categories.includes(opts.processedLabel))
    .flatMap(e => {
      const amountUSD = parseAmount(e.subject, opts.subjectRegex);
      if (amountUSD === null) return [];
      const notes = (parseEstablishment(e.body.content) ?? e.subject).toLowerCase();
      return [{ id: e.id, categories: e.categories, date: e.receivedDateTime.split('T')[0], amountUSD, notes, payee: matchPayee(notes) }];
    });
}

export async function markProcessed(
  token: string,
  id: string,
  existingCategories: string[],
  label: string,
): Promise<void> {
  await graphPatch(token, `/me/messages/${id}`, {
    categories: [...existingCategories, label],
  });
}
