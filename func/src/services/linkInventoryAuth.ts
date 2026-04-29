import { createSign, createHash, randomBytes, type KeyObject, createPrivateKey } from "node:crypto";
import { ManagedIdentityCredential } from "@azure/identity";
import { SPO_ORIGIN } from "./spoTokenProvider.js";

/**
 * Permission enforcement for the link-inventory endpoints.
 *
 * Two layers:
 *
 * 1. Group-based admin gate. Triggering a tenant scan or any write
 *    operation requires membership in the Entra group identified by
 *    LINK_INVENTORY_ADMIN_GROUP_ID. Checked via the `groups` claim in
 *    the user's bearer token (fast path) with a Graph fallback for the
 *    "groups overage" case (rare — only when a user has >200 groups).
 *
 * 2. Per-site SP permission enforcement. For every site whose data the
 *    user wants to read, we exchange the user's bearer token via OBO
 *    for an SP-audience token *as the user*, then call SP REST
 *    `/_api/web/effectiveBasePermissions`. The returned base permission
 *    mask tells us whether the user has ViewListItems / EditListItems.
 *    Sites where the user lacks read are silently dropped from the
 *    response. Sites where they lack write block any write operation.
 *
 * Token exchange uses the OBO grant
 * (`urn:ietf:params:oauth:grant-type:jwt-bearer`). The
 * `client_assertion` is produced one of two ways depending on
 * `LINK_INVENTORY_AUTH_MODE`:
 *
 *   - `cert` (default) — sign a short-lived JWT locally with the
 *     private key from `LINK_INVENTORY_CLIENT_CERT_PEM_BASE64`.
 *   - `federation` — fetch an MI-issued token for the
 *     `api://AzureADTokenExchange` audience and use that as the
 *     assertion. Only valid when MI and Entra app live in the same
 *     tenant (cross-tenant Entra-to-Entra federation is blocked by
 *     AADSTS700236).
 *
 * Both modes produce `appidacr=2` tokens, which SP REST requires
 * (secret-issued tokens with `appidacr=1` are rejected by the search
 * endpoint). The user's bearer token is the `assertion` parameter.
 */

const AUTH_MODE = (process.env.LINK_INVENTORY_AUTH_MODE ?? "cert").toLowerCase();
const TENANT_ID = process.env.LINK_INVENTORY_TENANT_ID ?? "";
const CLIENT_ID = process.env.LINK_INVENTORY_CLIENT_ID ?? "";
const CLIENT_CERT_B64 = process.env.LINK_INVENTORY_CLIENT_CERT_PEM_BASE64 ?? "";
const ADMIN_GROUP_ID = (process.env.LINK_INVENTORY_ADMIN_GROUP_ID ?? "").toLowerCase();

const TOKEN_ENDPOINT = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
const SP_OBO_READ_SCOPE = `${SPO_ORIGIN}/AllSites.Read`;
const SP_OBO_WRITE_SCOPE = `${SPO_ORIGIN}/AllSites.Write`;
const GRAPH_OBO_SCOPE = "https://graph.microsoft.com/.default";

// Re-export for backwards compat with the existing read-side code
const SP_OBO_SCOPE = SP_OBO_READ_SCOPE;

// In-memory caches per Function instance. Both are bounded TTL — short
// enough to reflect group-membership changes within a few minutes,
// long enough to avoid spamming Graph/SP on every poll from the UI.
const PERMISSION_CACHE_TTL_MS = 5 * 60 * 1000;
// Negative permission results (canRead=false / canWrite=false) cache
// for a much shorter window so a transient OBO failure or a freshly
// granted permission self-heals on the next call. Without this, a
// stale "no access" verdict can persist for the full 5-minute TTL,
// blocking the user even after the underlying issue is resolved.
const PERMISSION_CACHE_TTL_NEGATIVE_MS = 30 * 1000;
const TOKEN_CACHE_TTL_BUFFER_MS = 60 * 1000;

interface CachedUserToken {
  token: string;
  expiresAt: number;
  scope: string;
}

interface CachedPermission {
  expiresAt: number;
  canRead: boolean;
  canWrite: boolean;
}

interface CachedAdminCheck {
  expiresAt: number;
  isAdmin: boolean;
}

const userTokenCache = new Map<string, CachedUserToken>();
const permissionCache = new Map<string, CachedPermission>();
const adminCache = new Map<string, CachedAdminCheck>();

/**
 * Cert + private key parsed once per Function instance from the base64
 * PEM in app settings. We hold the parsed private key as a KeyObject
 * and the cert thumbprint base64url-encoded for the JWT `x5t` header.
 * Only populated when `AUTH_MODE === 'cert'`.
 */
let parsedKey: { privateKey: KeyObject; x5t: string } | undefined;
function getParsedKey(): { privateKey: KeyObject; x5t: string } {
  if (parsedKey) return parsedKey;
  if (!CLIENT_CERT_B64) {
    throw new AuthError(500, "LINK_INVENTORY_CLIENT_CERT_PEM_BASE64 must be set when auth mode is 'cert'");
  }
  const pem = Buffer.from(CLIENT_CERT_B64, "base64").toString("utf8");
  const certMatch = pem.match(/-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/);
  if (!certMatch) {
    throw new AuthError(500, "LINK_INVENTORY_CLIENT_CERT_PEM_BASE64 missing CERTIFICATE block");
  }
  const certDer = Buffer.from(
    certMatch[0].replace(/-----BEGIN CERTIFICATE-----|-----END CERTIFICATE-----|\s/g, ""),
    "base64",
  );
  // deepcode ignore InsecureHash: x5t is a base64url-encoded SHA-1 cert thumbprint per RFC 7515 §4.1.7; Entra rejects client_assertion JWTs that use any other algorithm here.
  const x5t = createHash("sha1").update(certDer).digest("base64url");
  const privateKey = createPrivateKey({ key: pem, format: "pem" });
  parsedKey = { privateKey, x5t };
  return parsedKey;
}

/**
 * Build and sign a short-lived JWT for use as `client_assertion` (cert
 * mode). Same shape as MSAL produces, RS256 signed with the private
 * key matching the cert uploaded to the Entra app.
 */
function signClientAssertion(): string {
  const { privateKey, x5t } = getParsedKey();
  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT", x5t };
  const payload = {
    aud: TOKEN_ENDPOINT,
    iss: CLIENT_ID,
    sub: CLIENT_ID,
    jti: randomBytes(16).toString("hex"),
    nbf: now,
    exp: now + 600,
  };
  const enc = (obj: object): string => Buffer.from(JSON.stringify(obj)).toString("base64url");
  const signingInput = `${enc(header)}.${enc(payload)}`;
  const sig = createSign("RSA-SHA256").update(signingInput).sign(privateKey).toString("base64url");
  return `${signingInput}.${sig}`;
}

/**
 * Federation-mode assertion: ask the Function App's MI for a token
 * scoped to `api://AzureADTokenExchange`. Entra accepts that token
 * as the `client_assertion` for the OBO call.
 */
let miCredential: ManagedIdentityCredential | undefined;
async function getMiAssertion(): Promise<string> {
  if (!miCredential) miCredential = new ManagedIdentityCredential();
  const tok = await miCredential.getToken("api://AzureADTokenExchange/.default");
  if (!tok) throw new AuthError(500, "Managed identity returned no AzureADTokenExchange token");
  return tok.token;
}

async function getClientAssertion(): Promise<string> {
  if (AUTH_MODE === "federation") return getMiAssertion();
  if (AUTH_MODE === "cert") return signClientAssertion();
  throw new AuthError(
    500,
    `Unsupported LINK_INVENTORY_AUTH_MODE='${AUTH_MODE}' (expected 'cert' or 'federation')`,
  );
}

interface UserPrincipal {
  /** oid (object id) claim — stable user identifier */
  userId: string;
  /** preferred_username / upn claim, for logging */
  upn?: string;
  /** groups claim, when present in the token */
  groups: string[];
  /** True if the token has a `_claim_names.groups` overage marker */
  hasGroupsOverage: boolean;
  /** Raw bearer token (we need it for OBO exchanges) */
  rawToken: string;
}

export class AuthError extends Error {
  constructor(public readonly status: number, message: string) {
    super(message);
  }
}

// Re-export the principal type for cross-module use (e.g. the writer
// needs UserPrincipal to call getSpoUserToken).
export type { UserPrincipal };

/**
 * Decode a JWT without verifying the signature. We don't need to verify
 * here because the token came in over HTTPS from a trusted SPFx caller
 * that obtained it from Entra (Easy Auth on the Function App handles
 * the upstream signature verification when we add it). The claims we
 * read are not used for authorization decisions on their own — every
 * authorization decision is paired with an OBO call back to the
 * authoritative service (SP or Graph), which fails the request if the
 * token isn't valid.
 *
 * For now (function-key auth in front of these endpoints), the bearer
 * token is supplied by the caller's SP-context aadHttpClient and
 * Entra-issued. When we wire EasyAuth, the function host will
 * additionally verify the signature for us via X-MS-CLIENT-PRINCIPAL.
 */
function decodeJwt(token: string): Record<string, unknown> {
  const parts = token.split(".");
  if (parts.length !== 3) throw new AuthError(401, "Malformed bearer token");
  try {
    const payload = Buffer.from(parts[1].replace(/-/g, "+").replace(/_/g, "/"), "base64").toString("utf8");
    return JSON.parse(payload) as Record<string, unknown>;
  } catch {
    throw new AuthError(401, "Malformed bearer token payload");
  }
}

/**
 * Pull the calling user's principal info from the Authorization header.
 * Throws AuthError(401) if the header is missing/malformed.
 */
export function parseUserPrincipal(authHeader: string | null): UserPrincipal {
  if (!authHeader || !authHeader.toLowerCase().startsWith("bearer ")) {
    throw new AuthError(401, "Missing or invalid Authorization header (expected 'Bearer ...')");
  }
  const rawToken = authHeader.slice(7).trim();
  const claims = decodeJwt(rawToken);

  const userId = (claims.oid as string | undefined) ?? (claims.sub as string | undefined);
  if (!userId) throw new AuthError(401, "Token missing oid/sub claim");

  const groupsClaim = claims.groups;
  const groups = Array.isArray(groupsClaim) ? (groupsClaim as string[]).map((g) => g.toLowerCase()) : [];
  const hasGroupsOverage = !!(claims._claim_names && (claims._claim_names as Record<string, unknown>).groups);

  return {
    userId,
    upn: (claims.upn as string | undefined) ?? (claims.preferred_username as string | undefined),
    groups,
    hasGroupsOverage,
    rawToken,
  };
}

/**
 * Exchange the user's bearer token for a downstream-audience token via
 * OBO. The result is cached per (userId, scope) until ~1 minute before
 * expiry.
 */
async function getOboToken(user: UserPrincipal, scope: string): Promise<string> {
  const cacheKey = `${user.userId}:${scope}`;
  const now = Date.now();
  const cached = userTokenCache.get(cacheKey);
  if (cached && cached.expiresAt - TOKEN_CACHE_TTL_BUFFER_MS > now) return cached.token;

  if (!TENANT_ID || !CLIENT_ID) {
    throw new AuthError(500, "LINK_INVENTORY_TENANT_ID / LINK_INVENTORY_CLIENT_ID not configured");
  }

  const clientAssertion = await getClientAssertion();

  const body = new URLSearchParams({
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    client_id: CLIENT_ID,
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    client_assertion: clientAssertion,
    assertion: user.rawToken,
    scope: scope,
    requested_token_use: "on_behalf_of",
  });

  const res = await fetch(TOKEN_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString(),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new AuthError(401, `OBO exchange failed (${res.status}): ${text.slice(0, 300)}`);
  }
  const json = (await res.json()) as { access_token: string; expires_in: number };
  const expiresAt = now + json.expires_in * 1000;
  userTokenCache.set(cacheKey, { token: json.access_token, expiresAt, scope });
  return json.access_token;
}

/**
 * True if the calling user is in the configured admin group. Fast path:
 * read the `groups` claim. Fallback: query Graph via OBO when the token
 * has a groups overage indicator.
 *
 * Result cached for 5 minutes per user.
 */
export async function isAdmin(user: UserPrincipal): Promise<boolean> {
  if (!ADMIN_GROUP_ID) {
    // Misconfigured — fail closed.
    throw new AuthError(500, "LINK_INVENTORY_ADMIN_GROUP_ID not configured");
  }

  const cacheKey = user.userId;
  const now = Date.now();
  const cached = adminCache.get(cacheKey);
  if (cached && cached.expiresAt > now) return cached.isAdmin;

  let isMember = false;

  if (!user.hasGroupsOverage && user.groups.length > 0) {
    isMember = user.groups.includes(ADMIN_GROUP_ID);
  } else {
    // Overage path (or token didn't carry groups for some reason) —
    // call Graph as the user.
    try {
      const graphToken = await getOboToken(user, GRAPH_OBO_SCOPE);
      const url = `https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.group?$select=id&$top=999`;
      const res = await fetch(url, { headers: { Authorization: `Bearer ${graphToken}` } });
      if (res.ok) {
        const data = (await res.json()) as { value: Array<{ id: string }> };
        isMember = data.value.some((g) => g.id.toLowerCase() === ADMIN_GROUP_ID);
      } else {
        // If Graph itself rejects the user, treat as non-admin
        isMember = false;
      }
    } catch {
      isMember = false;
    }
  }

  adminCache.set(cacheKey, { expiresAt: now + PERMISSION_CACHE_TTL_MS, isAdmin: isMember });
  return isMember;
}

/**
 * Throw AuthError(403) if the user is not in the admin group.
 */
export async function requireAdmin(user: UserPrincipal): Promise<void> {
  const ok = await isAdmin(user);
  if (!ok) {
    throw new AuthError(403, "This operation requires membership in the configured admin group");
  }
}

/**
 * SharePoint base permission bits we care about. The full mask is a
 * 64-bit value SP returns split into High/Low — we only need the bits
 * that fit in Low (the first 32). ViewListItems and EditListItems both
 * live in Low.
 *
 * Bit positions per SPBasePermissions enum:
 *   ViewListItems = 1 (1 << 0)
 *   EditListItems = 4 (1 << 2)
 */
const PERM_VIEW_LIST_ITEMS = 0x1;
const PERM_EDIT_LIST_ITEMS = 0x4;

/**
 * SP REST `effectiveBasePermissions` response shape (with `odata=nometadata`).
 * The Low/High pair is at the root of the response — there is no
 * `EffectiveBasePermissions` wrapper unless you use `odata=verbose`.
 */
interface SpEffectivePermissionsResponse {
  Low?: string;
  High?: string;
  // Older / verbose form, kept as a fallback
  EffectiveBasePermissions?: { Low: string; High: string };
}

/**
 * Check whether the user can read / write items on a given site.
 *
 * **CRITICAL: scope matters.** SharePoint's `effectiveBasePermissions`
 * endpoint masks the response based on the OBO token's scope. A token
 * acquired with `AllSites.Read` will have all write-related permission
 * bits stripped from the response, even if the user actually has
 * EditListItems on the site. To check for write access correctly, we
 * MUST request the token with `AllSites.Write`. Discovered the hard
 * way debugging "the tool says I don't have edit access but I clearly
 * do" — turned out the user had ManageWeb and EditListItems on the
 * site, but the read-scoped token returned `low=0x8231061` (ManageWeb
 * set, EditListItems unset) which is the read-masked view.
 *
 * The `purpose` parameter selects the scope:
 *   - `'read'` (default) — uses AllSites.Read; canWrite is unreliable
 *   - `'write'` — uses AllSites.Write; both canRead and canWrite are
 *     reliable. Slightly more expensive (a separate token in the
 *     OBO cache) but the only correct check for the write path.
 *
 * Caches positive results for 5 minutes (fast UX) and negative
 * results for 30 seconds (so transient OBO failures and freshly
 * granted permissions self-heal quickly). Cache key includes purpose
 * so read-purpose and write-purpose checks don't collide.
 */
export async function getSitePermissions(
  user: UserPrincipal,
  sitePath: string,
  context?: { log?: (msg: string) => void; warn?: (msg: string) => void; error?: (msg: string) => void },
  purpose: 'read' | 'write' = 'read',
): Promise<{ canRead: boolean; canWrite: boolean }> {
  const log = (msg: string): void => { if (context?.log) context.log(msg); };
  const warn = (msg: string): void => { if (context?.warn) context.warn(msg); };
  const cacheKey = `${user.userId}:${sitePath}:${purpose}`;
  const now = Date.now();
  const cached = permissionCache.get(cacheKey);
  if (cached && cached.expiresAt > now) {
    log(`[getSitePermissions/${purpose}] ${user.upn ?? user.userId} ${sitePath}: cache hit canRead=${cached.canRead} canWrite=${cached.canWrite}`);
    return { canRead: cached.canRead, canWrite: cached.canWrite };
  }

  let canRead = false;
  let canWrite = false;
  let outcome = "unknown";

  // Defensive: any failure on a per-site permission check (network
  // error, malformed URL, OBO failure, missing site, transient SP
  // error) is treated as "no access" — but cached only briefly so
  // the user can retry. The fail-closed default protects against
  // serving data to users who lack permission, while the short TTL
  // protects against transient errors blocking legitimate access.
  try {
    const scope = purpose === 'write' ? SP_OBO_WRITE_SCOPE : SP_OBO_SCOPE;
    const spUserToken = await getOboToken(user, scope);
    const url = `${SPO_ORIGIN}${sitePath}/_api/web/effectiveBasePermissions`;
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${spUserToken}`,
        Accept: "application/json;odata=nometadata",
      },
    });
    if (res.ok) {
      const data = (await res.json()) as SpEffectivePermissionsResponse;
      const lowStr = data.Low ?? data.EffectiveBasePermissions?.Low;
      const low = lowStr ? parseInt(lowStr, 10) : NaN;
      if (!Number.isNaN(low)) {
        canRead = (low & PERM_VIEW_LIST_ITEMS) !== 0;
        canWrite = (low & PERM_EDIT_LIST_ITEMS) !== 0;
        outcome = `200 low=0x${low.toString(16)} canRead=${canRead} canWrite=${canWrite}`;
      } else {
        outcome = `200 but unparseable Low: ${JSON.stringify(data).slice(0, 200)}`;
        warn(`[getSitePermissions/${purpose}] ${sitePath}: ${outcome}`);
      }
    } else {
      const body = (await res.text()).slice(0, 200);
      outcome = `${res.status} ${body}`;
      warn(`[getSitePermissions/${purpose}] ${sitePath}: HTTP ${res.status} — ${body}`);
    }
  } catch (err) {
    outcome = `threw: ${(err as Error).message}`;
    warn(`[getSitePermissions/${purpose}] ${sitePath}: ${outcome}`);
  }

  log(`[getSitePermissions/${purpose}] ${user.upn ?? user.userId} ${sitePath}: ${outcome}`);
  // Positive results cache for 5 min, negatives for 30s.
  const ttl = (canRead || canWrite) ? PERMISSION_CACHE_TTL_MS : PERMISSION_CACHE_TTL_NEGATIVE_MS;
  permissionCache.set(cacheKey, { expiresAt: now + ttl, canRead, canWrite });
  return { canRead, canWrite };
}

/**
 * Cache of the user's full visible-site set, computed once via SP
 * search and reused for ~1 hour. The cache is per Function instance
 * (lost on deploys/restarts) and keyed by user oid only — search
 * results don't depend on what we're filtering.
 */
interface CachedUserSiteSet {
  expiresAt: number;
  siteSet: Set<string>;
}
const USER_SITE_SET_TTL_MS = 60 * 60 * 1000; // 1 hour
const userSiteSetCache = new Map<string, CachedUserSiteSet>();

interface SearchResultRow {
  Cells: Array<{ Key: string; Value: string }>;
}
interface SearchResponse {
  PrimaryQueryResult?: {
    RelevantResults?: {
      RowCount: number;
      TotalRows: number;
      Table: { Rows: SearchResultRow[] };
    };
  };
}

/**
 * Enumerate every SPO site the calling user can see, via SP REST
 * search executed under the user's OBO token. SP's search index is
 * ACL-trimmed, so this returns the user's visible-site set with a
 * single API call instead of N per-site `effectiveBasePermissions`
 * checks.
 *
 * Limitations to be aware of:
 *   - Search ACL trim has independent indices for content vs ACLs;
 *     newly-granted access takes minutes to propagate. For our
 *     purposes (link inventory filtering) this is acceptable.
 *   - Sites not indexed (e.g. sites with "do not index" set) won't
 *     appear in search even if the user has access. Rare in practice.
 *   - Returns server-relative paths normalized to lowercase for the
 *     match-set, so callers should compare lowercased.
 */
async function enumerateUserVisibleSites(user: UserPrincipal): Promise<Set<string>> {
  const cacheKey = user.userId;
  const now = Date.now();
  const cached = userSiteSetCache.get(cacheKey);
  if (cached && cached.expiresAt > now) return cached.siteSet;

  const token = await getOboToken(user, SP_OBO_SCOPE);
  const out = new Set<string>();
  const baseUrl =
    `${SPO_ORIGIN}/_api/search/query` +
    `?querytext='contentclass:STS_Site'` +
    `&trimduplicates=false` +
    `&rowlimit=500` +
    `&selectproperties='Path'`;

  let startRow = 0;
  let calls = 0;
  const maxCalls = 50;
  const tenantHost = new URL(SPO_ORIGIN).hostname.toLowerCase();

  while (calls < maxCalls) {
    calls += 1;
    const res = await fetch(`${baseUrl}&startrow=${startRow}`, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
      },
    });
    if (!res.ok) {
      // Search may be unavailable for some users (rare). Cache an
      // empty set briefly so we don't hammer search on retry, and
      // let the caller fall back to a permissive default.
      userSiteSetCache.set(cacheKey, { expiresAt: now + 60_000, siteSet: out });
      return out;
    }
    const data = (await res.json()) as SearchResponse;
    const rels = data.PrimaryQueryResult?.RelevantResults;
    if (!rels) break;
    const rows = rels.Table?.Rows ?? [];
    if (rows.length === 0) break;
    for (const row of rows) {
      const cell = row.Cells.find((c) => c.Key === "Path");
      if (!cell || !cell.Value) continue;
      try {
        const u = new URL(cell.Value);
        if (u.hostname.toLowerCase() !== tenantHost) continue;
        out.add(u.pathname.replace(/\/$/, "").toLowerCase() || "/");
      } catch { /* skip malformed */ }
    }
    startRow += rows.length;
    if (startRow >= rels.TotalRows) break;
  }

  userSiteSetCache.set(cacheKey, { expiresAt: now + USER_SITE_SET_TTL_MS, siteSet: out });
  return out;
}

/**
 * Filter a list of site paths to those the user can read. Used by the
 * results endpoint to drop sites the user has no rights to before
 * returning the response.
 *
 * Strategy: one search call per page load (cached for 1 hour), instead
 * of N per-site `effectiveBasePermissions` checks. SP search returns
 * an ACL-trimmed list of sites the user can see, which is exactly the
 * filter set we need.
 *
 * Fallback: if search returns an empty set (e.g. search index
 * unavailable for this user), we fall back to per-site OBO checks
 * with concurrency 8 — slower but more accurate.
 */
export async function filterReadableSites(
  user: UserPrincipal,
  sites: string[],
): Promise<string[]> {
  // Admin bypass: Redirect Manager Admins see every site the function
  // itself can read. The search-based ACL trim has lag — sites where the
  // user was recently added as owner can be hidden for hours until the
  // index catches up. Admins shouldn't be subject to that lag because
  // they're the ones who triggered the scan and need full visibility.
  try {
    if (await isAdmin(user)) return [...sites];
  } catch {
    // isAdmin failure shouldn't block the read path — fall through to
    // the normal filter logic.
  }

  // Fast path: SP search ACL-trim
  let userSites: Set<string>;
  try {
    userSites = await enumerateUserVisibleSites(user);
  } catch {
    userSites = new Set();
  }

  if (userSites.size > 0) {
    const allowed = sites.filter((s) => userSites.has(s.toLowerCase()));
    return allowed;
  }

  // Fallback: search returned nothing. Per-site OBO checks at low
  // concurrency. Slow but unblocked when search is misbehaving.
  const PERM_CHECK_CONCURRENCY = 8;
  const allowed: string[] = [];
  let cursor = 0;
  async function worker(): Promise<void> {
    while (true) {
      const idx = cursor++;
      if (idx >= sites.length) return;
      const s = sites[idx];
      try {
        const perms = await getSitePermissions(user, s);
        if (perms.canRead) allowed.push(s);
      } catch { /* ignore */ }
    }
  }
  const workers: Promise<void>[] = [];
  for (let i = 0; i < Math.min(PERM_CHECK_CONCURRENCY, sites.length); i++) {
    workers.push(worker());
  }
  await Promise.all(workers);
  allowed.sort();
  return allowed;
}

/**
 * Get an SP-audience bearer token *as the user* with write scope. Used
 * by the find-and-replace writer so PATCH operations execute under the
 * user's identity (correct audit trail) and use the user's actual SP
 * permissions (no Sites.Selected per-site grants required).
 *
 * Requires the Entra app to have delegated `AllSites.Write` consented,
 * which it does (granted via Graph in this session).
 */
export async function getSpoUserWriteToken(user: UserPrincipal): Promise<string> {
  return getOboToken(user, SP_OBO_WRITE_SCOPE);
}

/**
 * Get an SP-audience bearer token *as the user* with read scope. Used
 * for tenant-wide site enumeration via SP REST search — that endpoint
 * rejects app-only tokens ("Unsupported app only token") regardless of
 * Sites.Read.All consent, so enumeration must run as the user. The
 * search index is ACL-trimmed, so the result is naturally filtered to
 * sites the calling user can see.
 */
export async function getSpoUserReadToken(user: UserPrincipal): Promise<string> {
  return getOboToken(user, SP_OBO_READ_SCOPE);
}
