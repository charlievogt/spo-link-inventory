import {
  ClientAssertionCredential,
  ClientCertificateCredential,
  ManagedIdentityCredential,
  type TokenCredential,
} from "@azure/identity";

/**
 * Token provider for the Link Inventory feature.
 *
 * Two auth modes, selected via `LINK_INVENTORY_AUTH_MODE`:
 *
 *   - `cert` (default) — client-credentials flow with a client
 *     certificate on the Entra app. PEM (cert + private key
 *     concatenated) is base64-encoded into
 *     `LINK_INVENTORY_CLIENT_CERT_PEM_BASE64`. Works any-tenant.
 *
 *   - `federation` — Function App's system-assigned managed identity
 *     federates into the Entra app. No credential at rest. Only works
 *     when the MI and the Entra app live in the SAME Entra tenant
 *     (cross-tenant Entra-to-Entra federation is blocked by Microsoft
 *     under AADSTS700236).
 *
 * Both modes produce `appidacr=2` tokens, which SharePoint REST
 * (notably the search endpoint) requires — `client_secret` flows
 * produce `appidacr=1` tokens that SP rejects with "Unsupported app
 * only token". That's why this module does NOT support secret-based
 * auth.
 */

import { getTenantOrigin } from "./config.js";

const AUTH_MODE = (process.env.LINK_INVENTORY_AUTH_MODE ?? "cert").toLowerCase();
const TENANT_ID = process.env.LINK_INVENTORY_TENANT_ID ?? "";
const CLIENT_ID = process.env.LINK_INVENTORY_CLIENT_ID ?? "";
const CLIENT_CERT_B64 = process.env.LINK_INVENTORY_CLIENT_CERT_PEM_BASE64 ?? "";
const SPO_RESOURCE = getTenantOrigin();

let cachedCredential: TokenCredential | undefined;
let cachedToken: { token: string; expiresOnTimestamp: number } | undefined;

function getCredential(): TokenCredential {
  if (cachedCredential) return cachedCredential;
  if (!TENANT_ID || !CLIENT_ID) {
    throw new Error(
      "LINK_INVENTORY_TENANT_ID and LINK_INVENTORY_CLIENT_ID must be set in app settings",
    );
  }

  if (AUTH_MODE === "federation") {
    // MI federates into the Entra app. The MI mints a token for the
    // `api://AzureADTokenExchange` audience; we hand that to Entra as
    // the `client_assertion`. Same-tenant only — see file header.
    const miCredential = new ManagedIdentityCredential();
    const getAssertion = async (): Promise<string> => {
      const tok = await miCredential.getToken("api://AzureADTokenExchange/.default");
      if (!tok) throw new Error("Managed identity returned no token for AzureADTokenExchange audience");
      return tok.token;
    };
    cachedCredential = new ClientAssertionCredential(TENANT_ID, CLIENT_ID, getAssertion);
    return cachedCredential;
  }

  if (AUTH_MODE !== "cert") {
    throw new Error(
      `Unsupported LINK_INVENTORY_AUTH_MODE='${AUTH_MODE}' (expected 'cert' or 'federation')`,
    );
  }

  if (!CLIENT_CERT_B64) {
    throw new Error("LINK_INVENTORY_CLIENT_CERT_PEM_BASE64 must be set when auth mode is 'cert'");
  }
  const pem = Buffer.from(CLIENT_CERT_B64, "base64").toString("utf8");
  cachedCredential = new ClientCertificateCredential(TENANT_ID, CLIENT_ID, {
    certificate: pem,
  });
  return cachedCredential;
}

/**
 * Returns a SharePoint-audience bearer token suitable for `Authorization:
 * Bearer <token>` against `https://contoso.sharepoint.com/_api/...`.
 *
 * Tokens are cached in-memory until 60s before expiry. Per Function-instance.
 */
export async function getSpoToken(): Promise<string> {
  const now = Date.now();
  if (cachedToken && cachedToken.expiresOnTimestamp - 60_000 > now) {
    return cachedToken.token;
  }

  const cred = getCredential();
  const tok = await cred.getToken(`${SPO_RESOURCE}/.default`);
  if (!tok) throw new Error("Token provider returned no SPO token");
  cachedToken = { token: tok.token, expiresOnTimestamp: tok.expiresOnTimestamp };
  return tok.token;
}

export const SPO_ORIGIN = SPO_RESOURCE;
