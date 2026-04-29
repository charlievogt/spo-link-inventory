/**
 * Environment-driven configuration for the function. All tenant-specific
 * values live here so the rest of the codebase stays generic.
 *
 * Loaded lazily on first read; throws with a clear message when a required
 * setting is missing rather than letting downstream calls fail with cryptic
 * errors deep inside SP REST or Graph.
 */

function required(name: string): string {
  const v = process.env[name];
  if (!v || !v.trim()) {
    throw new Error(
      `Missing required app setting: ${name}. ` +
        `Set it in local.settings.json (local) or as an Application Setting on the Function App (prod).`,
    );
  }
  return v.trim();
}

/** Tenant SharePoint hostname, e.g. "contoso.sharepoint.com" (no protocol, no trailing slash). */
export function getTenantHost(): string {
  return required("SPO_TENANT_HOST").toLowerCase().replace(/\/$/, "");
}

/** Absolute SPO origin, e.g. "https://contoso.sharepoint.com". */
export function getTenantOrigin(): string {
  return `https://${getTenantHost()}`;
}

/**
 * OBO target scope used by the writer service. Defaults to
 * `https://${SPO_TENANT_HOST}/.default` when unset, which is correct for
 * almost all deployments.
 */
export function getOboTargetScope(): string {
  const explicit = process.env.OBO_TARGET_SCOPE?.trim();
  if (explicit) return explicit;
  return `${getTenantOrigin()}/.default`;
}
