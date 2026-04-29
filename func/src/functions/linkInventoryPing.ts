import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { getSpoToken, SPO_ORIGIN } from "../services/spoTokenProvider.js";

/**
 * Auth spike for the Link Inventory feature.
 *
 * Proves the chain MI → federated → Entra app (Sites.Selected) → SP REST
 * works end-to-end on a single test site. Read-only, hardcoded site.
 *
 * Expected response:
 *   { ok: true, site, pageCount, samplePageTitles, tokenSource }
 *
 * If you get 401/403, see PICKUP-link-inventory-auth.md troubleshooting.
 */

const TEST_SITE_PATH = "/sites/charlie-test-site";

interface SitePagesResponse {
  value: Array<{ Id: number; Title?: string; FileLeafRef?: string }>;
}

async function pingHandler(
  _request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  try {
    const token = await getSpoToken();

    const url = `${SPO_ORIGIN}${TEST_SITE_PATH}/_api/web/lists/getbytitle('Site Pages')/items?$select=Id,Title,FileLeafRef&$top=5`;
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
      },
    });

    if (!res.ok) {
      const body = await res.text();
      context.error(`SPO REST ${res.status}: ${body.slice(0, 500)}`);
      return {
        status: res.status,
        jsonBody: {
          ok: false,
          step: "spo-rest-call",
          status: res.status,
          body: body.slice(0, 500),
          hint:
            res.status === 401 || res.status === 403
              ? "Likely missing Grant-PnPAzureADAppSitePermission for this site, or admin consent not granted on the Entra app. See func/PICKUP-link-inventory-auth.md"
              : undefined,
        },
      };
    }

    const data = (await res.json()) as SitePagesResponse;
    return {
      status: 200,
      jsonBody: {
        ok: true,
        site: TEST_SITE_PATH,
        pageCount: data.value.length,
        samplePageTitles: data.value.map((p) => p.Title ?? p.FileLeafRef ?? `#${p.Id}`),
        tokenSource: "managed-identity-federated",
      },
    };
  } catch (err) {
    const e = err as Error;
    context.error(`linkInventoryPing failed: ${e.message}\n${e.stack ?? ""}`);
    return {
      status: 500,
      jsonBody: {
        ok: false,
        step: "token-acquisition",
        error: e.message,
        hint: "Check LINK_INVENTORY_CLIENT_ID / LINK_INVENTORY_TENANT_ID app settings and that the Function App has a system-assigned managed identity with a federated credential on the Entra app.",
      },
    };
  }
}

app.http("linkInventoryPing", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "link-inventory/ping",
  handler: pingHandler,
});
