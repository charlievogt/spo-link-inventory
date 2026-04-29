import { app, type HttpResponseInit, type HttpRequest, type InvocationContext } from "@azure/functions";
import { AuthError, getSpoUserWriteToken, parseUserPrincipal, requireAdmin } from "../services/linkInventoryAuth.js";
import { SPO_ORIGIN } from "../services/spoTokenProvider.js";

/**
 * Admin-only: enable the Backlinks column on a target library.
 *
 * POST /api/link-inventory/backlinks-column/enable
 * Body: { siteUrl, libraryTitle }
 *
 * Creates the stub `RmgrBacklinks` multi-line text column, registers
 * the field customizer against it, and adds it to the library's
 * default view. The column holds no meaningful data — the field
 * customizer fetches the real backlink list from /backlinks at render
 * time with per-user ACL enforcement.
 *
 * Column naming:
 *   - Internal name: `RmgrBacklinks` (prefixed so it won't collide with
 *     a user-created "Backlinks" column)
 *   - Display name: `Backlinks`
 *
 * Pre-check: if either the internal name OR the display name is
 * already in use, abort with 409 so we don't overwrite someone else's
 * column or force-rename an unrelated field.
 *
 * Writes execute under the calling admin's OBO token (audit trail
 * pins the schema change to them, not to the Function MI).
 */

const COLUMN_INTERNAL_NAME = "RmgrBacklinks";
const COLUMN_DISPLAY_NAME = "Backlinks";

/**
 * Field customizer component id. Matches the manifest id emitted by
 * spfx/backlink-column-customizer when we build that package. Stored
 * as an env var so the function app can be updated without a redeploy
 * of func/ when the customizer guid changes (e.g., forced rotation
 * to bust the SPFx tenant manifest cache).
 */
const DEFAULT_CUSTOMIZER_COMPONENT_ID = "d4e71d3f-9a2c-4bf7-8f05-2e3d68b14c7a";
const CUSTOMIZER_COMPONENT_ID =
  process.env.BACKLINK_CUSTOMIZER_COMPONENT_ID || DEFAULT_CUSTOMIZER_COMPONENT_ID;

interface EnableBody {
  siteUrl?: string;
  libraryTitle?: string;
  /**
   * Values to serialize into the field's ClientSideComponentProperties
   * so the customizer knows how to reach the function. Admin web part
   * passes the same functionUrl + functionKey it uses itself.
   */
  customizerFunctionUrl?: string;
  customizerFunctionKey?: string;
}

interface EnableResponse {
  ok: boolean;
  siteUrl: string;
  libraryTitle: string;
  columnInternalName: string;
  columnDisplayName: string;
  addedToDefaultView: boolean;
  customizerRegistered: boolean;
}

function sitePathFromUrl(urlStr: string): string {
  const p = new URL(urlStr).pathname.replace(/\/$/, "");
  return p || "/";
}

function escapeODataSingleQuote(s: string): string {
  return s.replace(/'/g, "''");
}

async function checkFieldExists(sitePath: string, listTitle: string, nameOrTitle: string, token: string): Promise<boolean> {
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('${escapeODataSingleQuote(listTitle)}')` +
    `/fields/getbyinternalnameortitle('${escapeODataSingleQuote(nameOrTitle)}')?$select=Id`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}`, Accept: "application/json;odata=nometadata" },
  });
  if (res.ok) return true;
  if (res.status === 404) return false;
  // SP often returns 400 with a "does not exist" body when the named
  // field isn't on the list. Treat that as not-found rather than as
  // an error to bubble up.
  const body = await res.text();
  if (res.status === 400 && /does not exist/i.test(body)) return false;
  throw new Error(`Field existence check failed (${res.status}): ${body.slice(0, 300)}`);
}

async function createNoteField(sitePath: string, listTitle: string, token: string): Promise<void> {
  // Use CreateFieldAsXml for precise control over internal name.
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('${escapeODataSingleQuote(listTitle)}')` +
    `/fields/createfieldasxml`;
  const schemaXml =
    `<Field Type="Note" Name="${COLUMN_INTERNAL_NAME}" StaticName="${COLUMN_INTERNAL_NAME}"` +
    ` DisplayName="${COLUMN_DISPLAY_NAME}" RichText="FALSE" NumLines="1"` +
    ` Hidden="FALSE" Required="FALSE" />`;
  const body = {
    parameters: {
      SchemaXml: schemaXml,
      // AddToAllContentTypes = 4, AddFieldInternalNameHint = 8, combined = 12
      Options: 12,
    },
  };
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`CreateFieldAsXml failed (${res.status}): ${text.slice(0, 400)}`);
  }
}

async function registerCustomizer(
  sitePath: string,
  listTitle: string,
  token: string,
  properties: { functionUrl?: string; functionKey?: string },
): Promise<boolean> {
  if (!CUSTOMIZER_COMPONENT_ID) {
    // Component id not configured — leave the column in place but let
    // the caller know the renderer isn't wired yet.
    return false;
  }
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('${escapeODataSingleQuote(listTitle)}')` +
    `/fields/getbyinternalnameortitle('${COLUMN_INTERNAL_NAME}')`;
  const componentProperties = JSON.stringify({
    functionUrl: properties.functionUrl ?? "",
    functionKey: properties.functionKey ?? "",
  });
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "X-HTTP-Method": "MERGE",
      "If-Match": "*",
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
    },
    body: JSON.stringify({
      ClientSideComponentId: CUSTOMIZER_COMPONENT_ID,
      ClientSideComponentProperties: componentProperties,
    }),
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Customizer registration failed (${res.status}): ${text.slice(0, 400)}`);
  }
  return true;
}

async function addToDefaultView(sitePath: string, listTitle: string, token: string): Promise<boolean> {
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/lists/getbytitle('${escapeODataSingleQuote(listTitle)}')` +
    `/defaultview/viewfields/addviewfield('${COLUMN_INTERNAL_NAME}')`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
    },
  });
  if (!res.ok) {
    // Not fatal — the column was created; admin can add it manually.
    return false;
  }
  return true;
}

async function enableHandler(
  request: HttpRequest,
  context: InvocationContext,
): Promise<HttpResponseInit> {
  let user;
  try {
    user = parseUserPrincipal(request.headers.get("authorization"));
    await requireAdmin(user);
  } catch (err) {
    if (err instanceof AuthError) return { status: err.status, jsonBody: { ok: false, error: err.message } };
    throw err;
  }

  let body: EnableBody;
  try {
    body = (await request.json()) as EnableBody;
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid JSON body" } };
  }

  const { siteUrl, libraryTitle } = body;
  if (!siteUrl || !libraryTitle) {
    return { status: 400, jsonBody: { ok: false, error: "siteUrl and libraryTitle are required" } };
  }

  let sitePath: string;
  try {
    sitePath = sitePathFromUrl(siteUrl);
  } catch {
    return { status: 400, jsonBody: { ok: false, error: "Invalid siteUrl" } };
  }

  let token: string;
  try {
    token = await getSpoUserWriteToken(user);
  } catch (err) {
    return { status: 401, jsonBody: { ok: false, error: `Write token acquisition failed: ${(err as Error).message}` } };
  }

  // Pre-check for conflicts on both internal name and display name.
  try {
    if (await checkFieldExists(sitePath, libraryTitle, COLUMN_INTERNAL_NAME, token)) {
      return {
        status: 409,
        jsonBody: {
          ok: false,
          error: `A column with internal name "${COLUMN_INTERNAL_NAME}" already exists on this library`,
        },
      };
    }
    if (await checkFieldExists(sitePath, libraryTitle, COLUMN_DISPLAY_NAME, token)) {
      return {
        status: 409,
        jsonBody: {
          ok: false,
          error: `A column named "${COLUMN_DISPLAY_NAME}" already exists on this library — rename or delete it, then try again`,
        },
      };
    }
  } catch (err) {
    context.error(`enable-column pre-check failed: ${(err as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: (err as Error).message } };
  }

  try {
    await createNoteField(sitePath, libraryTitle, token);
  } catch (err) {
    context.error(`create field failed: ${(err as Error).message}`);
    return { status: 500, jsonBody: { ok: false, error: (err as Error).message } };
  }

  let customizerRegistered = false;
  try {
    customizerRegistered = await registerCustomizer(sitePath, libraryTitle, token, {
      functionUrl: body.customizerFunctionUrl,
      functionKey: body.customizerFunctionKey,
    });
  } catch (err) {
    // Field created but customizer not registered — partial success.
    // The admin can re-run once the customizer is deployed.
    context.warn(`customizer registration failed (non-fatal): ${(err as Error).message}`);
  }

  const addedToDefaultView = await addToDefaultView(sitePath, libraryTitle, token).catch(() => false);

  const payload: EnableResponse = {
    ok: true,
    siteUrl,
    libraryTitle,
    columnInternalName: COLUMN_INTERNAL_NAME,
    columnDisplayName: COLUMN_DISPLAY_NAME,
    customizerRegistered,
    addedToDefaultView,
  };
  return { status: 200, jsonBody: payload };
}

app.http("linkInventoryBacklinksColumnEnable", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "link-inventory/backlinks-column/enable",
  handler: enableHandler,
});
