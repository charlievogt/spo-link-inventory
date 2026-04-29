# Deploy guide

End-to-end deployment of the function backend. Estimated time: 15–25 minutes for a fresh tenant + Azure subscription.

## Prerequisites

- **Azure subscription** in any tenant. Free-tier always-free quota is plenty for a smoke test ($0/month expected).
- **SharePoint Online tenant** where you want the function to scan. Often a different tenant from the Azure subscription — both are supported.
- Tools installed locally:
  - [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli) ≥ 2.50
  - [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) v4
  - Node.js ≥ 18
  - Bash (Git for Windows works on Windows)
- **Permissions** in the SharePoint tenant: ability to create Entra app registrations and groups, and grant admin consent. Global Administrator or Cloud Application Administrator + Privileged Role Administrator works.

## Architecture recap

```
┌──────────────────────────────┐                ┌──────────────────────────────┐
│ Azure subscription           │                │ SharePoint tenant            │
│  (any tenant)                │                │                              │
│                              │  cert          │ • Entra app                  │
│  • Function App + system MI  │ ◄────────────► │ • Admin Entra group          │
│  • Storage account           │  OR federation │ • SharePoint                 │
│  • App Insights              │                │                              │
└──────────────────────────────┘                └──────────────────────────────┘
```

Two supported auth modes between the Function App and the Entra app in the SharePoint tenant:

- **`cert` (default)** — the Function App holds a private key in app settings (or a Key Vault reference for production); the matching public cert is uploaded to the Entra app. Works any-tenant, including the common case of *Azure subscription in tenant A, SharePoint in tenant B*.
- **`federation`** — the Function App's system-assigned managed identity federates into the Entra app via a federated credential. No credential at rest. **Only works when the MI and the Entra app live in the same Entra tenant** — Microsoft blocks Entra-to-Entra federation cross-tenant under [AADSTS700236](https://learn.microsoft.com/entra/identity-platform/reference-error-codes).

Both modes produce app-only tokens with `appidacr=2`, which SharePoint REST (notably the search endpoint) requires. The deprecated `client_secret` path produces `appidacr=1` tokens that SP rejects with *"Unsupported app only token"* — that's why this repo doesn't offer a secret mode.

For app-level SP permissions on the Entra app, two options:

- **`broad` (default)** — `Sites.Read.All` + `Sites.FullControl.All` (Application). Tenant-wide read enables the scanner to walk SP search across all sites. Find/replace writes still go through user OBO, so the *calling user* must have edit rights on a target site for any mutation to succeed.
- **`selected`** — `Sites.Selected` (Application) + per-site grants via `Grant-PnPAzureADAppSitePermission`. The app sees only sites you explicitly grant. Set `LINK_INVENTORY_SCAN_SITES` on the Function App with the explicit comma-separated site list.

## Cost guard (recommended)

Before creating Azure resources, set a low monthly budget on your subscription so a misconfiguration doesn't burn money. Smoke-test spend is genuinely $0 under the always-free tier, but the budget is cheap insurance:

```bash
SUB=<your-subscription-id>
EMAIL=<your-email>

az account set --subscription "$SUB"

cat > /tmp/budget.json <<EOF
{
  "properties": {
    "category": "Cost",
    "amount": 5,
    "timeGrain": "Monthly",
    "timePeriod": {
      "startDate": "$(date -u +%Y-%m-01T00:00:00Z)",
      "endDate":   "2030-04-01T00:00:00Z"
    },
    "notifications": {
      "Actual_GreaterThan_50_Percent": {
        "enabled": true, "operator": "GreaterThan", "threshold": 50,
        "contactEmails": ["$EMAIL"], "thresholdType": "Actual"
      },
      "Actual_GreaterThan_80_Percent": {
        "enabled": true, "operator": "GreaterThan", "threshold": 80,
        "contactEmails": ["$EMAIL"], "thresholdType": "Actual"
      }
    }
  }
}
EOF

az rest --method PUT \
  --uri "https://management.azure.com/subscriptions/$SUB/providers/Microsoft.Consumption/budgets/link-inventory-cap?api-version=2024-08-01" \
  --body @/tmp/budget.json
```

Budgets *alert*, they don't *block*. The hard stop is "I see the alert email and delete the resource group." Plain pay-as-you-go subs do not have a spending limit; only Free Trial subs auto-suspend.

## Step 1 — Register Azure resource providers (one-time per subscription)

Fresh Azure subscriptions don't have any resource providers registered. The Bicep deployment fails with cryptic "SubscriptionNotFound" errors otherwise.

```bash
az login    # sign in as your Azure user
az account set --subscription <your-subscription-id>
./scripts/register-providers.sh
```

Takes 1–3 minutes. Idempotent — safe to re-run.

## Step 2 — Set up the Entra app in your SharePoint tenant

```bash
# Sign in to your SP tenant. Use --allow-no-subscriptions because most
# SP-only tenants don't have an Azure subscription attached.
az login --tenant <your-sp-tenant>.onmicrosoft.com --allow-no-subscriptions

# Default — cert auth, broad app permissions (most common path):
./scripts/setup-entra.sh

# Alternatives:
#   ./scripts/setup-entra.sh --auth federation              # same-tenant only
#   ./scripts/setup-entra.sh --app-perm selected            # locked-down: per-site grants required
#   ./scripts/setup-entra.sh --auth cert --cert-path ./mycert.pem    # bring your own cert
```

The script:

1. Creates the Entra app (`link-inventory-func` by default)
2. Sets the Application ID URI to `api://<app-client-id>`
3. Adds the `user_impersonation` OAuth2 scope
4. Adds delegated SharePoint permissions (`AllSites.Read` + `AllSites.Write`) and Graph (`User.Read` + `GroupMember.Read.All`)
5. Adds the application-level SharePoint permissions implied by `--app-perm` and grants admin consent
6. Sets up the auth credential:
   - `cert` mode — generates a self-signed cert (or uses `--cert-path`), uploads the public cert to the Entra app, writes the combined PEM as base64 to `scripts/.entra-output.env`
   - `federation` mode — adds the federated credential immediately if `--mi-object-id` is given, otherwise defers to Step 4
7. Creates the "Link Inventory Admins" Entra group and adds you to it
8. Writes `scripts/.entra-output.env` with the values you need next

**Output:** the file `scripts/.entra-output.env` (chmod 600). For cert mode it includes:

```
ENTRA_TENANT_ID=...
ENTRA_CLIENT_ID=...
ENTRA_APP_OBJECT_ID=...
ADMIN_GROUP_ID=...
AUTH_MODE=cert
APP_PERM=broad
ENTRA_CLIENT_CERT_PEM_BASE64=...        # ← the long one. Treat as a secret.
ENTRA_CLIENT_CERT_THUMBPRINT=...
```

For federation mode, the cert fields are absent and `MI_OBJECT_ID` is added if provided.

> **Treat `.entra-output.env` as a secret in cert mode.** It contains the private key (base64-PEM) needed to authenticate as the Entra app. The repo's `.gitignore` excludes it, but mind any backup tools / sync clients pointed at the directory. For production deployments, replace the file-based cert with an Azure Key Vault reference (see *Production cert delivery* below).

## Step 3 — Deploy Azure resources

Switch back to your Azure-subscription context and deploy:

```bash
# Sign back in to your Azure subscription
az login --tenant <your-azure-tenant>.onmicrosoft.com   # or just `az login`
az account set --subscription <your-subscription-id>

# Pull the values from the previous step
source scripts/.entra-output.env

# Deploy. Pick a unique projectName — used as the suffix in all resource names.
# In cert mode, the combined PEM is passed via clientCertPemBase64.
az deployment sub create \
  --location centralus \
  --template-file infra/main.bicep \
  --parameters \
      projectName=linkinv01 \
      spoTenantHost=<your-sp-tenant>.sharepoint.com \
      entraTenantId=$ENTRA_TENANT_ID \
      entraClientId=$ENTRA_CLIENT_ID \
      adminGroupId=$ADMIN_GROUP_ID \
      authMode=$AUTH_MODE \
      clientCertPemBase64=${ENTRA_CLIENT_CERT_PEM_BASE64:-""}
```

For federation mode, omit `clientCertPemBase64` (or leave it empty — the Bicep template gates the cert app setting on `authMode == 'cert'`).

> **Windows users — if you hit `[WinError 5] Access is denied`** when `az` tries to subprocess Bicep, run `./scripts/install-bicep.sh` once, then compile to ARM JSON manually and pass that instead:
>
> ```bash
> ./scripts/install-bicep.sh
> ~/.azure/bin/bicep build infra/main.bicep --outfile /tmp/main.arm.json
> az deployment sub create --location centralus --template-file /tmp/main.arm.json --parameters ...
> ```
>
> Root cause: az CLI's bundled Bicep installer is blocked by some Windows AV / permission setups. Downloading the standalone binary and compiling explicitly sidesteps it. macOS / Linux users won't hit this.

Takes 3–5 minutes. The deployment outputs four values:

- `functionAppName` — pass to `func azure functionapp publish`
- `functionAppHostname` — for the SPFx web part property
- `miPrincipalId` — feeds into the next step
- `azureTenantId` — feeds into the next step

You can capture them with `--query`:

```bash
DEPLOY_OUT=$(az deployment sub create \
  --location centralus \
  --template-file infra/main.bicep \
  --parameters projectName=linkinv01 spoTenantHost=<your-sp-tenant>.sharepoint.com \
               entraTenantId=$ENTRA_TENANT_ID entraClientId=$ENTRA_CLIENT_ID adminGroupId=$ADMIN_GROUP_ID \
  --query properties.outputs -o json)

export FUNCTION_APP_NAME=$(echo "$DEPLOY_OUT" | node -p "JSON.parse(require('fs').readFileSync(0,'utf8')).functionAppName.value")
export MI_PRINCIPAL_ID=$(echo  "$DEPLOY_OUT" | node -p "JSON.parse(require('fs').readFileSync(0,'utf8')).miPrincipalId.value")
export AZURE_TENANT_ID=$(echo  "$DEPLOY_OUT" | node -p "JSON.parse(require('fs').readFileSync(0,'utf8')).azureTenantId.value")
```

## Step 4 — Wire up the federated credential (federation mode only)

**Skip this step if `AUTH_MODE=cert`** — the cert was uploaded in Step 2 and the Function App got the matching private key from the Bicep parameter in Step 3. You're done with auth setup.

For federation mode, go back to the SP-tenant context to add the federated credential on the Entra app:

```bash
az login --tenant <your-sp-tenant>.onmicrosoft.com --allow-no-subscriptions

./scripts/add-federated-credential.sh \
  --client-id $ENTRA_CLIENT_ID \
  --mi-principal-id $MI_PRINCIPAL_ID \
  --azure-tenant-id $AZURE_TENANT_ID
```

This is the moment the Function App's MI is allowed to act as the Entra app for OBO. Without it, the function will get 401s from AAD when trying to exchange MI tokens.

**Reminder:** federation only works if the Function App's MI and the Entra app live in the same Entra tenant. If they don't (e.g. Azure subscription in tenant A, SharePoint in tenant B), the federated credential exchange returns `AADSTS700236` and you'll need to redeploy in cert mode.

## Step 5 — Build and publish the function code

```bash
cd func
npm install
npm run build
func azure functionapp publish $FUNCTION_APP_NAME --typescript --build remote
```

The `--build remote` flag tells Azure (Oryx) to run `npm install` server-side, which is what a Linux Consumption plan needs.

## Step 6 — Smoke test

```bash
# Function alive (auth-gated, so 401 = success)
curl -s -o /dev/null -w "%{http_code}\n" "https://${FUNCTION_APP_NAME}.azurewebsites.net/api/link-inventory/whoami"
# expect: 401

# Function count
az functionapp function list --name $FUNCTION_APP_NAME --resource-group rg-linkinv01 --query "length(@)"
# expect: 30
```

A 401 here is the **good** outcome — it means Node started, the runtime registered functions, and the auth middleware is rejecting the unauthenticated request. The actual end-to-end auth check happens once the SPFx web part is configured to call the function.

## Step 7 — SPFx side (web part deployment)

### 7a. Build the .sppkg

```bash
cd spfx/link-inventory-admin
npm install
npm run build
```

`npm run build` runs `heft build --production && heft package-solution --production`. The artifact lands at:

```
spfx/link-inventory-admin/sharepoint/solution/link-inventory-admin.sppkg
```

(That path is gitignored — the .sppkg gets regenerated on every build.)

### 7b. Upload to the tenant app catalog

In any browser, sign in as a SharePoint admin and open:

```
https://<your-sp-tenant>.sharepoint.com/sites/appcatalog/AppCatalog/Forms/AllItems.aspx
```

If your tenant doesn't have an app catalog yet, create one first via the SharePoint admin center → *More features* → *Apps* → *Open* → *App Catalog* → *OK*. Provisioning takes a minute or two.

Drag the `.sppkg` from your local file system into the library. SharePoint shows a trust dialog:

- **Title:** "Do you trust Link Inventory Admin?"
- **Checkbox:** "Make this solution available to all sites in the organization" — **check it**
- Click **Deploy**

The package shows up in the catalog with **Deployed = Yes**.

### 7b-bis. Approve the API permission request

Uploading the .sppkg also creates a tenant-level API permission request: SharePoint needs to grant its built-in token-broker app (*SharePoint Online Client Extensibility Web Application Principal*) the `user_impersonation` scope on the function's Entra app, so the broker can mint OBO tokens on behalf of users when they load the web part.

Approve it:

1. Open the SharePoint admin center: `https://<your-sp-tenant>-admin.sharepoint.com`
2. Left nav: **Advanced** → **API access**
3. You'll see a pending request: resource = **link-inventory-func**, scope = **user_impersonation**
4. Click the request → **Approve**

Without this approval, the web part loads but errors with `AADSTS65001: The user or administrator has not consented to use the application … 'SharePoint Online Web Client Extensibility'` when it tries to call the function. The app registration's `user_impersonation` scope is user-consentable, but the broker still needs a tenant-wide grant to relay tokens for it.

(The .sppkg's `package-solution.json` declares this request via `webApiPermissionRequests` so admin approval is one click. If you ever need to re-approve, the request shows up in the same SP admin center page.)

### 7c. Add the web part to a page

Pick or create a modern site page where the web part will live (the admin site you provisioned earlier is a fine home). Open the page, click **Edit**.

Click the **+** button anywhere on the canvas to add a web part. In the picker:

- Type **Link Inventory** in the search box
- The tile reads **Link Inventory Admin** (group: *Tools*)
- Click it to insert

### 7d. Configure the two properties

The property pane opens automatically when you add the web part. Set:

| Property | Value |
|---|---|
| **Azure Function URL** | `https://<function-app-name>.azurewebsites.net` (the value the Bicep deployment printed as `functionAppHostname`, with `https://` prefix) |
| **Entra app api:// URI** | `api://<entra-client-id>` (use the URI form — the friendly display name intermittently fails to resolve. The value is in `scripts/.entra-output.env` as `ENTRA_CLIENT_ID`.) |

Click **Apply**, then **Save and publish** at the top of the page.

### 7e. First-time auth consent

When the first user opens the published page, AAD prompts for consent to the `user_impersonation` scope on your Entra app:

> **Permissions requested**
> link-inventory-func — Access link inventory function as you

Click **Accept**. (If you're a global admin, you can also click **Consent on behalf of your organization** to skip this prompt for everyone else.)

After consent the web part loads — three tabs (Link Inventory, Duplicates, Help). The Link Inventory tab will show "no scans yet" until you trigger your first scan from the **Run scan** dropdown.

After your first unified scan completes, the Link Inventory tab looks like this — pages and document files merged into one inventory, with sharing wrappers, AllItems URLs, mailto/external/relative all classified into the same table:

![Link inventory after first unified scan](images/inventory-overview.png)

Click any row to open the link details panel. Page rows expose a **Fix this link** button that rewrites the canvas under your identity:

![Link details panel for a page-row sharing wrapper](images/detail-page-row.png)

Document rows show the same suggestion but no write actions (rewriting hyperlinks inside binary OOXML/PDF files is out of scope):

![Link details panel for a doc-row — Read-only](images/detail-doc-row.png)

The **Duplicates** tab uses SHA-256 hashing across the scanned libraries to surface exact duplicates, stale copies (via version-history walk), and same-name pairs:

![Duplicates tab — Exact / Stale / Same-name groups](images/duplicates-tab.png)

### 7f. Verify the auth flow worked

A quick sanity check from any logged-in user's browser console on the page:

```javascript
// Should return 200 with { ok: true, userId, upn, isAdmin }
fetch('https://<function-app-name>.azurewebsites.net/api/link-inventory/whoami', {
  headers: { Authorization: 'Bearer ' + (await window.spfxAad.getClient('api://<entra-client-id>')).get())  // pseudo-code; SPFx wraps this
})
```

In practice you don't need to write this — if the web part renders and the **Run scan** button is visible (admin-only), the auth handshake is working end-to-end.

## Production cert delivery (recommended for cert mode)

The default flow ships the combined PEM as a plain Function App setting, which is fine for evaluation and dev tenants but not ideal for production: the value is visible to anyone with `Reader` on the Function App, doesn't rotate automatically, and isn't audited.

For production, replace the literal value with an Azure Key Vault reference:

1. Create a Key Vault in the same subscription as the Function App.
2. Import the combined PEM as a Key Vault *secret* (not a *certificate* — the Function App needs the raw base64 string the code already expects, and Key Vault certificates expose the PFX form):
   ```bash
   az keyvault secret set --vault-name <vault> --name link-inventory-client-cert --value "$(cat scripts/.cert/li-cert-combined.pem | base64 -w 0)"
   ```
3. Grant the Function App's system-assigned MI `Key Vault Secrets User` on the vault.
4. Replace the app setting value with the Key Vault reference syntax:
   ```
   LINK_INVENTORY_CLIENT_CERT_PEM_BASE64=@Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/link-inventory-client-cert)
   ```
5. Rotate by updating the secret value in Key Vault — the Function App picks up the new version automatically (the reference doesn't pin a version unless you add `/<version>` to the URI).

Bicep doesn't enforce the Key Vault path because the OSS scripts default to file-based cert delivery for first-run simplicity. Once you've validated the deploy, switch to Key Vault and delete the local PEM files.

## Updating the deployed code

Whenever you change the func code:

```bash
cd func
npm run build
func azure functionapp publish $FUNCTION_APP_NAME --typescript --build remote
```

That's it — settings persist across deploys.

## Cleanup

To remove everything:

```bash
# Azure side
az group delete --name rg-<projectName> --yes --no-wait

# SP tenant side
az ad app delete --id <ENTRA_CLIENT_ID>
az ad group delete --group <ADMIN_GROUP_ID>
```

## Troubleshooting

**Windows: `[WinError 5] Access is denied`** when running `az deployment sub create` or `az bicep build` — `az` can't subprocess its bundled Bicep binary. Workaround: run `./scripts/install-bicep.sh`, then compile to ARM JSON manually (`~/.azure/bin/bicep build infra/main.bicep --outfile /tmp/main.arm.json`) and deploy the JSON file. See the note inline in Step 3 above.

**`SubscriptionNotFound` during Bicep deploy** — resource providers aren't registered. Run `./scripts/register-providers.sh` and retry.

**`func azure functionapp publish` succeeds but `/api/link-inventory/whoami` returns 404** — usually means `--build remote` was omitted and dependencies weren't installed server-side. Re-publish with `--build remote`.

**Function returns 500** — check Application Insights traces:
```bash
APP_ID=$(az monitor app-insights component show --app appi-<projectName> --resource-group rg-<projectName> --query appId -o tsv)
az monitor app-insights query --app $APP_ID --analytics-query "exceptions | top 10 by timestamp desc | project timestamp, type, outerMessage"
```

**`AADSTS500011` (resource not found) when SPFx calls the function** — `entraAppApiUri` web part property must be the URI form (`api://<guid>`), not the friendly display name. Friendly-name lookup intermittently fails in the SP approved-permissions list.

**`AADSTS700236` during federated credential exchange** — Microsoft blocks Entra-to-Entra federation cross-tenant. If your Function App's MI lives in a different Entra tenant from the Entra app (common when Azure subscription is in tenant A and SharePoint in tenant B), redeploy in `cert` mode: re-run `setup-entra.sh` (no flags = cert default) and re-deploy the Bicep with `authMode=cert clientCertPemBase64=$ENTRA_CLIENT_CERT_PEM_BASE64`.

**SP search returns `Unsupported app only token`** — SharePoint REST search requires `appidacr=2` tokens (cert or federated). If you somehow ended up with a `client_secret`-issued token, replace with a cert-issued credential. This repo doesn't offer a secret mode for that reason.

**Cert appears uploaded but token call fails with `invalid_client`** — usually a thumbprint mismatch between what the Entra app expects and what the JWT `x5t` header carries. Verify the public PEM uploaded to the app matches the private key in `LINK_INVENTORY_CLIENT_CERT_PEM_BASE64`:
```bash
# Public part the Entra app holds (compare customKeyIdentifier to your local thumbprint):
az ad app show --id $ENTRA_CLIENT_ID --query "keyCredentials[].customKeyIdentifier" -o tsv

# Local thumbprint:
openssl x509 -in scripts/.cert/li-cert.pem -noout -fingerprint -sha1 | sed 's/^.*=//' | tr -d ':'
```

## References

Canonical Microsoft Learn docs for the patterns this guide builds on. Read these if any step's *why* feels unclear.

**Federated identity / OBO** (Step 2, Step 4):
- [Workload identity federation](https://learn.microsoft.com/entra/workload-id/workload-identity-federation) — overview of the trust pattern between an external identity (the Function App's MI in your Azure subscription) and an Entra app (in your SP tenant).
- [Configure an application to trust a managed identity](https://learn.microsoft.com/entra/workload-id/workload-identity-federation-config-app-trust-managed-identity) — exact field semantics for the federated credential (`issuer`, `subject`, `audiences`). The `add-federated-credential.sh` script automates this.
- [Microsoft Graph permissions reference](https://learn.microsoft.com/graph/permissions-reference) — authoritative list of the Graph permission GUIDs `setup-entra.sh` requests (`User.Read`, `GroupMember.Read.All`).

**Bicep** (Step 3):
- [Subscription deployments with Bicep files](https://learn.microsoft.com/azure/azure-resource-manager/bicep/deploy-to-subscription) — `targetScope = 'subscription'` semantics. `infra/main.bicep` uses this scope so the Bicep can create the resource group itself.

**Azure Functions** (Step 5):
- [Azure Functions Node.js developer guide (v4 model)](https://learn.microsoft.com/azure/azure-functions/functions-reference-node) — explains the `app.http(...)` registration pattern and the `main` field in `package.json` that loads `dist/src/index.js`.
- [Deployment technologies in Azure Functions](https://learn.microsoft.com/azure/azure-functions/functions-deployment-technologies) — covers zip deploy, `WEBSITE_RUN_FROM_PACKAGE`, and the Linux Consumption "remote build" flow that `--build remote` triggers (Oryx running `npm install` on the server).

**SPFx** (Step 7):
- [Connect to Entra ID-secured APIs in SharePoint Framework solutions](https://learn.microsoft.com/sharepoint/dev/spfx/use-aadhttpclient) — explains how `AadHttpClient.getClient(api://<guid>)` works under the covers (the SharePoint Online Client Extensibility service principal, the implicit OAuth flow). Useful context for why we configure `entraAppApiUri` as a web part property.

**SharePoint REST** (used internally by `setup-entra.sh` and the validation flow):
- [Manage modern SharePoint sites using REST](https://learn.microsoft.com/sharepoint/dev/apis/site-creation-rest) — the `_api/SPSiteManager/create` endpoint used to provision the admin site.
