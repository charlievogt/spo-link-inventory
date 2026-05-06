#!/usr/bin/env bash
# Set up the Entra app, scope, API permissions, admin consent, and admin
# group in the *SharePoint tenant*.
#
# Run this AFTER you've signed in to your SP tenant:
#   az login --tenant <sp-tenant>.onmicrosoft.com --allow-no-subscriptions
#
# Usage:
#   ./scripts/setup-entra.sh \
#       [--app-name NAME] \
#       [--group-name NAME] \
#       [--add-user UPN] \
#       [--auth cert|federation] \
#       [--cert-path PATH_TO_PUBLIC_PEM] \
#       [--mi-object-id OID] \
#       [--app-perm broad|selected]
#
# Flag details:
#   --auth          (default: cert)
#                   - cert:       client-credentials with a certificate.
#                                 Works any-tenant. If --cert-path is not
#                                 given, a self-signed cert is generated
#                                 and the combined PEM is written to
#                                 .entra-output.env (base64).
#                   - federation: Function App MI federates into the
#                                 Entra app. Same-tenant only (Microsoft
#                                 blocks Entra-to-Entra federation
#                                 cross-tenant via AADSTS700236).
#                                 Requires --mi-object-id.
#
#   --cert-path     (cert mode only) Path to a PEM file containing the
#                   public certificate. The matching private key must be
#                   in the same file OR alongside it as <basename>-key.pem.
#                   Use this when the cert is sourced from Key Vault or
#                   another managed CA. Without this flag, a fresh
#                   self-signed cert is created.
#
#   --mi-object-id  (federation mode only) The Function App's
#                   system-assigned managed identity object id. Find via:
#                     az functionapp identity show -n <fn> -g <rg> \
#                         --query principalId -o tsv
#
#   --app-perm      (default: broad)
#                   - broad:    grants Sites.Read.All + Sites.FullControl.All
#                               (Application) on SharePoint. Tenant-wide
#                               read enables `enumerateSites` to walk SP
#                               search across the whole tenant. Find/replace
#                               writes still go through user OBO so the
#                               calling user must have edit rights.
#                   - selected: grants Sites.Selected (Application). The app
#                               sees only sites you explicitly grant via
#                               Grant-PnPAzureADAppSitePermission. Tenant
#                               enumeration won't return anything until
#                               you grant per-site, so configure
#                               LINK_INVENTORY_SCAN_SITES on the Function
#                               App with an explicit list.
#
# Outputs to scripts/.entra-output.env (sourced by the bicep step):
#   ENTRA_TENANT_ID
#   ENTRA_CLIENT_ID
#   ENTRA_APP_OBJECT_ID
#   ADMIN_GROUP_ID
#   AUTH_MODE
#   APP_PERM
#   ENTRA_CLIENT_CERT_PEM_BASE64   (cert mode only)
#   ENTRA_CLIENT_CERT_THUMBPRINT   (cert mode only)
#   MI_OBJECT_ID                   (federation mode only)

set -euo pipefail

APP_NAME="link-inventory-func"
GROUP_NAME="Link Inventory Admins"
ADD_USER=""
AUTH_MODE="cert"
CERT_PATH=""
MI_OBJECT_ID=""
APP_PERM="broad"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --app-name) APP_NAME="$2"; shift 2 ;;
    --group-name) GROUP_NAME="$2"; shift 2 ;;
    --add-user) ADD_USER="$2"; shift 2 ;;
    --auth) AUTH_MODE="$2"; shift 2 ;;
    --cert-path) CERT_PATH="$2"; shift 2 ;;
    --mi-object-id) MI_OBJECT_ID="$2"; shift 2 ;;
    --app-perm) APP_PERM="$2"; shift 2 ;;
    *) echo "Unknown option: $1" >&2; exit 1 ;;
  esac
done

case "$AUTH_MODE" in
  cert|federation) ;;
  *) echo "Error: --auth must be 'cert' or 'federation' (got '$AUTH_MODE')" >&2; exit 1 ;;
esac
case "$APP_PERM" in
  broad|selected) ;;
  *) echo "Error: --app-perm must be 'broad' or 'selected' (got '$APP_PERM')" >&2; exit 1 ;;
esac
# --mi-object-id is optional for federation mode. If absent, the
# federated credential must be added later via add-federated-credential.sh
# (after Bicep emits the MI principal id).

# Verify we're signed in to a tenant context
TENANT_ID=$(az account show --query tenantId -o tsv 2>/dev/null || true)
SIGNED_IN_USER=$(az ad signed-in-user show --query userPrincipalName -o tsv 2>/dev/null || true)
if [[ -z "$TENANT_ID" || -z "$SIGNED_IN_USER" ]]; then
  echo "Error: not signed in. Run:" >&2
  echo "  az login --tenant <your-sp-tenant>.onmicrosoft.com --allow-no-subscriptions" >&2
  exit 1
fi

if [[ -z "$ADD_USER" ]]; then
  ADD_USER="$SIGNED_IN_USER"
fi

echo "Tenant:           $TENANT_ID"
echo "Signed in as:     $SIGNED_IN_USER"
echo "App name:         $APP_NAME"
echo "Admin group:      $GROUP_NAME"
echo "Adding user:      $ADD_USER"
echo "Auth mode:        $AUTH_MODE"
echo "App permission:   $APP_PERM"
[[ "$AUTH_MODE" == "federation" ]] && echo "MI object id:     $MI_OBJECT_ID"
echo

# Resource + scope GUIDs are global / well-known
SHAREPOINT_RESOURCE_ID="00000003-0000-0ff1-ce00-000000000000"
SHAREPOINT_ALLSITES_READ="4e0d77b0-96ba-4398-af14-3baa780278f4"
SHAREPOINT_ALLSITES_WRITE="640ddd16-e5b7-4d71-9690-3f4022699ee7"
# Application permissions on SharePoint
SHAREPOINT_SITES_READ_ALL_APP="d13f72ca-a275-4b96-b789-48ebcc4da984"
SHAREPOINT_SITES_FULLCONTROL_ALL_APP="678536fe-1083-478a-9c59-b99265e6b0d3"
SHAREPOINT_SITES_SELECTED_APP="20d37865-089c-4dee-8c41-6967602d4ac8"
GRAPH_RESOURCE_ID="00000003-0000-0000-c000-000000000000"
GRAPH_USER_READ="e1fe6dd8-ba31-4d61-89e7-88639da4683d"
GRAPH_GROUPMEMBER_READ_ALL="bc024368-1153-4739-b217-4326f2e966d0"

# 1. Create app
echo "==> Creating Entra app $APP_NAME ..."
APP_JSON=$(az ad app create --display-name "$APP_NAME" --sign-in-audience AzureADMyOrg -o json)
APP_ID=$(echo "$APP_JSON" | grep -o '"appId": *"[^"]*"' | head -1 | cut -d'"' -f4)
APP_OBJECT_ID=$(echo "$APP_JSON" | grep -o '"id": *"[^"]*"' | head -1 | cut -d'"' -f4)
echo "    App ID:        $APP_ID"
echo "    Object ID:     $APP_OBJECT_ID"

# 2. Set Application ID URI
echo "==> Setting Application ID URI to api://$APP_ID ..."
az ad app update --id "$APP_ID" --identifier-uris "api://$APP_ID" >/dev/null

# 3. Create service principal
echo "==> Creating service principal ..."
APP_SP_ID=$(az ad sp create --id "$APP_ID" --query id -o tsv)
echo "    Service Principal ID: $APP_SP_ID"

# 4. Add user_impersonation scope via Graph PATCH
echo "==> Adding user_impersonation scope ..."
SCOPE_ID=$(node -e "process.stdout.write(require('crypto').randomUUID())")
SCOPE_BODY=$(cat <<EOF
{
  "api": {
    "oauth2PermissionScopes": [
      {
        "id": "$SCOPE_ID",
        "adminConsentDescription": "Allow the function to act on behalf of the signed-in user when calling SharePoint",
        "adminConsentDisplayName": "Access link inventory function as user",
        "userConsentDescription": "Allow the function to act on your behalf when calling SharePoint",
        "userConsentDisplayName": "Access link inventory function as you",
        "value": "user_impersonation",
        "type": "User",
        "isEnabled": true
      }
    ]
  }
}
EOF
)
echo "$SCOPE_BODY" > /tmp/spli-scope.json
az rest --method PATCH \
  --url "https://graph.microsoft.com/v1.0/applications/$APP_OBJECT_ID" \
  --headers "Content-Type=application/json" \
  --body "@/tmp/spli-scope.json" >/dev/null
rm -f /tmp/spli-scope.json

# 5. Delegated permissions (always)
echo "==> Adding SharePoint AllSites.Read + AllSites.Write (delegated) ..."
az ad app permission add --id "$APP_ID" \
  --api "$SHAREPOINT_RESOURCE_ID" \
  --api-permissions "${SHAREPOINT_ALLSITES_READ}=Scope" "${SHAREPOINT_ALLSITES_WRITE}=Scope" \
  >/dev/null 2>&1 || true

echo "==> Adding Graph User.Read + GroupMember.Read.All (delegated) ..."
# User.Read is required for /me/* endpoints (the function's admin-group
# fallback calls /me/transitiveMemberOf when the bearer token doesn't
# carry a groups claim). Without it Graph returns 403 and isAdmin
# silently defaults to false — admin features then look broken even
# though the user IS in the configured admin group.
az ad app permission add --id "$APP_ID" \
  --api "$GRAPH_RESOURCE_ID" \
  --api-permissions "${GRAPH_USER_READ}=Scope" "${GRAPH_GROUPMEMBER_READ_ALL}=Scope" \
  >/dev/null 2>&1 || true

# 5b. Application permissions on SharePoint (depends on --app-perm)
APP_PERM_ROLE_IDS=()
if [[ "$APP_PERM" == "broad" ]]; then
  echo "==> Adding SharePoint Sites.Read.All + Sites.FullControl.All (Application) ..."
  APP_PERM_ROLE_IDS=("$SHAREPOINT_SITES_READ_ALL_APP" "$SHAREPOINT_SITES_FULLCONTROL_ALL_APP")
else
  echo "==> Adding SharePoint Sites.Selected (Application) ..."
  APP_PERM_ROLE_IDS=("$SHAREPOINT_SITES_SELECTED_APP")
fi
PERM_ARGS=()
for id in "${APP_PERM_ROLE_IDS[@]}"; do
  PERM_ARGS+=("${id}=Role")
done
az ad app permission add --id "$APP_ID" \
  --api "$SHAREPOINT_RESOURCE_ID" \
  --api-permissions "${PERM_ARGS[@]}" \
  >/dev/null 2>&1 || true

# 6. Admin consent (delegated). Application permissions sometimes don't
# pick up via the CLI's admin-consent path, so we belt-and-suspenders
# them with explicit appRoleAssignments on the SP.
echo "==> Granting admin consent (delegated scopes) ..."
sleep 5
az ad app permission admin-consent --id "$APP_ID" >/dev/null 2>&1 || {
  echo "    Note: admin-consent CLI returned non-zero. Continuing — appRoleAssignments will fix the app permissions."
}

echo "==> Resolving SharePoint service principal id ..."
SPO_SP_ID=$(az ad sp show --id "$SHAREPOINT_RESOURCE_ID" --query id -o tsv)
echo "    SPO SP ID: $SPO_SP_ID"

for role_id in "${APP_PERM_ROLE_IDS[@]}"; do
  # Skip if admin-consent already created this assignment (the
  # `az ad app permission admin-consent` step above sometimes does it
  # for application permissions too — depends on tenant setup).
  EXISTING=$(az rest --method GET \
    --url "https://graph.microsoft.com/v1.0/servicePrincipals/$APP_SP_ID/appRoleAssignments" \
    --query "value[?appRoleId=='$role_id' && resourceId=='$SPO_SP_ID'] | [0].id" -o tsv 2>/dev/null || true)
  if [[ -n "$EXISTING" ]]; then
    echo "==> Application role $role_id already granted (id=$EXISTING)"
    continue
  fi
  echo "==> Granting application role $role_id ..."
  # Tolerate "already exists" — the prior admin-consent step may have
  # created the assignment between our GET-check and this POST (Graph
  # propagation race). Treating it as success keeps the script idempotent.
  GRANT_OUT=$(az rest --method POST \
    --url "https://graph.microsoft.com/v1.0/servicePrincipals/$APP_SP_ID/appRoleAssignedTo" \
    --headers "Content-Type=application/json" \
    --body "{\"principalId\": \"$APP_SP_ID\", \"resourceId\": \"$SPO_SP_ID\", \"appRoleId\": \"$role_id\"}" \
    2>&1) || {
    if echo "$GRANT_OUT" | grep -q "already exists"; then
      echo "    (already granted — skipping)"
    else
      echo "    Grant POST failed: $GRANT_OUT" >&2
      exit 1
    fi
  }
done

# 7. Auth credential — cert OR federated credential
CERT_PEM_BASE64=""
CERT_THUMBPRINT=""
if [[ "$AUTH_MODE" == "cert" ]]; then
  if [[ -n "$CERT_PATH" ]]; then
    echo "==> Uploading existing public cert from $CERT_PATH ..."
    if [[ ! -f "$CERT_PATH" ]]; then
      echo "Error: --cert-path '$CERT_PATH' not found" >&2; exit 1
    fi
    az ad app credential reset --id "$APP_ID" --cert "@$CERT_PATH" --append --display-name "func-cert" --years 3 >/dev/null
    # Caller is responsible for delivering the matching private key to
    # the Function App. We don't write a base64 PEM in this case.
    CERT_THUMBPRINT=$(openssl x509 -in "$CERT_PATH" -noout -fingerprint -sha1 | sed 's/^.*=//' | tr -d ':')
    echo "    Thumbprint: $CERT_THUMBPRINT"
    echo "    Note: --cert-path mode does NOT bundle the private key. Provide it"
    echo "          to the Function App separately as LINK_INVENTORY_CLIENT_CERT_PEM_BASE64,"
    echo "          ideally via an Azure Key Vault reference."
  else
    echo "==> Generating self-signed cert (3-year validity) ..."
    CERT_DIR="$(dirname "$0")/.cert"
    mkdir -p "$CERT_DIR"
    PEM_PUB="$CERT_DIR/li-cert.pem"
    PEM_KEY="$CERT_DIR/li-cert-key.pem"
    PEM_COMBINED="$CERT_DIR/li-cert-combined.pem"
    # MSYS_NO_PATHCONV avoids Git Bash mangling the openssl -subj argument
    # on Windows (it treats /CN=... as a path).
    MSYS_NO_PATHCONV=1 openssl req -x509 -newkey rsa:2048 -nodes -sha256 \
      -days 1095 \
      -keyout "$PEM_KEY" -out "$PEM_PUB" \
      -subj "/CN=link-inventory-func-cert" 2>/dev/null
    cat "$PEM_PUB" "$PEM_KEY" > "$PEM_COMBINED"
    chmod 600 "$PEM_KEY" "$PEM_COMBINED"

    az ad app credential reset --id "$APP_ID" --cert "@$PEM_PUB" --append --display-name "func-cert" --years 3 >/dev/null
    CERT_THUMBPRINT=$(openssl x509 -in "$PEM_PUB" -noout -fingerprint -sha1 | sed 's/^.*=//' | tr -d ':')
    CERT_PEM_BASE64=$(base64 -w 0 "$PEM_COMBINED")
    echo "    Thumbprint: $CERT_THUMBPRINT"
    echo "    Combined PEM: $PEM_COMBINED (chmod 600)"
    echo "    Base64 length: ${#CERT_PEM_BASE64}"
    echo
    echo "    IMPORTANT: $PEM_COMBINED contains the private key."
    echo "               Treat as a secret. Don't commit it. The script's .gitignore"
    echo "               excludes scripts/.cert/ — verify before pushing."
  fi
else
  if [[ -n "$MI_OBJECT_ID" ]]; then
    echo "==> Adding federated credential (issuer=this tenant, subject=MI object id) ..."
    FEDCRED_BODY=$(cat <<EOF
{
  "name": "func-mi",
  "issuer": "https://login.microsoftonline.com/$TENANT_ID/v2.0",
  "subject": "$MI_OBJECT_ID",
  "audiences": ["api://AzureADTokenExchange"]
}
EOF
)
    echo "$FEDCRED_BODY" > /tmp/spli-fedcred.json
    az ad app federated-credential create --id "$APP_ID" --parameters "@/tmp/spli-fedcred.json" >/dev/null
    rm -f /tmp/spli-fedcred.json
    echo "    Federated credential created."
  else
    echo "==> Federation mode: deferring federated credential creation."
    echo "    After Bicep emits the Function App's MI principal id, run:"
    echo "      ./scripts/add-federated-credential.sh \\"
    echo "          --client-id $APP_ID \\"
    echo "          --mi-principal-id <miPrincipalId> \\"
    echo "          --azure-tenant-id <azureTenantId>"
  fi
fi

# 8. Admin group
echo "==> Creating Entra group: $GROUP_NAME ..."
GROUP_NICKNAME=$(echo "$GROUP_NAME" | tr 'A-Z ' 'a-z' | tr -cd 'a-z0-9')
GROUP_ID=$(az ad group create \
  --display-name "$GROUP_NAME" \
  --mail-nickname "$GROUP_NICKNAME" \
  --query id -o tsv)
echo "    Group ID: $GROUP_ID"

# 9. Add user to group
echo "==> Adding $ADD_USER to $GROUP_NAME ..."
USER_ID=$(az ad user show --id "$ADD_USER" --query id -o tsv)
az ad group member add --group "$GROUP_ID" --member-id "$USER_ID" >/dev/null

# Output
OUTPUT_FILE="$(dirname "$0")/.entra-output.env"
{
  echo "# Generated by setup-entra.sh on $(date -u +"%Y-%m-%dT%H:%M:%SZ")"
  echo "ENTRA_TENANT_ID=$TENANT_ID"
  echo "ENTRA_CLIENT_ID=$APP_ID"
  echo "ENTRA_APP_OBJECT_ID=$APP_OBJECT_ID"
  echo "ADMIN_GROUP_ID=$GROUP_ID"
  echo "AUTH_MODE=$AUTH_MODE"
  echo "APP_PERM=$APP_PERM"
  if [[ -n "$CERT_PEM_BASE64" ]]; then
    echo "ENTRA_CLIENT_CERT_PEM_BASE64=$CERT_PEM_BASE64"
  fi
  if [[ -n "$CERT_THUMBPRINT" ]]; then
    echo "ENTRA_CLIENT_CERT_THUMBPRINT=$CERT_THUMBPRINT"
  fi
  if [[ -n "$MI_OBJECT_ID" ]]; then
    echo "MI_OBJECT_ID=$MI_OBJECT_ID"
  fi
} > "$OUTPUT_FILE"
chmod 600 "$OUTPUT_FILE"

echo
echo "============================================================"
echo "Done. Results saved to: $OUTPUT_FILE (chmod 600)"
echo "  source $OUTPUT_FILE"
echo "============================================================"
if [[ "$APP_PERM" == "selected" ]]; then
  echo
  echo "Sites.Selected mode reminder:"
  echo "  Grant the app per-site read access via PnP.PowerShell:"
  echo "    Connect-PnPOnline -Url https://<tenant>.sharepoint.com/sites/<site> -Interactive"
  echo "    Grant-PnPAzureADAppSitePermission -AppId $APP_ID -DisplayName '$APP_NAME' \\"
  echo "        -Site https://<tenant>.sharepoint.com/sites/<site> -Permissions Read"
  echo "  Set LINK_INVENTORY_SCAN_SITES on the Function App to the explicit comma-separated list."
fi
