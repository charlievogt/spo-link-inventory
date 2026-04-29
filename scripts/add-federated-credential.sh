#!/usr/bin/env bash
# Add a federated credential to the Entra app, pointing at the Function App's
# managed identity. Run this AFTER the Bicep deployment, in the SP tenant
# context.
#
# Usage:
#   ./scripts/add-federated-credential.sh \
#     --client-id <ENTRA_CLIENT_ID> \
#     --mi-principal-id <MI_PRINCIPAL_ID> \
#     --azure-tenant-id <AZURE_TENANT_ID> \
#     [--name <credential-name>]
#
# All three required values are emitted by:
#   - setup-entra.sh                  → ENTRA_CLIENT_ID
#   - az deployment sub create        → MI_PRINCIPAL_ID, AZURE_TENANT_ID
#
# Sign in to your SP tenant first:
#   az login --tenant <sp-tenant>.onmicrosoft.com --allow-no-subscriptions

set -euo pipefail

NAME="function-mi"
CLIENT_ID=""
MI_PRINCIPAL_ID=""
AZURE_TENANT_ID=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --client-id) CLIENT_ID="$2"; shift 2 ;;
    --mi-principal-id) MI_PRINCIPAL_ID="$2"; shift 2 ;;
    --azure-tenant-id) AZURE_TENANT_ID="$2"; shift 2 ;;
    --name) NAME="$2"; shift 2 ;;
    *) echo "Unknown option: $1" >&2; exit 1 ;;
  esac
done

if [[ -z "$CLIENT_ID" || -z "$MI_PRINCIPAL_ID" || -z "$AZURE_TENANT_ID" ]]; then
  echo "Error: --client-id, --mi-principal-id, --azure-tenant-id are all required." >&2
  exit 1
fi

# Verify we're in the right tenant context (must be SP tenant, where the app lives)
SIGNED_IN_USER=$(az ad signed-in-user show --query userPrincipalName -o tsv 2>/dev/null || true)
if [[ -z "$SIGNED_IN_USER" ]]; then
  echo "Error: not signed in. Run:" >&2
  echo "  az login --tenant <your-sp-tenant>.onmicrosoft.com --allow-no-subscriptions" >&2
  exit 1
fi

ISSUER="https://login.microsoftonline.com/$AZURE_TENANT_ID/v2.0"

echo "Adding federated credential to app $CLIENT_ID:"
echo "  Name:     $NAME"
echo "  Issuer:   $ISSUER"
echo "  Subject:  $MI_PRINCIPAL_ID"
echo "  Audience: api://AzureADTokenExchange"
echo

cat > /tmp/spli-fic.json <<EOF
{
  "name": "$NAME",
  "description": "Function App managed identity acts as this app for OBO token exchange",
  "issuer": "$ISSUER",
  "subject": "$MI_PRINCIPAL_ID",
  "audiences": ["api://AzureADTokenExchange"]
}
EOF

az ad app federated-credential create \
  --id "$CLIENT_ID" \
  --parameters "@/tmp/spli-fic.json" \
  --query "{name:name, issuer:issuer, subject:subject}" -o json

rm -f /tmp/spli-fic.json
echo
echo "Done."
