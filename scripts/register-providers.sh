#!/usr/bin/env bash
# Register the resource providers spo-link-inventory needs.
# One-time, per Azure subscription. Idempotent.
#
# Usage: ./scripts/register-providers.sh [<subscription-id>]
#
# If no subscription ID is given, uses the currently-active one.

set -euo pipefail

SUBSCRIPTION="${1:-}"
if [[ -n "$SUBSCRIPTION" ]]; then
  az account set --subscription "$SUBSCRIPTION"
fi

CURRENT_SUB=$(az account show --query id -o tsv)
echo "Registering providers on subscription: $CURRENT_SUB"
echo

PROVIDERS=(
  Microsoft.Storage
  Microsoft.Web
  Microsoft.Insights
  Microsoft.OperationalInsights
  Microsoft.Authorization
)

for ns in "${PROVIDERS[@]}"; do
  echo "  registering $ns..."
  az provider register --namespace "$ns" --wait >/dev/null 2>&1 &
done
wait

echo
echo "All providers registered:"
az provider list \
  --query "[?namespace=='Microsoft.Storage' || namespace=='Microsoft.Web' || namespace=='microsoft.insights' || namespace=='Microsoft.OperationalInsights' || namespace=='Microsoft.Authorization'].{ns:namespace, state:registrationState}" \
  -o table
