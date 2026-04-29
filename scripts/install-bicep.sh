#!/usr/bin/env bash
# Install or refresh the Bicep CLI binary used by `az`.
#
# Why this exists: on some Windows setups (corporate AV, strict permissions)
# `az` fails to subprocess its bundled bicep.exe with "[WinError 5] Access
# is denied", which blocks `az deployment ... --template-file infra/main.bicep`.
# Downloading bicep.exe directly to ~/.azure/bin/ sidesteps the bundled
# installer and works reliably.
#
# Usage:
#   ./scripts/install-bicep.sh              # latest from GitHub releases
#
# After running, you may also need to compile bicep manually before
# deploying — see "Windows: Bicep subprocess errors" in docs/deploy.md.

set -euo pipefail

# Detect platform
case "$(uname -s)" in
  MINGW*|MSYS*|CYGWIN*) PLATFORM=win-x64 ; EXT=.exe ;;
  Linux*)               PLATFORM=linux-x64 ; EXT= ;;
  Darwin*)              PLATFORM=osx-x64 ; EXT= ;;
  *) echo "Unsupported platform: $(uname -s)" >&2; exit 1 ;;
esac

DEST_DIR="$HOME/.azure/bin"
DEST="$DEST_DIR/bicep$EXT"
URL="https://github.com/Azure/bicep/releases/latest/download/bicep-${PLATFORM}${EXT}"

mkdir -p "$DEST_DIR"
echo "Downloading: $URL"
curl -fsSL "$URL" -o "$DEST"
chmod +x "$DEST"

echo
echo "Installed: $DEST"
"$DEST" --version
