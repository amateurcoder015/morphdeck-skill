#!/usr/bin/env bash
# Install morphdeck as a Claude Code skill, plus its font and Python packages.
#   ./install.sh          copy the skill into ~/.claude/skills/morphdeck
#   ./install.sh --link   symlink instead (edits in this repo apply immediately)
set -euo pipefail

REPO="$(cd "$(dirname "$0")" && pwd)"
DEST="$HOME/.claude/skills/morphdeck"
PY="${PYTHON:-python3}"

echo "→ installing skill to $DEST"
mkdir -p "$HOME/.claude/skills"
if [ -e "$DEST" ] || [ -L "$DEST" ]; then
  echo "  replacing existing $DEST"
  rm -rf "$DEST"
fi
if [ "${1:-}" = "--link" ]; then
  ln -s "$REPO/morphdeck" "$DEST"
else
  cp -R "$REPO/morphdeck" "$DEST"
fi

echo "→ checking Python packages"
pip_install() {
  "$PY" -m pip install --user --quiet "$@" || echo "  ! pip could not install $* (install it manually: $PY -m pip install --user $*)"
}
"$PY" -c "import pptx, PIL, certifi" 2>/dev/null || pip_install python-pptx pillow certifi
if [ "$(uname)" = "Darwin" ]; then
  "$PY" -c "import pymupdf" 2>/dev/null || pip_install pymupdf
fi

echo "→ installing the Unbounded font"
"$PY" "$DEST/scripts/fonts.py"

CFG="$HOME/.config/morphdeck/.env"
if [ ! -f "$CFG" ]; then
  mkdir -p "$(dirname "$CFG")"
  cat > "$CFG" <<'EOF'
# Optional: free AI images via Cloudflare Workers AI (see README → AI images)
# CF_ACCOUNT_ID=
# CF_API_TOKEN=
# Optional: Pexels stock photos instead of Openverse
# PEXELS_API_KEY=
EOF
  chmod 600 "$CFG"
  echo "→ created $CFG (add Cloudflare keys there for AI images)"
fi

echo
echo "✓ morphdeck installed. Restart Claude Code, then try:  /morphdeck Black holes"
