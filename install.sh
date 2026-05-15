#!/usr/bin/env bash
# captioner install — clone-and-symlink into the Claude Code skills directory.
#
# Usage:
#   git clone https://github.com/jbenhart44/captioner.git
#   cd captioner
#   bash install.sh
#
# Then restart Claude Code and run:  /captioner <path-to-deck-or-folder>
set -euo pipefail

SKILLS_DIR="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET="$SKILLS_DIR/captioner"

echo "captioner installer"
echo "  source : $REPO_DIR"
echo "  target : $TARGET"

mkdir -p "$SKILLS_DIR"

if [ -e "$TARGET" ] || [ -L "$TARGET" ]; then
  echo "  note   : $TARGET already exists — removing old link/dir"
  rm -rf "$TARGET"
fi

ln -s "$REPO_DIR" "$TARGET"
echo "  linked : $TARGET -> $REPO_DIR"

echo
echo "Checking Python dependencies..."
if python3 -c "import pptx" 2>/dev/null; then
  echo "  python-pptx: OK"
else
  echo "  python-pptx: MISSING — run: pip install -r \"$REPO_DIR/requirements.txt\""
fi

echo
echo "Done. Restart Claude Code, then run:  /captioner <path-to-.pptx-or-folder>"
